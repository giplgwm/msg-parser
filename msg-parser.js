/**
 * msg-parser.js — Parse Outlook .msg files entirely in the browser (or Node).
 *
 * Zero dependencies, single ES module. Implements:
 *   - CFB / Compound File Binary reader (MS-CFB) — the container format
 *   - MSG property extraction (MS-OXMSG): fixed & variable-length properties,
 *     multi-valued properties, named properties (MS-OXPROPS)
 *   - Recipients, attachments, recursively embedded messages
 *   - Compressed RTF body decompression (MS-OXRTFCP, LZFu)
 *
 * Quick start:
 *   import { parseMsg } from './msg-parser.js';
 *   const msg = parseMsg(await file.arrayBuffer());
 *   console.log(msg.subject, msg.from, msg.recipients, msg.attachments);
 *
 * Lower-level building blocks are exported too: CompoundFile, decompressRTF,
 * PropertyTags, PropertyTypes.
 */

// ---------------------------------------------------------------------------
// Small utilities
// ---------------------------------------------------------------------------

const FREESECT = 0xffffffff;
const ENDOFCHAIN = 0xfffffffe;
const NOSTREAM = 0xffffffff;

function toUint8(data) {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) {
    return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  }
  throw new TypeError('parseMsg: expected ArrayBuffer or Uint8Array');
}

function dataView(u8) {
  return new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
}

const UTF16 = new TextDecoder('utf-16le');

function decodeUtf16(u8) {
  // Strip a trailing NUL terminator if present.
  let len = u8.byteLength & ~1;
  if (len >= 2 && u8[len - 1] === 0 && u8[len - 2] === 0) len -= 2;
  return UTF16.decode(u8.subarray(0, len));
}

function latin1(u8) {
  let s = '';
  for (let i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]);
  return s;
}

/** Map a Windows codepage id to a TextDecoder label. */
const CODEPAGES = {
  437: 'ibm437', 850: 'ibm850', 866: 'ibm866',
  874: 'windows-874', 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5',
  1200: 'utf-16le', 1201: 'utf-16be',
  1250: 'windows-1250', 1251: 'windows-1251', 1252: 'windows-1252',
  1253: 'windows-1253', 1254: 'windows-1254', 1255: 'windows-1255',
  1256: 'windows-1256', 1257: 'windows-1257', 1258: 'windows-1258',
  20127: 'windows-1252' /* us-ascii, superset ok */, 20866: 'koi8-r', 21866: 'koi8-u',
  28591: 'iso-8859-1', 28592: 'iso-8859-2', 28595: 'iso-8859-5',
  28597: 'iso-8859-7', 28599: 'iso-8859-9', 28605: 'iso-8859-15',
  50220: 'iso-2022-jp', 51932: 'euc-jp', 54936: 'gb18030', 65001: 'utf-8',
};

function decodeWithCodepage(u8, codepage) {
  // Strip single trailing NUL common in ANSI streams.
  let end = u8.byteLength;
  if (end > 0 && u8[end - 1] === 0) end -= 1;
  const body = u8.subarray(0, end);
  const label = CODEPAGES[codepage] || 'windows-1252';
  try {
    return new TextDecoder(label).decode(body);
  } catch {
    try {
      return new TextDecoder('windows-1252').decode(body);
    } catch {
      return latin1(body);
    }
  }
}

function filetimeToDate(lo, hi) {
  // FILETIME: 100ns intervals since 1601-01-01 UTC.
  const ms = (hi * 4294967296 + lo) / 10000 - 11644473600000;
  return new Date(ms);
}

function guidToString(u8, off = 0) {
  const dv = dataView(u8);
  const hex = (n, w) => n.toString(16).toUpperCase().padStart(w, '0');
  let s = hex(dv.getUint32(off, true), 8) + '-' +
    hex(dv.getUint16(off + 4, true), 4) + '-' +
    hex(dv.getUint16(off + 6, true), 4) + '-';
  for (let i = 8; i < 10; i++) s += hex(u8[off + i], 2);
  s += '-';
  for (let i = 10; i < 16; i++) s += hex(u8[off + i], 2);
  return s;
}

// ---------------------------------------------------------------------------
// CFB — Compound File Binary reader (MS-CFB)
// ---------------------------------------------------------------------------

/**
 * A directory entry inside a compound file.
 * @typedef {Object} CFBEntry
 * @property {number} id            Stream/storage id (index in directory)
 * @property {string} name
 * @property {number} type          0=unknown 1=storage 2=stream 5=root
 * @property {number} startSector
 * @property {number} size
 * @property {CFBEntry[]} children  Populated for storages and root
 * @property {Map<string,CFBEntry>} childMap
 */

export class CompoundFile {
  /**
   * @param {ArrayBuffer|Uint8Array} data
   * @returns {CompoundFile}
   */
  static parse(data) {
    return new CompoundFile(toUint8(data));
  }

  /** @param {Uint8Array} u8 */
  constructor(u8) {
    this.u8 = u8;
    this.dv = dataView(u8);
    this.#parseHeader();
    this.#loadFAT();
    this.#loadDirectory();
    this.#loadMiniFAT();
    this.#buildTree();
  }

  #parseHeader() {
    const dv = this.dv;
    if (this.u8.byteLength < 512 ||
        dv.getUint32(0, true) !== 0xe011cfd0 || dv.getUint32(4, true) !== 0xe11ab1a1) {
      throw new Error('Not a valid MSG/compound file (bad CFB signature)');
    }
    this.majorVersion = dv.getUint16(26, true);
    this.sectorShift = dv.getUint16(30, true);
    this.sectorSize = 1 << this.sectorShift;           // 512 (v3) or 4096 (v4)
    this.miniSectorSize = 1 << dv.getUint16(32, true); // 64
    this.numFatSectors = dv.getUint32(44, true);
    this.firstDirSector = dv.getUint32(48, true);
    this.miniStreamCutoff = dv.getUint32(56, true);    // 4096
    this.firstMiniFatSector = dv.getUint32(60, true);
    this.numMiniFatSectors = dv.getUint32(64, true);
    this.firstDifatSector = dv.getUint32(68, true);
    this.numDifatSectors = dv.getUint32(72, true);
    this.maxSector = Math.ceil(this.u8.byteLength / this.sectorSize) + 1;
  }

  #sectorOffset(sect) {
    return (sect + 1) * this.sectorSize;
  }

  /** Read one sector as Uint8Array (may be short at EOF for tolerant parsing). */
  #sector(sect) {
    const off = this.#sectorOffset(sect);
    return this.u8.subarray(off, Math.min(off + this.sectorSize, this.u8.byteLength));
  }

  #loadFAT() {
    const dv = this.dv;
    const entriesPerSector = this.sectorSize / 4;
    // DIFAT: 109 entries in header + optional DIFAT sector chain.
    const difat = [];
    for (let i = 0; i < 109; i++) difat.push(dv.getUint32(76 + i * 4, true));
    let ds = this.firstDifatSector;
    let guard = this.numDifatSectors + 4;
    while (ds !== ENDOFCHAIN && ds !== FREESECT && guard-- > 0) {
      const off = this.#sectorOffset(ds);
      for (let i = 0; i < entriesPerSector - 1; i++) {
        difat.push(dv.getUint32(off + i * 4, true));
      }
      ds = dv.getUint32(off + (entriesPerSector - 1) * 4, true);
    }
    // FAT: concatenation of all FAT sectors listed in the DIFAT.
    const fat = new Uint32Array(this.numFatSectors * entriesPerSector || entriesPerSector);
    let n = 0;
    for (const fs of difat) {
      if (fs === FREESECT || fs === ENDOFCHAIN) continue;
      const off = this.#sectorOffset(fs);
      if (off + this.sectorSize > this.u8.byteLength) continue;
      for (let i = 0; i < entriesPerSector; i++) {
        if (n < fat.length) fat[n++] = dv.getUint32(off + i * 4, true);
      }
    }
    this.fat = fat;
  }

  /**
   * Follow a FAT chain and concatenate sectors, truncated to `size`.
   * @param {number} start @param {number} size
   */
  #readChain(start, size) {
    const out = new Uint8Array(size);
    let sect = start, written = 0, guard = this.maxSector + 16;
    while (sect !== ENDOFCHAIN && sect !== FREESECT && written < size && guard-- > 0) {
      const chunk = this.#sector(sect);
      const take = Math.min(chunk.byteLength, size - written);
      out.set(chunk.subarray(0, take), written);
      written += take;
      sect = sect < this.fat.length ? this.fat[sect] : ENDOFCHAIN;
    }
    return out;
  }

  #loadDirectory() {
    // Directory stream: follow FAT chain from firstDirSector; unknown size, so
    // read whole sectors until end of chain.
    const sectors = [];
    let sect = this.firstDirSector, guard = this.maxSector + 16;
    while (sect !== ENDOFCHAIN && sect !== FREESECT && guard-- > 0) {
      sectors.push(this.#sector(sect));
      sect = sect < this.fat.length ? this.fat[sect] : ENDOFCHAIN;
    }
    /** @type {CFBEntry[]} */
    const dirs = [];
    let id = 0;
    for (const sec of sectors) {
      const dv = dataView(sec);
      for (let o = 0; o + 128 <= sec.byteLength; o += 128) {
        const nameLen = dv.getUint16(o + 64, true);
        const type = sec[o + 66];
        const name = nameLen >= 2
          ? UTF16.decode(sec.subarray(o, o + Math.min(nameLen - 2, 64)))
          : '';
        dirs.push({
          id: id++,
          name,
          type,
          left: dv.getUint32(o + 68, true),
          right: dv.getUint32(o + 72, true),
          child: dv.getUint32(o + 76, true),
          startSector: dv.getUint32(o + 116, true),
          // 64-bit size; high dword is 0 for v3 files.
          size: dv.getUint32(o + 120, true) + dv.getUint32(o + 124, true) * 4294967296,
          children: [],
          childMap: new Map(),
        });
      }
    }
    this.dirs = dirs;
    this.root = dirs.find((d) => d.type === 5) || dirs[0];
    if (!this.root) throw new Error('Compound file has no root entry');
  }

  #loadMiniFAT() {
    const bytes = this.#readChain(
      this.firstMiniFatSector,
      this.numMiniFatSectors * this.sectorSize,
    );
    const dv = dataView(bytes);
    const miniFat = new Uint32Array(bytes.byteLength / 4);
    for (let i = 0; i < miniFat.length; i++) miniFat[i] = dv.getUint32(i * 4, true);
    this.miniFat = miniFat;
    // The mini stream lives in the root entry's regular FAT chain.
    this.miniStream = this.root.size > 0
      ? this.#readChain(this.root.startSector, this.root.size)
      : new Uint8Array(0);
  }

  #buildTree() {
    const dirs = this.dirs;
    const collectChildren = (storage) => {
      const result = [];
      const seen = new Set();
      const stack = [storage.child];
      while (stack.length) {
        const idx = stack.pop();
        if (idx === NOSTREAM || idx >= dirs.length || seen.has(idx)) continue;
        seen.add(idx);
        const e = dirs[idx];
        result.push(e);
        stack.push(e.left, e.right);
      }
      return result;
    };
    for (const d of dirs) {
      if (d.type === 1 || d.type === 5) {
        d.children = collectChildren(d);
        for (const c of d.children) d.childMap.set(c.name, c);
      }
    }
  }

  /**
   * Read a stream entry's full contents.
   * @param {CFBEntry} entry
   * @returns {Uint8Array}
   */
  readStream(entry) {
    if (entry.type !== 2) throw new Error(`"${entry.name}" is not a stream`);
    if (entry.size >= this.miniStreamCutoff) {
      return this.#readChain(entry.startSector, entry.size);
    }
    // Mini stream: 64-byte sectors within this.miniStream, chained via miniFAT.
    const out = new Uint8Array(entry.size);
    let sect = entry.startSector, written = 0;
    let guard = (this.miniStream.byteLength / this.miniSectorSize) + 16;
    while (sect !== ENDOFCHAIN && sect !== FREESECT && written < entry.size && guard-- > 0) {
      const off = sect * this.miniSectorSize;
      const take = Math.min(this.miniSectorSize, entry.size - written,
        Math.max(0, this.miniStream.byteLength - off));
      out.set(this.miniStream.subarray(off, off + take), written);
      written += take;
      sect = sect < this.miniFat.length ? this.miniFat[sect] : ENDOFCHAIN;
    }
    return out;
  }

  /**
   * Look up an entry by '/'-separated path from the root, e.g.
   * "__attach_version1.0_#00000000/__substg1.0_37010102".
   * @param {string} path
   * @returns {CFBEntry|null}
   */
  getEntry(path) {
    let cur = this.root;
    for (const part of path.split('/')) {
      if (!part) continue;
      cur = cur.childMap ? cur.childMap.get(part) : undefined;
      if (!cur) return null;
    }
    return cur;
  }
}

// ---------------------------------------------------------------------------
// Compressed RTF (MS-OXRTFCP)
// ---------------------------------------------------------------------------

const LZFU_DICT_INIT =
  '{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman ' +
  '\\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New Roman' +
  'Courier{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20' +
  '\\b\\i\\u\\tab\\tx';

/**
 * Decompress a PidTagRtfCompressed (0x1009) value.
 * Handles both compressed ("LZFu") and stored-uncompressed ("MELA") payloads.
 * @param {Uint8Array} data
 * @returns {Uint8Array} raw RTF bytes
 */
export function decompressRTF(data) {
  data = toUint8(data);
  if (data.byteLength < 16) return new Uint8Array(0);
  const dv = dataView(data);
  const compSize = dv.getUint32(0, true); // size of everything after this field
  const rawSize = dv.getUint32(4, true);
  const compType = dv.getUint32(8, true);
  const end = Math.min(4 + compSize, data.byteLength);

  if (compType === 0x414c454d) { // 'MELA' — uncompressed
    return data.slice(16, Math.min(16 + rawSize, data.byteLength));
  }
  if (compType !== 0x75465a4c) { // 'LZFu'
    throw new Error('Unknown RTF compression type 0x' + compType.toString(16));
  }

  const dict = new Uint8Array(4096);
  for (let i = 0; i < LZFU_DICT_INIT.length; i++) dict[i] = LZFU_DICT_INIT.charCodeAt(i);
  let wp = LZFU_DICT_INIT.length; // 207

  const out = new Uint8Array(rawSize);
  let op = 0;
  let ip = 16;
  while (ip < end && op < rawSize) {
    const control = data[ip++];
    for (let bit = 0; bit < 8 && op < rawSize; bit++) {
      if ((control >> bit) & 1) {
        if (ip + 1 >= end + 1) return out.subarray(0, op);
        const b1 = data[ip++], b2 = data[ip++];
        const offset = (b1 << 4) | (b2 >> 4);
        const length = (b2 & 0x0f) + 2;
        if (offset === (wp & 0x0fff)) return out.subarray(0, op); // end marker
        for (let i = 0; i < length && op < rawSize; i++) {
          const c = dict[(offset + i) & 0x0fff];
          out[op++] = c;
          dict[wp & 0x0fff] = c;
          wp++;
        }
      } else {
        if (ip >= end) return out.subarray(0, op);
        const c = data[ip++];
        out[op++] = c;
        dict[wp & 0x0fff] = c;
        wp++;
      }
    }
  }
  return out.subarray(0, op);
}

// ---------------------------------------------------------------------------
// MAPI property machinery (MS-OXMSG / MS-OXPROPS)
// ---------------------------------------------------------------------------

/** Property type ids. */
export const PropertyTypes = {
  INT16: 0x0002, INT32: 0x0003, FLOAT: 0x0004, DOUBLE: 0x0005,
  CURRENCY: 0x0006, APPTIME: 0x0007, ERROR: 0x000a, BOOLEAN: 0x000b,
  OBJECT: 0x000d, INT64: 0x0014, STRING8: 0x001e, UNICODE: 0x001f,
  SYSTIME: 0x0040, CLSID: 0x0048, BINARY: 0x0102,
  MV_FLAG: 0x1000,
};

/** Friendly names for common property ids (not exhaustive — for display/debug). */
export const PropertyTags = {
  0x001a: 'messageClass',        0x0037: 'subject',
  0x003d: 'subjectPrefix',       0x0e1d: 'normalizedSubject',
  0x0070: 'conversationTopic',   0x0071: 'conversationIndex',
  0x1000: 'body',                0x1013: 'bodyHtml',
  0x1009: 'rtfCompressed',       0x007d: 'transportMessageHeaders',
  0x0039: 'clientSubmitTime',    0x0e06: 'messageDeliveryTime',
  0x3007: 'creationTime',        0x3008: 'lastModificationTime',
  0x0c1a: 'senderName',          0x0c1f: 'senderEmail',
  0x0c1e: 'senderAddressType',   0x5d01: 'senderSmtpAddress',
  0x0042: 'sentRepresentingName', 0x0065: 'sentRepresentingEmail',
  0x0064: 'sentRepresentingAddressType', 0x5d02: 'sentRepresentingSmtpAddress',
  0x0e04: 'displayTo',           0x0e03: 'displayCc',
  0x0e02: 'displayBcc',          0x0e07: 'messageFlags',
  0x0017: 'importance',          0x0036: 'sensitivity',
  0x1035: 'internetMessageId',   0x1039: 'internetReferences',
  0x1042: 'inReplyToId',         0x3fde: 'internetCodepage',
  0x3ffd: 'messageCodepage',     0x0e08: 'messageSize',
  // Recipient
  0x3001: 'displayName',         0x3002: 'addressType',
  0x3003: 'emailAddress',        0x39fe: 'smtpAddress',
  0x0c15: 'recipientType',       0x5ff6: 'recipientDisplayName',
  // Attachment
  0x3701: 'attachData',          0x3703: 'attachExtension',
  0x3704: 'attachFilename',      0x3707: 'attachLongFilename',
  0x3705: 'attachMethod',        0x370e: 'attachMimeTag',
  0x3712: 'attachContentId',     0x370b: 'renderingPosition',
  0x7ffe: 'attachmentHidden',    0x3714: 'attachFlags',
  0x3702: 'attachEncoding',      0x371d: 'attachPayloadClass',
};

const RE_SUBSTG = /^__substg1\.0_([0-9A-F]{4})([0-9A-F]{4})(?:-([0-9A-F]{8}))?$/i;
const RE_RECIP = /^__recip_version1\.0_#([0-9A-F]{8})$/i;
const RE_ATTACH = /^__attach_version1\.0_#([0-9A-F]{8})$/i;
const NAMEID_STORAGE = '__nameid_version1.0';
const PROPS_STREAM = '__properties_version1.0';
const EMBEDDED_MSG_STREAM = '__substg1.0_3701000D';

const PS_MAPI = '00020328-0000-0000-C000-000000000046';
const PS_PUBLIC_STRINGS = '00020329-0000-0000-C000-000000000046';

/**
 * One parsed property.
 * @typedef {Object} MsgProperty
 * @property {number} id       16-bit property id (e.g. 0x0037)
 * @property {number} type     16-bit property type (e.g. 0x001F)
 * @property {number} tag      Full 32-bit tag (id << 16 | type)
 * @property {string|null} name   Friendly or named-property name if known
 * @property {*} value
 * @property {{guid:string, name?:string, lid?:number}|null} named
 */

/** Decode a fixed-size property value from an 8-byte slot. */
function decodeFixed(type, dv, off) {
  switch (type) {
    case PropertyTypes.INT16: return dv.getInt16(off, true);
    case PropertyTypes.INT32: return dv.getInt32(off, true);
    case PropertyTypes.FLOAT: return dv.getFloat32(off, true);
    case PropertyTypes.DOUBLE: return dv.getFloat64(off, true);
    case PropertyTypes.CURRENCY: return dv.getBigInt64
      ? Number(dv.getBigInt64(off, true)) / 10000
      : (dv.getInt32(off, true) + dv.getInt32(off + 4, true) * 4294967296) / 10000;
    case PropertyTypes.APPTIME: {
      // Days since 1899-12-30.
      const days = dv.getFloat64(off, true);
      return new Date(Date.UTC(1899, 11, 30) + days * 86400000);
    }
    case PropertyTypes.BOOLEAN: return dv.getUint8(off) !== 0;
    case PropertyTypes.INT64:
      return Number(dv.getBigInt64(off, true));
    case PropertyTypes.SYSTIME:
      return filetimeToDate(dv.getUint32(off, true), dv.getUint32(off + 4, true));
    case PropertyTypes.ERROR: return dv.getUint32(off, true);
    default: return null;
  }
}

/** Decode a variable-length property value from its own stream. */
function decodeVariable(type, bytes, codepage) {
  switch (type) {
    case PropertyTypes.UNICODE: return decodeUtf16(bytes);
    case PropertyTypes.STRING8: return decodeWithCodepage(bytes, codepage);
    case PropertyTypes.BINARY: return bytes;
    case PropertyTypes.CLSID: return bytes.byteLength >= 16 ? guidToString(bytes) : null;
    default: return bytes;
  }
}

const FIXED_SIZES = {
  0x0002: 2, 0x0003: 4, 0x0004: 4, 0x0005: 8, 0x0006: 8, 0x0007: 8,
  0x000a: 4, 0x000b: 2, 0x0014: 8, 0x0040: 8, 0x0048: 16,
};

/** Decode a multi-valued fixed-size property stored as a flat array stream. */
function decodeMultiFixed(baseType, bytes) {
  const size = FIXED_SIZES[baseType];
  if (!size) return bytes;
  const dv = dataView(bytes);
  const out = [];
  for (let o = 0; o + size <= bytes.byteLength; o += size) {
    out.push(baseType === 0x000b ? dv.getUint16(o, true) !== 0 : decodeFixed(baseType, dv, o));
  }
  return out;
}

/** Parse the __nameid_version1.0 storage into propIndex → named-property info. */
function parseNamedProperties(cfb, nameidStorage) {
  const map = new Map(); // property id (>= 0x8000) → {guid, name?, lid?}
  if (!nameidStorage) return map;
  const read = (name) => {
    const e = nameidStorage.childMap.get(name);
    return e && e.type === 2 ? cfb.readStream(e) : new Uint8Array(0);
  };
  const guidBytes = read('__substg1.0_00020102');
  const entryBytes = read('__substg1.0_00030102');
  const stringBytes = read('__substg1.0_00040102');

  const guids = [];
  for (let o = 0; o + 16 <= guidBytes.byteLength; o += 16) {
    guids.push(guidToString(guidBytes, o));
  }
  const edv = dataView(entryBytes);
  const sdv = dataView(stringBytes);
  for (let o = 0; o + 8 <= entryBytes.byteLength; o += 8) {
    const idOrOffset = edv.getUint32(o, true);
    const indexKind = edv.getUint16(o + 4, true);
    const propIndex = edv.getUint16(o + 6, true);
    const kind = indexKind & 1;
    const guidIndex = indexKind >> 1;
    const guid = guidIndex === 1 ? PS_MAPI
      : guidIndex === 2 ? PS_PUBLIC_STRINGS
      : guids[guidIndex - 3] || null;
    const info = { guid };
    if (kind === 1) {
      if (idOrOffset + 4 <= stringBytes.byteLength) {
        const len = sdv.getUint32(idOrOffset, true);
        const s = idOrOffset + 4;
        info.name = UTF16.decode(
          stringBytes.subarray(s, Math.min(s + len, stringBytes.byteLength)));
      }
    } else {
      info.lid = idOrOffset;
    }
    map.set(0x8000 + propIndex, info);
  }
  return map;
}

/**
 * Parse all properties of one storage (message, recipient, or attachment).
 * @returns {Map<number, MsgProperty>} keyed by 16-bit property id
 */
function parseProperties(cfb, storage, propsHeaderSize, namedProps) {
  /** @type {Map<number, MsgProperty>} */
  const props = new Map();
  const put = (id, type, value) => {
    const existing = props.get(id);
    // Prefer unicode strings over ansi duplicates.
    if (existing && existing.type === PropertyTypes.UNICODE && type === PropertyTypes.STRING8) return;
    const named = id >= 0x8000 ? (namedProps.get(id) || null) : null;
    props.set(id, {
      id, type,
      tag: ((id << 16) | type) >>> 0,
      name: named ? (named.name || (named.lid != null ? 'lid_0x' + named.lid.toString(16) : null))
                  : (PropertyTags[id] || null),
      named,
      value,
    });
  };

  // Pass 1: raw variable-length streams (also discover the codepage).
  const varStreams = [];
  const mvParts = new Map(); // "iiiitttt" → array of {index, bytes}
  let codepage = 1252;
  for (const child of storage.children) {
    const m = RE_SUBSTG.exec(child.name);
    if (!m || child.type !== 2) continue;
    const id = parseInt(m[1], 16);
    const type = parseInt(m[2], 16);
    const bytes = cfb.readStream(child);
    if (m[3] !== undefined) {
      const key = m[1] + m[2];
      if (!mvParts.has(key)) mvParts.set(key, []);
      mvParts.get(key).push({ index: parseInt(m[3], 16), id, type, bytes });
    } else {
      varStreams.push({ id, type, bytes });
    }
  }

  // The fixed-size properties stream also tells us codepages; parse it first.
  const propsEntry = storage.childMap.get(PROPS_STREAM);
  const fixedEntries = [];
  if (propsEntry && propsEntry.type === 2) {
    const bytes = cfb.readStream(propsEntry);
    const dv = dataView(bytes);
    for (let o = propsHeaderSize; o + 16 <= bytes.byteLength; o += 16) {
      const type = dv.getUint16(o, true);
      const id = dv.getUint16(o + 2, true);
      fixedEntries.push({ id, type, dv, off: o + 8 });
      if ((id === 0x3ffd || id === 0x3fde) && type === PropertyTypes.INT32) {
        const cp = dv.getInt32(o + 8, true);
        if (id === 0x3ffd || codepage === 1252) codepage = cp;
      }
    }
  }

  // Pass 2: decode everything.
  for (const { id, type, dv, off } of fixedEntries) {
    if (type in FIXED_SIZES || type === PropertyTypes.BOOLEAN) {
      put(id, type, decodeFixed(type, dv, off));
    }
    // Variable-length entries here only carry sizes; actual data is in substg
    // streams handled below.
  }
  for (const { id, type, bytes } of varStreams) {
    if (type & PropertyTypes.MV_FLAG) {
      const baseType = type & 0x0fff;
      if (FIXED_SIZES[baseType]) put(id, type, decodeMultiFixed(baseType, bytes));
      // Variable-length MV base streams only hold lengths — values arrive via
      // the indexed "-XXXXXXXX" streams collected in mvParts.
    } else {
      put(id, type, decodeVariable(type, bytes, codepage));
    }
  }
  for (const parts of mvParts.values()) {
    parts.sort((a, b) => a.index - b.index);
    const { id, type } = parts[0];
    const baseType = type & 0x0fff;
    put(id, type, parts.map((p) => decodeVariable(baseType, p.bytes, codepage)));
  }
  return props;
}

// ---------------------------------------------------------------------------
// High-level message model
// ---------------------------------------------------------------------------

const RECIPIENT_TYPES = { 1: 'to', 2: 'cc', 3: 'bcc' };
const ATTACH_METHODS = {
  0: 'none', 1: 'byValue', 2: 'byReference', 3: 'byReferenceResolve',
  4: 'byReferenceOnly', 5: 'embeddedMessage', 6: 'ole',
};

/** @typedef {Object} MsgRecipient
 *  @property {string|null} name
 *  @property {string|null} email
 *  @property {'to'|'cc'|'bcc'|null} type
 *  @property {Map<number, MsgProperty>} properties
 */

/** @typedef {Object} MsgAttachment
 *  @property {string|null} fileName      Best available file name
 *  @property {string|null} extension
 *  @property {string|null} mimeType
 *  @property {string|null} contentId     For resolving cid: references in HTML
 *  @property {boolean} hidden
 *  @property {string} method             'byValue' | 'embeddedMessage' | ...
 *  @property {Uint8Array|null} content   Raw bytes (byValue attachments)
 *  @property {MsgMessage|null} innerMsg  Parsed embedded message, if any
 *  @property {Map<number, MsgProperty>} properties
 */

export class MsgMessage {
  /**
   * @param {CompoundFile} cfb
   * @param {CFBEntry} storage
   * @param {number} depth  0 = top-level message
   * @param {Map<number, object>|null} namedProps
   */
  constructor(cfb, storage, depth = 0, namedProps = null) {
    // Named-property mapping only exists at the top level and applies to the
    // whole file, embedded messages included.
    if (!namedProps) {
      namedProps = parseNamedProperties(cfb, storage.childMap.get(NAMEID_STORAGE));
    }
    const headerSize = depth === 0 ? 32 : 24;

    /** @type {Map<number, MsgProperty>} All raw properties, keyed by property id. */
    this.properties = parseProperties(cfb, storage, headerSize, namedProps);
    /** Named-property definitions found in this file (id → {guid, name?, lid?}). */
    this.namedProperties = namedProps;

    /** @type {MsgRecipient[]} */
    this.recipients = [];
    /** @type {MsgAttachment[]} */
    this.attachments = [];

    const recipStorages = [];
    const attachStorages = [];
    for (const child of storage.children) {
      if (child.type !== 1) continue;
      let m;
      if ((m = RE_RECIP.exec(child.name))) recipStorages.push({ i: parseInt(m[1], 16), child });
      else if ((m = RE_ATTACH.exec(child.name))) attachStorages.push({ i: parseInt(m[1], 16), child });
    }
    recipStorages.sort((a, b) => a.i - b.i);
    attachStorages.sort((a, b) => a.i - b.i);

    for (const { child } of recipStorages) {
      const p = parseProperties(cfb, child, 8, namedProps);
      const g = (id) => (p.get(id) ? p.get(id).value : null);
      this.recipients.push({
        name: g(0x3001),
        email: g(0x39fe) || g(0x3003),
        addressType: g(0x3002),
        type: RECIPIENT_TYPES[g(0x0c15)] || null,
        properties: p,
      });
    }

    for (const { child } of attachStorages) {
      const p = parseProperties(cfb, child, 8, namedProps);
      const g = (id) => (p.get(id) ? p.get(id).value : null);
      const method = ATTACH_METHODS[g(0x3705) ?? 1] || 'byValue';
      let content = null;
      let innerMsg = null;
      const dataProp = p.get(0x3701);
      if (dataProp && dataProp.value instanceof Uint8Array) content = dataProp.value;
      const embedded = child.childMap.get(EMBEDDED_MSG_STREAM);
      if (embedded && embedded.type === 1) {
        try {
          innerMsg = new MsgMessage(cfb, embedded, depth + 1, namedProps);
        } catch { /* tolerate malformed embedded messages */ }
      }
      const inner = innerMsg;
      this.attachments.push({
        fileName: g(0x3707) || g(0x3704) ||
          (inner && inner.subject ? inner.subject + '.msg' : null),
        extension: g(0x3703),
        mimeType: g(0x370e),
        contentId: g(0x3712),
        hidden: g(0x7ffe) === true,
        method,
        content,
        innerMsg,
        properties: p,
      });
    }
  }

  /**
   * Get a property value by numeric id (e.g. 0x0037), friendly name
   * (e.g. "subject"), or named-property name (e.g. "x-custom").
   * @param {number|string} key
   * @returns {*} value, or undefined
   */
  getProperty(key) {
    if (typeof key === 'number') {
      const p = this.properties.get(key);
      return p ? p.value : undefined;
    }
    const lower = String(key).toLowerCase();
    for (const p of this.properties.values()) {
      if (p.name && p.name.toLowerCase() === lower) return p.value;
      if (p.named && p.named.name && p.named.name.toLowerCase() === lower) return p.value;
    }
    return undefined;
  }

  #v(id) {
    const p = this.properties.get(id);
    return p ? p.value : null;
  }

  /** e.g. "IPM.Note" */
  get messageClass() { return this.#v(0x001a); }
  get subject() { return this.#v(0x0037); }
  /** Plain-text body. */
  get body() { return this.#v(0x1000); }
  /** HTML body as a string (decoded from PidTagHtml), or null. */
  get bodyHtml() {
    const p = this.properties.get(0x1013);
    if (!p) return null;
    if (typeof p.value === 'string') return p.value;
    if (p.value instanceof Uint8Array) {
      const cp = this.#v(0x3fde) || this.#v(0x3ffd) || 65001;
      return decodeWithCodepage(p.value, cp);
    }
    return null;
  }
  /** Raw compressed RTF bytes (PidTagRtfCompressed), or null. */
  get rtfCompressed() {
    const v = this.#v(0x1009);
    return v instanceof Uint8Array ? v : null;
  }
  /** Decompressed RTF source text, or null. */
  get rtfBody() {
    const c = this.rtfCompressed;
    if (!c) return null;
    try {
      return decodeWithCodepage(decompressRTF(c), this.#v(0x3ffd) || 1252);
    } catch {
      return null;
    }
  }
  /** Raw SMTP transport headers string, or null. */
  get headers() { return this.#v(0x007d); }
  /** {name, email} of the sender. */
  get from() {
    const email = this.#v(0x5d01) || this.#v(0x5d02) ||
      (String(this.#v(0x0c1e) || '').toUpperCase() === 'SMTP' ? this.#v(0x0c1f) : null) ||
      this.#v(0x0c1f) || this.#v(0x0065);
    return {
      name: this.#v(0x0c1a) || this.#v(0x0042),
      email: email || null,
    };
  }
  get to() { return this.recipients.filter((r) => r.type === 'to'); }
  get cc() { return this.recipients.filter((r) => r.type === 'cc'); }
  get bcc() { return this.recipients.filter((r) => r.type === 'bcc'); }
  /** Date sent (falls back to delivery/creation time). @returns {Date|null} */
  get date() {
    return this.#v(0x0039) || this.#v(0x0e06) || this.#v(0x3007);
  }
  get deliveryDate() { return this.#v(0x0e06); }
  get messageId() { return this.#v(0x1035); }
  get conversationTopic() { return this.#v(0x0070); }
  /** 0 low, 1 normal, 2 high (PidTagImportance). */
  get importance() { return this.#v(0x0017); }

  /** JSON-safe snapshot (attachment bytes replaced with lengths). */
  toJSON() {
    return {
      messageClass: this.messageClass,
      subject: this.subject,
      from: this.from,
      date: this.date,
      recipients: this.recipients.map((r) => ({
        name: r.name, email: r.email, type: r.type,
      })),
      body: this.body,
      bodyHtml: this.bodyHtml,
      headers: this.headers,
      messageId: this.messageId,
      attachments: this.attachments.map((a) => ({
        fileName: a.fileName,
        mimeType: a.mimeType,
        contentId: a.contentId,
        method: a.method,
        hidden: a.hidden,
        size: a.content ? a.content.byteLength : null,
        innerMsg: a.innerMsg ? a.innerMsg.toJSON() : null,
      })),
    };
  }
}

/**
 * Parse an Outlook .msg file.
 * @param {ArrayBuffer|Uint8Array} data  File contents.
 * @returns {MsgMessage}
 */
export function parseMsg(data) {
  const cfb = CompoundFile.parse(data);
  return new MsgMessage(cfb, cfb.root, 0);
}

export default parseMsg;
