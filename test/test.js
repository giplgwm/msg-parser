/**
 * Test suite for msg-parser.js.
 *
 * Generates a synthetic-but-valid .msg file from scratch (its own minimal CFB
 * writer, independent of the parser) and asserts the parser extracts
 * everything correctly. Run with:  node test/test.js
 */
import { strict as assert } from 'node:assert';
import { writeFileSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseMsg, decompressRTF, CompoundFile } from '../msg-parser.js';

const __dir = dirname(fileURLToPath(import.meta.url));

// ---------------------------------------------------------------------------
// Minimal CFB writer (v3, 512-byte sectors) — test-only
// ---------------------------------------------------------------------------

const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const NOSTREAM = 0xffffffff;
const FATSECT = 0xfffffffd;
const SEC = 512, MINISEC = 64, CUTOFF = 4096;

/**
 * @param {{name:string, children?:any[], data?:Buffer}} rootSpec
 * @returns {Buffer} full compound file
 */
function writeCFB(rootSpec) {
  // Flatten tree, assign ids, build sibling chains (degenerate right-chains).
  const entries = [];
  function add(spec, type) {
    const e = {
      name: spec.name, type,
      left: NOSTREAM, right: NOSTREAM, child: NOSTREAM,
      data: spec.data || null, start: ENDOFCHAIN, size: 0,
    };
    entries.push(e);
    if (spec.children) {
      let prev = null;
      for (const c of spec.children) {
        const ce = add(c, c.children ? 1 : 2);
        if (prev) prev.right = entries.indexOf(ce);
        else e.child = entries.indexOf(ce);
        prev = ce;
      }
    }
    return e;
  }
  add(rootSpec, 5);

  // Allocate mini sectors (small streams) and count big-stream sectors.
  const miniFat = [];
  const miniChunks = [];
  const bigStreams = [];
  for (const e of entries) {
    if (e.type !== 2 || !e.data) continue;
    e.size = e.data.length;
    if (e.data.length === 0) continue;
    if (e.data.length < CUTOFF) {
      const nsec = Math.ceil(e.data.length / MINISEC);
      e.start = miniFat.length;
      for (let i = 0; i < nsec; i++) {
        miniFat.push(i === nsec - 1 ? ENDOFCHAIN : miniFat.length + 1);
      }
      const padded = Buffer.alloc(nsec * MINISEC);
      e.data.copy(padded);
      miniChunks.push(padded);
    } else {
      bigStreams.push(e);
    }
  }
  const miniStream = Buffer.concat(miniChunks);
  const root = entries[0];
  root.size = miniStream.length;

  // Sector layout: [dir][miniFAT][miniStream][big streams][FAT sectors]
  const dirSectors = Math.ceil((entries.length * 128) / SEC);
  const miniFatSectors = Math.ceil((miniFat.length * 4) / SEC);
  const miniStreamSectors = Math.ceil(miniStream.length / SEC);
  const bigSectorCounts = bigStreams.map((e) => Math.ceil(e.data.length / SEC));
  const bigSectors = bigSectorCounts.reduce((a, b) => a + b, 0);
  let fatSectors = 1;
  for (let i = 0; i < 8; i++) {
    const total = dirSectors + miniFatSectors + miniStreamSectors + bigSectors + fatSectors;
    const need = Math.ceil(total / (SEC / 4));
    if (need === fatSectors) break;
    fatSectors = need;
  }
  const totalSectors = dirSectors + miniFatSectors + miniStreamSectors + bigSectors + fatSectors;

  const dirStart = 0;
  const miniFatStart = dirStart + dirSectors;
  const miniStreamStart = miniFatStart + miniFatSectors;
  const bigStart = miniStreamStart + miniStreamSectors;
  const fatStart = bigStart + bigSectors;

  root.start = miniStream.length > 0 ? miniStreamStart : ENDOFCHAIN;
  {
    let s = bigStart;
    for (let i = 0; i < bigStreams.length; i++) {
      bigStreams[i].start = s;
      s += bigSectorCounts[i];
    }
  }

  // Build the FAT.
  const fat = new Array(fatSectors * (SEC / 4)).fill(FREESECT);
  const chain = (start, count) => {
    for (let i = 0; i < count; i++) fat[start + i] = i === count - 1 ? ENDOFCHAIN : start + i + 1;
  };
  chain(dirStart, dirSectors);
  if (miniFatSectors) chain(miniFatStart, miniFatSectors);
  if (miniStreamSectors) chain(miniStreamStart, miniStreamSectors);
  {
    let s = bigStart;
    for (const n of bigSectorCounts) { chain(s, n); s += n; }
  }
  for (let i = 0; i < fatSectors; i++) fat[fatStart + i] = FATSECT;

  // Serialize.
  const file = Buffer.alloc(SEC + totalSectors * SEC);
  const sec = (n) => SEC + n * SEC;

  // Header
  Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(file, 0);
  file.writeUInt16LE(0x003e, 24);       // minor version
  file.writeUInt16LE(3, 26);            // major version
  file.writeUInt16LE(0xfffe, 28);       // byte order
  file.writeUInt16LE(9, 30);            // sector shift
  file.writeUInt16LE(6, 32);            // mini sector shift
  file.writeUInt32LE(fatSectors, 44);
  file.writeUInt32LE(dirStart, 48);
  file.writeUInt32LE(CUTOFF, 56);
  file.writeUInt32LE(miniFatSectors ? miniFatStart : ENDOFCHAIN, 60);
  file.writeUInt32LE(miniFatSectors, 64);
  file.writeUInt32LE(ENDOFCHAIN, 68);   // first DIFAT sector (none)
  file.writeUInt32LE(0, 72);            // num DIFAT sectors
  for (let i = 0; i < 109; i++) {
    file.writeUInt32LE(i < fatSectors ? fatStart + i : FREESECT, 76 + i * 4);
  }

  // Directory entries
  for (let i = 0; i < entries.length; i++) {
    const e = entries[i];
    const o = sec(dirStart) + i * 128;
    const name = Buffer.from(e.name, 'utf16le');
    name.copy(file, o, 0, Math.min(name.length, 62));
    file.writeUInt16LE(name.length + 2, o + 64); // incl. NUL terminator
    file.writeUInt8(e.type, o + 66);
    file.writeUInt8(1, o + 67); // black
    file.writeUInt32LE(e.left, o + 68);
    file.writeUInt32LE(e.right, o + 72);
    file.writeUInt32LE(e.child, o + 76);
    file.writeUInt32LE(e.type === 2 || e.type === 5 ? (e.start >>> 0) : 0, o + 116);
    file.writeUInt32LE(e.size, o + 120);
  }

  // miniFAT
  for (let i = 0; i < miniFat.length; i++) {
    file.writeUInt32LE(miniFat[i] >>> 0, sec(miniFatStart) + i * 4);
  }
  // mini stream
  miniStream.copy(file, sec(miniStreamStart));
  // big streams
  for (const e of bigStreams) e.data.copy(file, sec(e.start));
  // FAT
  for (let i = 0; i < fat.length; i++) {
    file.writeUInt32LE(fat[i] >>> 0, sec(fatStart) + i * 4);
  }
  return file;
}

// ---------------------------------------------------------------------------
// Synthetic MSG content
// ---------------------------------------------------------------------------

const u16 = (s) => Buffer.from(s, 'utf16le');
const substg = (id, type, data) => ({
  name: `__substg1.0_${id.toString(16).toUpperCase().padStart(4, '0')}` +
        `${type.toString(16).toUpperCase().padStart(4, '0')}`,
  data,
});

function propsStream(headerSize, fixed, { recips = 0, attachs = 0 } = {}) {
  const buf = Buffer.alloc(headerSize + fixed.length * 16);
  if (headerSize >= 24) {
    buf.writeUInt32LE(recips, 8);   // next recipient id
    buf.writeUInt32LE(attachs, 12); // next attachment id
    buf.writeUInt32LE(recips, 16);  // recipient count
    buf.writeUInt32LE(attachs, 20); // attachment count
  }
  fixed.forEach(({ id, type, value }, i) => {
    const o = headerSize + i * 16;
    buf.writeUInt32LE(((id << 16) | type) >>> 0, o);
    buf.writeUInt32LE(0x06, o + 4); // flags: readable|writable
    value.copy(buf, o + 8);
  });
  return { name: '__properties_version1.0', data: buf };
}

const i32 = (n) => { const b = Buffer.alloc(8); b.writeInt32LE(n, 0); return b; };
const boolv = (v) => { const b = Buffer.alloc(8); b.writeUInt8(v ? 1 : 0, 0); return b; };
const filetime = (date) => {
  const b = Buffer.alloc(8);
  b.writeBigUInt64LE((BigInt(date.getTime()) + 11644473600000n) * 10000n, 0);
  return b;
};

// Canonical LZFu test vector from MS-OXRTFCP §3.1:
// decompresses to '{\rtf1\ansi\ansicpg1252\pard hello world}\r\n'
const RTF_COMPRESSED = Buffer.from([
  0x2d, 0x00, 0x00, 0x00, 0x2b, 0x00, 0x00, 0x00,
  0x4c, 0x5a, 0x46, 0x75, 0xf1, 0xc5, 0xc7, 0xa7,
  0x03, 0x00, 0x0a, 0x00, 0x72, 0x63, 0x70, 0x67,
  0x31, 0x32, 0x35, 0x42, 0x32, 0x0a, 0xf3, 0x20,
  0x68, 0x65, 0x6c, 0x09, 0x00, 0x20, 0x62, 0x77,
  0x05, 0xb0, 0x6c, 0x64, 0x7d, 0x0a, 0x80, 0x0f,
  0xa0,
]);
const RTF_EXPECTED = '{\\rtf1\\ansi\\ansicpg1252\\pard hello world}\r\n';

const SENT = new Date(Date.UTC(2020, 0, 2, 3, 4, 5));

// Big attachment (> 4096 bytes) exercises the regular-FAT stream path.
const BIG = Buffer.alloc(8000);
for (let i = 0; i < BIG.length; i++) BIG[i] = i % 251;

// Named property definitions: one string-named prop "TestProp" -> 0x8000.
const NAMED_GUID = Buffer.from([
  0x44, 0x33, 0x22, 0x11, 0x66, 0x55, 0x88, 0x77,
  0x99, 0xaa, 0xbb, 0xcc, 0xdd, 0xee, 0xff, 0x00,
]); // renders as 11223344-5566-7788-99AA-BBCCDDEEFF00
const nameEntry = Buffer.alloc(8);
nameEntry.writeUInt32LE(0, 0);            // string offset 0
nameEntry.writeUInt16LE((3 << 1) | 1, 4); // guidIndex 3 (first custom), kind=string
nameEntry.writeUInt16LE(0, 6);            // propIndex 0 -> id 0x8000
const nameString = Buffer.concat([i32('TestProp'.length * 2).subarray(0, 4), u16('TestProp')]);

const msgSpec = {
  name: 'Root Entry',
  children: [
    propsStream(32, [
      { id: 0x0039, type: 0x0040, value: filetime(SENT) },
      { id: 0x0017, type: 0x0003, value: i32(2) },       // importance: high
      { id: 0x3ffd, type: 0x0003, value: i32(1252) },    // message codepage
    ], { recips: 2, attachs: 2 }),
    substg(0x001a, 0x001f, u16('IPM.Note')),
    substg(0x0037, 0x001f, u16('Test Subject ✓')),
    substg(0x1000, 0x001f, u16('This is the body.\r\nLine two.')),
    substg(0x007d, 0x001f, u16('From: sender@example.com\r\nTo: alice@example.com\r\n')),
    substg(0x0c1a, 0x001f, u16('Sender Person')),
    substg(0x0c1e, 0x001f, u16('SMTP')),
    substg(0x0c1f, 0x001f, u16('sender@example.com')),
    substg(0x1009, 0x0102, RTF_COMPRESSED),
    substg(0x8000, 0x001f, u16('named value')),
    // Multi-valued unicode property 0x0F0F
    substg(0x0f0f, 0x101f, Buffer.alloc(8)), // length stream (ignored by parser)
    { name: '__substg1.0_0F0F101F-00000000', data: u16('one') },
    { name: '__substg1.0_0F0F101F-00000001', data: u16('two') },
    {
      name: '__nameid_version1.0',
      children: [
        { name: '__substg1.0_00020102', data: NAMED_GUID },
        { name: '__substg1.0_00030102', data: nameEntry },
        { name: '__substg1.0_00040102', data: nameString },
      ],
    },
    {
      name: '__recip_version1.0_#00000000',
      children: [
        propsStream(8, [{ id: 0x0c15, type: 0x0003, value: i32(1) }]),
        substg(0x3001, 0x001f, u16('Alice Example')),
        substg(0x39fe, 0x001f, u16('alice@example.com')),
      ],
    },
    {
      name: '__recip_version1.0_#00000001',
      children: [
        propsStream(8, [{ id: 0x0c15, type: 0x0003, value: i32(2) }]),
        substg(0x3001, 0x001f, u16('Bob Example')),
        substg(0x3003, 0x001f, u16('bob@example.com')),
      ],
    },
    {
      name: '__attach_version1.0_#00000000',
      children: [
        propsStream(8, [
          { id: 0x3705, type: 0x0003, value: i32(1) },
          { id: 0x7ffe, type: 0x000b, value: boolv(false) },
        ]),
        substg(0x3707, 0x001f, u16('hello.bin')),
        substg(0x3703, 0x001e, Buffer.from('.bin\0', 'latin1')), // ANSI string
        substg(0x370e, 0x001f, u16('application/octet-stream')),
        substg(0x3701, 0x0102, BIG),
      ],
    },
    {
      name: '__attach_version1.0_#00000001',
      children: [
        propsStream(8, [{ id: 0x3705, type: 0x0003, value: i32(5) }]),
        {
          name: '__substg1.0_3701000D',
          children: [
            propsStream(24, []),
            substg(0x0037, 0x001f, u16('Inner subject')),
            substg(0x1000, 0x001f, u16('Inner body')),
          ],
        },
      ],
    },
  ],
};

// ---------------------------------------------------------------------------
// Run tests
// ---------------------------------------------------------------------------

let passed = 0;
function test(name, fn) {
  fn();
  passed++;
  console.log(`  ✓ ${name}`);
}

const fileBuf = writeCFB(msgSpec);
writeFileSync(join(__dir, 'sample.msg'), fileBuf);
console.log(`Generated test/sample.msg (${fileBuf.length} bytes)\n`);

const msg = parseMsg(fileBuf);

test('CFB layer parses and exposes streams', () => {
  const cfb = CompoundFile.parse(fileBuf);
  const e = cfb.getEntry('__substg1.0_0037001F');
  assert.ok(e);
  assert.equal(e.type, 2);
});

test('subject (unicode, incl. non-ASCII)', () => {
  assert.equal(msg.subject, 'Test Subject ✓');
});

test('body and headers', () => {
  assert.equal(msg.body, 'This is the body.\r\nLine two.');
  assert.match(msg.headers, /^From: sender@example\.com/);
});

test('message class and importance', () => {
  assert.equal(msg.messageClass, 'IPM.Note');
  assert.equal(msg.importance, 2);
});

test('sender', () => {
  assert.deepEqual(msg.from, { name: 'Sender Person', email: 'sender@example.com' });
});

test('sent date (FILETIME)', () => {
  assert.equal(msg.date.getTime(), SENT.getTime());
});

test('recipients with types', () => {
  assert.equal(msg.recipients.length, 2);
  assert.equal(msg.to.length, 1);
  assert.equal(msg.cc.length, 1);
  assert.deepEqual(
    msg.recipients.map((r) => [r.name, r.email, r.type]),
    [['Alice Example', 'alice@example.com', 'to'],
     ['Bob Example', 'bob@example.com', 'cc']],
  );
});

test('binary attachment (> 4096 bytes, regular FAT stream)', () => {
  const a = msg.attachments[0];
  assert.equal(a.fileName, 'hello.bin');
  assert.equal(a.mimeType, 'application/octet-stream');
  assert.equal(a.method, 'byValue');
  assert.equal(a.hidden, false);
  assert.equal(a.content.byteLength, BIG.length);
  assert.deepEqual(Array.from(a.content.subarray(0, 6)), [0, 1, 2, 3, 4, 5]);
  assert.equal(a.content[7999], 7999 % 251);
});

test('ANSI (STRING8) property decoded with codepage', () => {
  assert.equal(msg.attachments[0].extension, '.bin');
});

test('embedded message attachment parsed recursively', () => {
  const a = msg.attachments[1];
  assert.equal(a.method, 'embeddedMessage');
  assert.ok(a.innerMsg);
  assert.equal(a.innerMsg.subject, 'Inner subject');
  assert.equal(a.innerMsg.body, 'Inner body');
  assert.equal(a.fileName, 'Inner subject.msg');
});

test('compressed RTF body (LZFu, spec test vector)', () => {
  assert.equal(msg.rtfBody, RTF_EXPECTED);
});

test('decompressRTF handles MELA (stored uncompressed)', () => {
  const raw = Buffer.from('{\\rtf1 plain}', 'latin1');
  const hdr = Buffer.alloc(16);
  hdr.writeUInt32LE(raw.length + 12, 0);
  hdr.writeUInt32LE(raw.length, 4);
  hdr.writeUInt32LE(0x414c454d, 8); // 'MELA'
  const out = decompressRTF(Buffer.concat([hdr, raw]));
  assert.equal(Buffer.from(out).toString('latin1'), '{\\rtf1 plain}');
});

test('named property resolved via __nameid mapping', () => {
  assert.equal(msg.getProperty('TestProp'), 'named value');
  const p = msg.properties.get(0x8000);
  assert.equal(p.named.name, 'TestProp');
  assert.equal(p.named.guid, '11223344-5566-7788-99AA-BBCCDDEEFF00');
});

test('multi-valued unicode property', () => {
  assert.deepEqual(msg.getProperty(0x0f0f), ['one', 'two']);
});

test('getProperty by id and friendly name', () => {
  assert.equal(msg.getProperty(0x0037), 'Test Subject ✓');
  assert.equal(msg.getProperty('subject'), 'Test Subject ✓');
});

test('toJSON is JSON-serializable and complete', () => {
  const j = JSON.parse(JSON.stringify(msg.toJSON()));
  assert.equal(j.subject, 'Test Subject ✓');
  assert.equal(j.attachments.length, 2);
  assert.equal(j.attachments[0].size, BIG.length);
  assert.equal(j.attachments[1].innerMsg.subject, 'Inner subject');
});

test('rejects non-CFB data', () => {
  assert.throws(() => parseMsg(new Uint8Array(600)), /signature/);
});

console.log(`\nAll ${passed} tests passed.`);
