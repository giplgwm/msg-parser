# msg-parser

A zero-dependency JavaScript library for parsing Outlook `.msg` files entirely in the browser (Node works too). Single ES module, no build step.

It implements the full stack the format requires: the CFB/Compound File Binary container ([MS-CFB]), MAPI property extraction including named and multi-valued properties ([MS-OXMSG]), recipients, attachments, recursively embedded messages, and compressed-RTF body decompression ([MS-OXRTFCP]).

## Quick start

```html
<script type="module">
  import { parseMsg } from './msg-parser.js';

  const file = /* a File from an <input> or drag-and-drop */;
  const msg = parseMsg(await file.arrayBuffer());

  msg.subject;        // "Quarterly report"
  msg.from;           // { name: "Jane Doe", email: "jane@example.com" }
  msg.recipients;     // [{ name, email, type: 'to'|'cc'|'bcc', properties }]
  msg.to; msg.cc; msg.bcc;
  msg.date;           // Date sent (falls back to delivery/creation time)
  msg.body;           // plain-text body
  msg.bodyHtml;       // HTML body, decoded with the right codepage
  msg.rtfBody;        // decompressed RTF source (LZFu / MELA)
  msg.headers;        // raw SMTP transport headers
  msg.attachments;    // see below
</script>
```

Try `demo.html` for a full drag-and-drop viewer (serve the folder with any static server, e.g. `python -m http.server` — ES modules don't load from `file://`). `test/sample.msg` is a generated file you can drop in.

## API

### `parseMsg(data) → MsgMessage`

`data` is an `ArrayBuffer`, `Uint8Array`, or Node `Buffer`. Throws if the data is not a compound file.

### `MsgMessage`

Convenience getters: `subject`, `body`, `bodyHtml`, `rtfBody`, `rtfCompressed`, `headers`, `from`, `recipients`, `to`, `cc`, `bcc`, `date`, `deliveryDate`, `messageId`, `conversationTopic`, `importance`, `messageClass`, `attachments`.

Attachments look like:

```js
{
  fileName: 'photo.jpg',      // best available name
  extension: '.jpg',
  mimeType: 'image/jpeg',
  contentId: 'img1@...',      // for resolving cid: refs in bodyHtml
  hidden: false,
  method: 'byValue',          // or 'embeddedMessage', 'ole', ...
  content: Uint8Array | null, // raw bytes for byValue attachments
  innerMsg: MsgMessage | null,// parsed .msg attached to a .msg, recursively
  properties: Map,            // all raw attachment properties
}
```

Every MAPI property is also available raw:

```js
msg.properties;               // Map<propertyId, {id, type, tag, name, named, value}>
msg.getProperty(0x0037);      // by numeric id
msg.getProperty('subject');   // by friendly name
msg.getProperty('Keywords');  // by named-property name (resolved via __nameid)
msg.toJSON();                 // JSON-safe snapshot (bytes → lengths)
```

Values are decoded by MAPI type: unicode/ANSI strings (codepage-aware via `PidTagMessageCodepage`/`PidTagInternetCodepage`), `FILETIME` → `Date`, integers, floats, booleans, currency, GUIDs, binary → `Uint8Array`, and multi-valued variants → arrays.

### Lower-level exports

For building your own tooling on top:

```js
import { CompoundFile, decompressRTF, PropertyTags, PropertyTypes } from './msg-parser.js';

const cfb = CompoundFile.parse(buffer);          // any CFB file, not just .msg
cfb.root.children;                               // directory tree
cfb.getEntry('__attach_version1.0_#00000000/__substg1.0_37010102');
cfb.readStream(entry);                           // → Uint8Array

decompressRTF(bytes);                            // PidTagRtfCompressed → RTF bytes
```

## Tests

```
node test/test.js
```

The suite generates a synthetic-but-valid `.msg` with its own independent CFB writer (unicode + ANSI strings, mini-stream and regular-FAT streams, recipients, a binary attachment, an embedded message, named properties, multi-valued properties, and the compressed-RTF test vector from MS-OXRTFCP §3.1) and asserts the parser reads it all back.

## Notes and limitations

- Read-only: parsing, not writing.
- `rtfBody` returns RTF *source*; converting RTF to plain text/HTML is out of scope. (Most modern messages carry `bodyHtml` anyway; when only RTF exists it often wraps the HTML — search for `\fromhtml`.)
- OLE (`method: 'ole'`) attachment payloads are returned as raw bytes without further decoding.
- Corrupt files are handled tolerantly (bounded chain walks, cycle guards) — you get as much data as can be recovered, or a clear error for non-CFB input.
