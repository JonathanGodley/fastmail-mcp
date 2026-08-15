// Minimal raw-JMAP helper for the live probes: session, method calls, blob
// upload/download, and a tiny PNG generator for image fixtures.
// References the FASTMAIL_API_TOKEN env var by NAME only; never prints its value.
import { deflateSync } from 'node:zlib';
import { createHash } from 'node:crypto';

const TOKEN = process.env.FASTMAIL_API_TOKEN;
if (!TOKEN) {
  console.error('FASTMAIL_API_TOKEN is not set in the child environment. Aborting.');
  process.exit(2);
}

const AUTH = { Authorization: `Bearer ${TOKEN}` };

export async function getSession() {
  const res = await fetch('https://api.fastmail.com/jmap/session', { headers: AUTH });
  if (res.status === 401) {
    console.error('AUTH FAILED: 401 from /jmap/session. Stopping.');
    process.exit(3);
  }
  if (!res.ok) throw new Error(`session ${res.status}`);
  const s = await res.json();
  const accountId = s.primaryAccounts['urn:ietf:params:jmap:mail'];
  return { accountId, apiUrl: s.apiUrl, uploadUrl: s.uploadUrl, downloadUrl: s.downloadUrl };
}

export async function jmap(session, methodCalls, extraUsing = []) {
  const res = await fetch(session.apiUrl, {
    method: 'POST',
    headers: { ...AUTH, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail', ...extraUsing],
      methodCalls,
    }),
  });
  if (res.status === 401) {
    console.error('AUTH FAILED: 401 from the JMAP API. Stopping.');
    process.exit(3);
  }
  const text = await res.text();
  let body;
  try { body = JSON.parse(text); } catch { throw new Error(`non-JSON response ${res.status}: ${text.slice(0, 400)}`); }
  if (!res.ok) throw new Error(`JMAP ${res.status}: ${JSON.stringify(body).slice(0, 600)}`);
  return body;
}

export async function upload(session, bytes, type = 'image/png') {
  const url = session.uploadUrl.replace('{accountId}', encodeURIComponent(session.accountId));
  const res = await fetch(url, {
    method: 'POST',
    headers: { ...AUTH, 'Content-Type': type },
    body: bytes,
  });
  if (res.status === 401) { console.error('AUTH FAILED: 401 on upload. Stopping.'); process.exit(3); }
  if (!res.ok) throw new Error(`upload ${res.status}: ${(await res.text()).slice(0, 300)}`);
  return res.json(); // { accountId, blobId, type, size }
}

export async function downloadBlob(session, blobId, name = 'blob', type = 'application/octet-stream') {
  const url = session.downloadUrl
    .replace('{accountId}', encodeURIComponent(session.accountId))
    .replace('{blobId}', encodeURIComponent(blobId))
    .replace('{name}', encodeURIComponent(name))
    .replace('{type}', encodeURIComponent(type));
  const res = await fetch(url, { headers: AUTH });
  if (res.status === 401) { console.error('AUTH FAILED: 401 on download. Stopping.'); process.exit(3); }
  if (!res.ok) throw new Error(`download ${res.status}`);
  return Buffer.from(await res.arrayBuffer());
}

export function sha256(buf) {
  return createHash('sha256').update(buf).digest('hex');
}

// --- tiny valid PNG generator (1x1, given RGB) ---
function crc32(buf) {
  let c, crc = 0xffffffff;
  for (let n = 0; n < buf.length; n++) {
    c = (crc ^ buf[n]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = c ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}
function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
}
export function makePng(r, g, b) {
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(1, 0); ihdr.writeUInt32BE(1, 4);
  ihdr[8] = 8; ihdr[9] = 2; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0; // 8-bit truecolor
  const raw = Buffer.from([0x00, r, g, b]); // filter byte + one RGB pixel
  return Buffer.concat([sig, chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}
