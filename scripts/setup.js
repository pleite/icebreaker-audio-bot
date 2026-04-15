#!/usr/bin/env node
'use strict';

/**
 * scripts/setup.js
 *
 * Interactive setup helper that:
 *  1. Generates placeholder Teams app icons (color.png + outline.png)
 *  2. Patches manifest/manifest.json with the real bot App ID
 *  3. Packages manifest/ into manifest.zip ready for Teams upload
 *
 * Run:  node scripts/setup.js
 */

const fs = require('fs');
const path = require('path');
const readline = require('readline');
const crypto = require('crypto');
const zlib = require('zlib');

const MANIFEST_DIR = path.join(__dirname, '..', 'manifest');
const MANIFEST_FILE = path.join(MANIFEST_DIR, 'manifest.json');
const ZIP_OUT = path.join(MANIFEST_DIR, 'manifest.zip');

// ---------------------------------------------------------------------------
// Minimal PNG generation (no external deps)
// Creates a solid-colour PNG image using raw PNG chunks.
// ---------------------------------------------------------------------------

function crc32(buf) {
  let crc = 0xffffffff;
  for (const byte of buf) {
    crc ^= byte;
    for (let i = 0; i < 8; i++) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function uint32BE(n) {
  const buf = Buffer.alloc(4);
  buf.writeUInt32BE(n);
  return buf;
}

function pngChunk(type, data) {
  const typeBytes = Buffer.from(type, 'ascii');
  const crc = crc32(Buffer.concat([typeBytes, data]));
  return Buffer.concat([uint32BE(data.length), typeBytes, data, uint32BE(crc)]);
}

/**
 * Create a minimal solid-colour PNG.
 * @param {number} width
 * @param {number} height
 * @param {[number,number,number]} rgb  e.g. [0, 120, 212]
 * @returns {Buffer}
 */
function createSolidPng(width, height, rgb) {
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

  // IHDR
  const ihdrData = Buffer.concat([
    uint32BE(width),
    uint32BE(height),
    Buffer.from([8, 2, 0, 0, 0]), // bit depth=8, colourType=2 (RGB), no interlace
  ]);
  const ihdr = pngChunk('IHDR', ihdrData);

  // Build raw pixel data (filter byte 0x00 before each row)
  const rowBytes = width * 3;
  const raw = Buffer.alloc((rowBytes + 1) * height);
  for (let y = 0; y < height; y++) {
    raw[y * (rowBytes + 1)] = 0; // filter type: None
    for (let x = 0; x < width; x++) {
      const offset = y * (rowBytes + 1) + 1 + x * 3;
      raw[offset] = rgb[0];
      raw[offset + 1] = rgb[1];
      raw[offset + 2] = rgb[2];
    }
  }

  const compressed = zlib.deflateSync(raw);
  const idat = pngChunk('IDAT', compressed);
  const iend = pngChunk('IEND', Buffer.alloc(0));

  return Buffer.concat([signature, ihdr, idat, iend]);
}

// ---------------------------------------------------------------------------
// Minimal ZIP writer (no external deps)
// ---------------------------------------------------------------------------

function zipFiles(files) {
  // files: Array<{ name: string, data: Buffer }>
  const entries = [];
  let offset = 0;

  for (const file of files) {
    const nameBytes = Buffer.from(file.name, 'utf8');
    const crc = crc32(file.data);
    const localHeader = Buffer.concat([
      Buffer.from([0x50, 0x4b, 0x03, 0x04]), // signature
      Buffer.from([0x14, 0x00]),              // version needed: 2.0
      Buffer.from([0x00, 0x00]),              // general purpose flags
      Buffer.from([0x00, 0x00]),              // compression: stored
      Buffer.from([0x00, 0x00, 0x00, 0x00]), // mod time/date
      uint32BE(crc),                          // CRC-32
      uint32BE(file.data.length),             // compressed size
      uint32BE(file.data.length),             // uncompressed size
      Buffer.from([nameBytes.length & 0xff, (nameBytes.length >> 8) & 0xff]), // name len
      Buffer.from([0x00, 0x00]),              // extra field len
      nameBytes,
    ]);

    entries.push({ nameBytes, crc, size: file.data.length, offset, localHeader, data: file.data });
    offset += localHeader.length + file.data.length;
  }

  const centralDirChunks = [];
  for (const e of entries) {
    const central = Buffer.concat([
      Buffer.from([0x50, 0x4b, 0x01, 0x02]), // central dir signature
      Buffer.from([0x14, 0x00]),              // version made by
      Buffer.from([0x14, 0x00]),              // version needed
      Buffer.from([0x00, 0x00]),              // flags
      Buffer.from([0x00, 0x00]),              // compression
      Buffer.from([0x00, 0x00, 0x00, 0x00]), // mod time/date
      uint32BE(e.crc),
      uint32BE(e.size),
      uint32BE(e.size),
      Buffer.from([e.nameBytes.length & 0xff, (e.nameBytes.length >> 8) & 0xff]),
      Buffer.from([0x00, 0x00]),              // extra len
      Buffer.from([0x00, 0x00]),              // comment len
      Buffer.from([0x00, 0x00]),              // disk start
      Buffer.from([0x00, 0x00]),              // int attr
      Buffer.from([0x00, 0x00, 0x00, 0x00]), // ext attr
      uint32BE(e.offset),
      e.nameBytes,
    ]);
    centralDirChunks.push(central);
  }

  const centralDir = Buffer.concat(centralDirChunks);
  const centralDirOffset = offset;

  const eocd = Buffer.concat([
    Buffer.from([0x50, 0x4b, 0x05, 0x06]),              // end-of-central-dir
    Buffer.from([0x00, 0x00, 0x00, 0x00]),              // disk numbers
    Buffer.from([entries.length & 0xff, (entries.length >> 8) & 0xff]),
    Buffer.from([entries.length & 0xff, (entries.length >> 8) & 0xff]),
    uint32BE(centralDir.length),
    uint32BE(centralDirOffset),
    Buffer.from([0x00, 0x00]),                          // comment len
  ]);

  const localParts = entries.map((e) => Buffer.concat([e.localHeader, e.data]));
  return Buffer.concat([...localParts, centralDir, eocd]);
}

// ---------------------------------------------------------------------------
// Prompt helper
// ---------------------------------------------------------------------------

function prompt(rl, question) {
  return new Promise((resolve) => rl.question(question, resolve));
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  console.log('\n🚀  Icebreaker Audio Bot — Setup\n');

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

  // 1. Read the Microsoft App ID
  const appId = await prompt(
    rl,
    'Enter your Microsoft App ID (from Azure Bot Service)\n' +
      '  [press Enter to keep the placeholder]: '
  );

  // 2. Patch manifest.json
  const manifest = JSON.parse(fs.readFileSync(MANIFEST_FILE, 'utf8'));
  const newId = (appId.trim() || manifest.id).trim();
  manifest.id = newId;
  manifest.bots[0].botId = newId;
  fs.writeFileSync(MANIFEST_FILE, JSON.stringify(manifest, null, 2) + '\n', 'utf8');
  console.log(`\n✅  manifest.json updated with App ID: ${newId}`);

  // 3. Generate icons
  const colorPng = createSolidPng(192, 192, [0, 120, 212]);   // Teams blue
  const outlinePng = createSolidPng(32, 32, [255, 255, 255]); // white
  fs.writeFileSync(path.join(MANIFEST_DIR, 'color.png'), colorPng);
  fs.writeFileSync(path.join(MANIFEST_DIR, 'outline.png'), outlinePng);
  console.log('✅  Placeholder icons generated (color.png 192×192, outline.png 32×32)');
  console.log('    Replace them with your own branded icons before publishing to the Teams store.');

  // 4. Package into manifest.zip
  const zip = zipFiles([
    { name: 'manifest.json', data: Buffer.from(JSON.stringify(manifest, null, 2) + '\n', 'utf8') },
    { name: 'color.png', data: colorPng },
    { name: 'outline.png', data: outlinePng },
  ]);
  fs.writeFileSync(ZIP_OUT, zip);
  console.log(`✅  ${ZIP_OUT} created`);

  rl.close();

  console.log('\n📦  Next steps:');
  console.log('   1. Upload manifest/manifest.zip in Teams Admin Centre → Manage Apps → Upload');
  console.log('      (or use Teams Developer Portal at https://dev.teams.microsoft.com)');
  console.log('   2. Copy .env.example to .env and fill in all values');
  console.log('   3. Run: npm install && npm start\n');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
