/**
 * create-icons.js
 * Generates minimal valid PNG icons for the add-in using only Node.js built-ins.
 * Run once: node create-icons.js
 */
const zlib = require("zlib");
const fs = require("fs");
const path = require("path");

// CRC-32 used by PNG chunk checksums
function crc32(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    crc ^= buf[i];
    for (let j = 0; j < 8; j++) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

// Wrap data in a PNG chunk: [length][type][data][crc]
function pngChunk(type, data) {
  const typeBytes = Buffer.from(type, "ascii");
  const lenBuf = Buffer.alloc(4);
  lenBuf.writeUInt32BE(data.length, 0);
  const crcInput = Buffer.concat([typeBytes, data]);
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc32(crcInput), 0);
  return Buffer.concat([lenBuf, typeBytes, data, crcBuf]);
}

/**
 * Create a solid-color N×N PNG.
 * @param {number} size  Pixel dimension (width = height)
 * @param {number} r     Red   0-255
 * @param {number} g     Green 0-255
 * @param {number} b     Blue  0-255
 */
function makeSolidPNG(size, r, g, b) {
  // PNG file signature
  const sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

  // IHDR: width(4) height(4) bitDepth(1) colorType(1=gray,2=RGB,3=idx,4=grayA,6=RGBA) comp filter interlace
  const ihdrData = Buffer.alloc(13);
  ihdrData.writeUInt32BE(size, 0); // width
  ihdrData.writeUInt32BE(size, 4); // height
  ihdrData[8] = 8;  // bit depth: 8 bits per channel
  ihdrData[9] = 2;  // color type: RGB (3 channels)
  // bytes 10-12 default to 0: compression=deflate, filter=adaptive, interlace=none

  // Raw scanlines: each row = [filter_byte=0, R, G, B, R, G, B, ...]
  const rowLen = 1 + size * 3;
  const raw = Buffer.alloc(size * rowLen);
  for (let y = 0; y < size; y++) {
    const offset = y * rowLen;
    raw[offset] = 0; // filter: None
    for (let x = 0; x < size; x++) {
      raw[offset + 1 + x * 3 + 0] = r;
      raw[offset + 1 + x * 3 + 1] = g;
      raw[offset + 1 + x * 3 + 2] = b;
    }
  }

  const compressed = zlib.deflateSync(raw);

  return Buffer.concat([
    sig,
    pngChunk("IHDR", ihdrData),
    pngChunk("IDAT", compressed),
    pngChunk("IEND", Buffer.alloc(0)),
  ]);
}

// Office blue: #0078D4 = rgb(0, 120, 212)
const R = 0, G = 120, B = 212;

const assetsDir = path.join(__dirname, "assets");
fs.mkdirSync(assetsDir, { recursive: true });

fs.writeFileSync(path.join(assetsDir, "icon-16.png"), makeSolidPNG(16, R, G, B));
fs.writeFileSync(path.join(assetsDir, "icon-32.png"), makeSolidPNG(32, R, G, B));
fs.writeFileSync(path.join(assetsDir, "icon-80.png"), makeSolidPNG(80, R, G, B));

console.log("Created assets/icon-16.png (16×16), icon-32.png (32×32), icon-80.png (80×80)");
