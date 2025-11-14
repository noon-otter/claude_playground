#!/usr/bin/env node

/**
 * Simple icon generator for Office Add-in
 * Creates placeholder PNG icons if they don't exist
 */

const fs = require('fs');
const path = require('path');

// SVG template for icons (will be converted to PNG-like format)
const createSVG = (size) => `
<svg width="${size}" height="${size}" xmlns="http://www.w3.org/2000/svg">
  <rect width="${size}" height="${size}" fill="#0078D4"/>
  <rect x="${size * 0.2}" y="${size * 0.2}" width="${size * 0.6}" height="${size * 0.6}" fill="white" opacity="0.3"/>
  <text x="${size / 2}" y="${size / 2}" font-family="Arial" font-size="${size * 0.4}" fill="white" text-anchor="middle" dominant-baseline="middle">M</text>
</svg>
`.trim();

// Create a simple base64 encoded 1x1 PNG (blue)
const createSimplePNG = () => {
  // This is a valid 1x1 blue PNG in base64
  const base64PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
  return Buffer.from(base64PNG, 'base64');
};

// Create a proper PNG file with color
const createColoredPNG = (size, color = '#0078D4') => {
  // Convert hex to RGB
  const r = parseInt(color.substr(1, 2), 16);
  const g = parseInt(color.substr(3, 2), 16);
  const b = parseInt(color.substr(5, 2), 16);

  // Create a minimal valid PNG file
  // This is a simplified approach - creates a solid color square
  const chunks = [];

  // PNG signature
  chunks.push(Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]));

  // IHDR chunk
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);  // Width
  ihdr.writeUInt32BE(size, 4);  // Height
  ihdr.writeUInt8(8, 8);        // Bit depth
  ihdr.writeUInt8(2, 9);        // Color type (RGB)
  ihdr.writeUInt8(0, 10);       // Compression
  ihdr.writeUInt8(0, 11);       // Filter
  ihdr.writeUInt8(0, 12);       // Interlace
  chunks.push(createChunk('IHDR', ihdr));

  // IDAT chunk (image data) - simplified solid color
  const scanlineSize = size * 3 + 1; // RGB + filter byte
  const idatData = Buffer.alloc(scanlineSize * size);

  for (let y = 0; y < size; y++) {
    const offset = y * scanlineSize;
    idatData[offset] = 0; // Filter type: None
    for (let x = 0; x < size; x++) {
      const pixelOffset = offset + 1 + x * 3;
      idatData[pixelOffset] = r;
      idatData[pixelOffset + 1] = g;
      idatData[pixelOffset + 2] = b;
    }
  }

  // Compress IDAT data (simplified - just wrap in zlib)
  const zlib = require('zlib');
  const compressed = zlib.deflateSync(idatData);
  chunks.push(createChunk('IDAT', compressed));

  // IEND chunk
  chunks.push(createChunk('IEND', Buffer.alloc(0)));

  return Buffer.concat(chunks);
};

function createChunk(type, data) {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length, 0);

  const typeBuffer = Buffer.from(type, 'ascii');
  const crc = require('zlib').crc32(Buffer.concat([typeBuffer, data]));
  const crcBuffer = Buffer.alloc(4);
  crcBuffer.writeUInt32BE(crc, 0);

  return Buffer.concat([length, typeBuffer, data, crcBuffer]);
}

const assetsDir = path.join(__dirname, 'assets');
const sizes = [16, 32, 64, 80];

console.log('🎨 Generating placeholder icons...\n');

// Ensure assets directory exists
if (!fs.existsSync(assetsDir)) {
  fs.mkdirSync(assetsDir, { recursive: true });
}

sizes.forEach(size => {
  const iconPath = path.join(assetsDir, `icon-${size}.png`);

  // Only create if doesn't exist or is a text file
  let shouldCreate = false;

  if (!fs.existsSync(iconPath)) {
    shouldCreate = true;
  } else {
    const content = fs.readFileSync(iconPath);
    // Check if it's a text file (not a real PNG)
    if (!content.includes(Buffer.from([0x89, 0x50, 0x4E, 0x47]))) {
      shouldCreate = true;
    }
  }

  if (shouldCreate) {
    try {
      const pngBuffer = createColoredPNG(size, '#0078D4');
      fs.writeFileSync(iconPath, pngBuffer);
      console.log(`✅ Created icon-${size}.png`);
    } catch (error) {
      console.log(`⚠️  Could not create icon-${size}.png:`, error.message);
    }
  } else {
    console.log(`✓ icon-${size}.png already exists`);
  }
});

console.log('\n✨ Icon generation complete!');
console.log('\nNote: These are placeholder icons. For production, replace with proper brand icons.\n');
