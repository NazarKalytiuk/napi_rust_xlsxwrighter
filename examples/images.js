/**
 * Images example - demonstrates inserting images from files and buffers
 *
 * Note: This example creates a small PNG programmatically so it runs
 * without needing external image files.
 */

const { Workbook, Image } = require('@nazarkalytiuk/rusc-xlsx');
const fs = require('fs');
const path = require('path');

// Create a minimal valid 2x2 red PNG for the demo
// (PNG header + IHDR + IDAT + IEND)
function createDemoPng() {
  const buf = Buffer.from([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, // PNG signature
    0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
    0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x02, // 2x2 px
    0x08, 0x02, 0x00, 0x00, 0x00, 0xfd, 0xd4, 0x9a, 0x73, // 8-bit RGB
    0x00, 0x00, 0x00, 0x12, 0x49, 0x44, 0x41, 0x54, // IDAT chunk
    0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00, 0x00,
    0x00, 0x06, 0x00, 0x01, 0x00, 0x00, 0x02, 0x00,
    0x02, 0x1a, 0xd5, 0xc4, 0xc0, 0x1c,
    0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, // IEND chunk
    0xae, 0x42, 0x60, 0x82,
  ]);
  return buf;
}

const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Images');

worksheet.setColumnWidth(0, 25);
worksheet.setColumnWidth(2, 25);

worksheet.write(0, 0, 'Image from Buffer:');

// Insert image from a buffer
const pngBuffer = createDemoPng();
const img1 = Image.fromBuffer(pngBuffer);
img1.setScaleWidth(50);
img1.setScaleHeight(50);
img1.setAltText('Demo red square image');
worksheet.insertImage(1, 0, img1);

worksheet.write(5, 0, 'Image with URL link:');

const img2 = Image.fromBuffer(pngBuffer);
img2.setScaleWidth(30);
img2.setScaleHeight(30);
img2.setUrl('https://github.com/NazarKalytiuk/napi_rust_xlsxwrighter');
img2.setAltText('Click to visit the project');
worksheet.insertImage(6, 0, img2);

worksheet.write(10, 0, 'Decorative image:');

const img3 = Image.fromBuffer(pngBuffer);
img3.setScaleWidth(20);
img3.setScaleHeight(20);
img3.setDecorative(true);
worksheet.insertImage(11, 0, img3);

workbook.save('images.xlsx');
console.log('Created images.xlsx with images from buffers, URL links, and decorative mode');
