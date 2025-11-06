const { Workbook, Image } = require('../wrapper.js');
const fs = require('fs');
const path = require('path');

describe('Images', () => {
  let workbook, worksheet;
  const testImagePath = path.join(__dirname, 'test-image.png');

  beforeEach(() => {
    workbook = new Workbook();
    worksheet = workbook.addWorksheet('Images');
  });

  afterEach(() => {
    try {
      if (fs.existsSync('test-images.xlsx')) {
        fs.unlinkSync('test-images.xlsx');
      }
    } catch (err) {
      // Ignore cleanup errors
    }
  });

  describe('Image Creation', () => {
    it('should create image from file path', () => {
      expect(() => {
        const image = new Image(testImagePath);
        expect(image).toBeDefined();
      }).not.toThrow();
    });

    it('should create image from buffer', () => {
      const buffer = fs.readFileSync(testImagePath);
      expect(() => {
        const image = Image.fromBuffer(buffer);
        expect(image).toBeDefined();
      }).not.toThrow();
    });

    it('should throw error for invalid path', () => {
      expect(() => {
        new Image('nonexistent.png');
      }).toThrow();
    });
  });

  describe('Image Insertion', () => {
    it('should insert image at cell position', () => {
      const image = new Image(testImagePath);
      expect(() => {
        worksheet.insertImage(0, 0, image);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
      expect(fs.existsSync('test-images.xlsx')).toBe(true);
    });

    it('should insert image with offset', () => {
      const image = new Image(testImagePath);
      expect(() => {
        worksheet.insertImageWithOffset(1, 1, image, 10, 20);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });

    it('should insert multiple images', () => {
      const image1 = new Image(testImagePath);
      const image2 = new Image(testImagePath);
      expect(() => {
        worksheet.insertImage(0, 0, image1);
        worksheet.insertImage(2, 2, image2);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });

    it('should insert image fit to cell', () => {
      const image = new Image(testImagePath);
      expect(() => {
        worksheet.insertImageFitToCell(0, 0, image, true);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });

    it('should insert image fit to cell without aspect ratio', () => {
      const image = new Image(testImagePath);
      expect(() => {
        worksheet.insertImageFitToCell(0, 0, image, false);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });
  });

  describe('Image Formatting', () => {
    it('should set image width', () => {
      const image = new Image(testImagePath);
      expect(() => {
        image.setWidth(200);
        worksheet.insertImage(0, 0, image);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });

    it('should set image height', () => {
      const image = new Image(testImagePath);
      expect(() => {
        image.setHeight(150);
        worksheet.insertImage(0, 0, image);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });

    it('should set width and height', () => {
      const image = new Image(testImagePath);
      expect(() => {
        image.setWidth(300).setHeight(200);
        worksheet.insertImage(0, 0, image);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });

    it('should scale image width', () => {
      const image = new Image(testImagePath);
      expect(() => {
        image.setScaleWidth(1.5);
        worksheet.insertImage(0, 0, image);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });

    it('should scale image height', () => {
      const image = new Image(testImagePath);
      expect(() => {
        image.setScaleHeight(0.5);
        worksheet.insertImage(0, 0, image);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });

    it('should scale to specific size with aspect ratio', () => {
      const image = new Image(testImagePath);
      expect(() => {
        image.setScaleToSize(400, 300, true);
        worksheet.insertImage(0, 0, image);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });

    it('should scale to specific size without aspect ratio', () => {
      const image = new Image(testImagePath);
      expect(() => {
        image.setScaleToSize(400, 300, false);
        worksheet.insertImage(0, 0, image);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });
  });

  describe('Image Properties', () => {
    it('should set alt text', () => {
      const image = new Image(testImagePath);
      expect(() => {
        image.setAltText('Company logo');
        worksheet.insertImage(0, 0, image);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });

    it('should mark as decorative', () => {
      const image = new Image(testImagePath);
      expect(() => {
        image.setDecorative(true);
        worksheet.insertImage(0, 0, image);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });

    it('should set object movement to moveAndSize', () => {
      const image = new Image(testImagePath);
      expect(() => {
        image.setObjectMovement('moveAndSize');
        worksheet.insertImage(0, 0, image);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });

    it('should set object movement to moveOnly', () => {
      const image = new Image(testImagePath);
      expect(() => {
        image.setObjectMovement('moveOnly');
        worksheet.insertImage(0, 0, image);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });

    it('should set object movement to none', () => {
      const image = new Image(testImagePath);
      expect(() => {
        image.setObjectMovement('none');
        worksheet.insertImage(0, 0, image);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });

    it('should throw error for invalid object movement', () => {
      const image = new Image(testImagePath);
      expect(() => {
        image.setObjectMovement('invalid');
      }).toThrow('Invalid object movement option');
    });

    it('should set URL link', () => {
      const image = new Image(testImagePath);
      expect(() => {
        image.setUrl('https://example.com');
        worksheet.insertImage(0, 0, image);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });
  });

  describe('Method Chaining', () => {
    it('should chain image methods', () => {
      const image = new Image(testImagePath);
      expect(() => {
        image
          .setWidth(200)
          .setHeight(150)
          .setAltText('Product image')
          .setObjectMovement('moveOnly');

        worksheet.insertImage(0, 0, image);
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });
  });

  describe('Real World Scenarios', () => {
    it('should create report with header image', () => {
      const logo = new Image(testImagePath);
      logo.setWidth(100).setHeight(50);

      expect(() => {
        worksheet.insertImage(0, 0, logo);
        worksheet.write(2, 0, 'Monthly Report');
        worksheet.write(3, 0, 'Sales Data');
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });

    it('should create product catalog with multiple images', () => {
      expect(() => {
        for (let i = 0; i < 3; i++) {
          const image = new Image(testImagePath);
          image.setScaleWidth(0.5).setScaleHeight(0.5);
          worksheet.insertImage(i * 5, 0, image);
          worksheet.write(i * 5, 1, `Product ${i + 1}`);
        }
        workbook.save('test-images.xlsx');
      }).not.toThrow();
    });
  });
});
