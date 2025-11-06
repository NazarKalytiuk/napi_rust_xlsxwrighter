const { Workbook, Format, ProtectionOptions, RichText } = require('../wrapper.js');
const fs = require('fs');

describe('Protection Options', () => {
  let workbook, worksheet;

  beforeEach(() => {
    workbook = new Workbook();
    worksheet = workbook.addWorksheet('Protected');
  });

  afterEach(() => {
    try {
      if (fs.existsSync('test-protection.xlsx')) {
        fs.unlinkSync('test-protection.xlsx');
      }
    } catch (err) {
      // Ignore cleanup errors
    }
  });

  describe('ProtectionOptions Creation', () => {
    it('should create default protection options', () => {
      expect(() => {
        const options = new ProtectionOptions();
        worksheet.protectWithOptions(options);
        workbook.save('test-protection.xlsx');
      }).not.toThrow();
      expect(fs.existsSync('test-protection.xlsx')).toBe(true);
    });

    it('should support method chaining', () => {
      const options = new ProtectionOptions()
        .setInsertRows(true)
        .setInsertColumns(true)
        .setFormatCells(true);

      expect(options).toBeDefined();
      expect(() => {
        worksheet.protectWithOptions(options);
        workbook.save('test-protection.xlsx');
      }).not.toThrow();
    });
  });

  describe('Protection Permissions', () => {
    it('should allow inserting rows', () => {
      const options = new ProtectionOptions().setInsertRows(true);
      expect(() => {
        worksheet.protectWithOptions(options);
        workbook.save('test-protection.xlsx');
      }).not.toThrow();
    });

    it('should allow inserting columns', () => {
      const options = new ProtectionOptions().setInsertColumns(true);
      expect(() => {
        worksheet.protectWithOptions(options);
        workbook.save('test-protection.xlsx');
      }).not.toThrow();
    });

    it('should allow formatting cells', () => {
      const options = new ProtectionOptions().setFormatCells(true);
      expect(() => {
        worksheet.protectWithOptions(options);
        workbook.save('test-protection.xlsx');
      }).not.toThrow();
    });

    it('should allow sorting', () => {
      const options = new ProtectionOptions().setSort(true);
      expect(() => {
        worksheet.protectWithOptions(options);
        workbook.save('test-protection.xlsx');
      }).not.toThrow();
    });

    it('should allow using autofilter', () => {
      const options = new ProtectionOptions().setUseAutofilter(true);
      expect(() => {
        worksheet.protectWithOptions(options);
        workbook.save('test-protection.xlsx');
      }).not.toThrow();
    });

    it('should configure multiple permissions', () => {
      const options = new ProtectionOptions()
        .setInsertRows(true)
        .setInsertColumns(true)
        .setDeleteRows(true)
        .setDeleteColumns(true)
        .setSort(true)
        .setUseAutofilter(true);

      expect(() => {
        worksheet.write(0, 0, 'Protected Sheet');
        worksheet.protectWithOptions(options);
        workbook.save('test-protection.xlsx');
      }).not.toThrow();
    });
  });

  describe('Protection with Password', () => {
    it('should protect with password and options', () => {
      const options = new ProtectionOptions()
        .setInsertRows(true)
        .setFormatCells(false);

      expect(() => {
        worksheet.protectWithPasswordAndOptions('secret123', options);
        workbook.save('test-protection.xlsx');
      }).not.toThrow();
    });

    it('should create protected form template', () => {
      const options = new ProtectionOptions()
        .setSelectLockedCells(true)
        .setSelectUnlockedCells(true)
        .setFormatCells(false)
        .setFormatRows(false)
        .setFormatColumns(false);

      const lockedFormat = new Format().setBold();
      const unlockedFormat = new Format().setUnlocked().setBackgroundColor('yellow');

      expect(() => {
        worksheet.write(0, 0, 'Name:', lockedFormat);
        worksheet.write(0, 1, '', unlockedFormat);
        worksheet.write(1, 0, 'Email:', lockedFormat);
        worksheet.write(1, 1, '', unlockedFormat);

        worksheet.protectWithPasswordAndOptions('form123', options);
        workbook.save('test-protection.xlsx');
      }).not.toThrow();
    });
  });

  describe('All Protection Methods', () => {
    it('should use all protection setter methods', () => {
      const options = new ProtectionOptions()
        .setSelectLockedCells(false)
        .setSelectUnlockedCells(true)
        .setFormatCells(true)
        .setFormatColumns(true)
        .setFormatRows(true)
        .setInsertColumns(true)
        .setInsertRows(true)
        .setInsertLinks(true)
        .setDeleteColumns(true)
        .setDeleteRows(true)
        .setSort(true)
        .setUseAutofilter(true)
        .setUsePivotTables(true)
        .setEditScenarios(true)
        .setEditObjects(true)
        .setContents(false);

      expect(() => {
        worksheet.protectWithOptions(options);
        workbook.save('test-protection.xlsx');
      }).not.toThrow();
    });
  });
});

describe('Rich Text Strings', () => {
  let workbook, worksheet;

  beforeEach(() => {
    workbook = new Workbook();
    worksheet = workbook.addWorksheet('RichText');
  });

  afterEach(() => {
    try {
      if (fs.existsSync('test-richtext.xlsx')) {
        fs.unlinkSync('test-richtext.xlsx');
      }
    } catch (err) {
      // Ignore cleanup errors
    }
  });

  describe('RichText Creation', () => {
    it('should create empty rich text', () => {
      const richText = new RichText();
      expect(richText).toBeDefined();
    });

    it('should add segments', () => {
      const richText = new RichText();
      const normalFormat = new Format();

      expect(() => {
        richText.addSegment('Hello', normalFormat);
      }).not.toThrow();
    });
  });

  describe('Writing Rich Text', () => {
    it('should write simple rich text string', () => {
      const normal = new Format();
      const bold = new Format().setBold();

      const richText = new RichText()
        .addSegment('This is ', normal)
        .addSegment('bold', bold);

      expect(() => {
        worksheet.writeRichString(0, 0, richText);
        workbook.save('test-richtext.xlsx');
      }).not.toThrow();
      expect(fs.existsSync('test-richtext.xlsx')).toBe(true);
    });

    it('should write rich text with multiple formats', () => {
      const normal = new Format();
      const bold = new Format().setBold();
      const italic = new Format().setItalic();
      const boldItalic = new Format().setBold().setItalic();

      const richText = new RichText()
        .addSegment('Normal ', normal)
        .addSegment('bold ', bold)
        .addSegment('italic ', italic)
        .addSegment('both', boldItalic);

      expect(() => {
        worksheet.writeRichString(0, 0, richText);
        workbook.save('test-richtext.xlsx');
      }).not.toThrow();
    });

    it('should write rich text with colors', () => {
      const red = new Format().setFontColor('red').setBold();
      const blue = new Format().setFontColor('blue').setItalic();
      const green = new Format().setFontColor('green');

      const richText = new RichText()
        .addSegment('Red ', red)
        .addSegment('blue ', blue)
        .addSegment('and green', green);

      expect(() => {
        worksheet.writeRichString(0, 0, richText);
        workbook.save('test-richtext.xlsx');
      }).not.toThrow();
    });

    it('should write rich text with cell format', () => {
      const bold = new Format().setBold();
      const italic = new Format().setItalic();

      const richText = new RichText()
        .addSegment('Bold ', bold)
        .addSegment('and italic', italic);

      const cellFormat = new Format()
        .setAlign('center')
        .setBackgroundColor('yellow');

      expect(() => {
        worksheet.writeRichStringWithFormat(0, 0, richText, cellFormat);
        workbook.save('test-richtext.xlsx');
      }).not.toThrow();
    });
  });

  describe('Complex Rich Text', () => {
    it('should create formatted document with rich text', () => {
      const header = new Format().setBold().setFontSize(14).setFontColor('blue');
      const normal = new Format();
      const important = new Format().setFontColor('red').setBold();
      const code = new Format().setFontName('Courier New').setBackgroundColor('silver');

      const title = new RichText()
        .addSegment('Document ', header)
        .addSegment('Title', header);

      const paragraph = new RichText()
        .addSegment('This is ', normal)
        .addSegment('important', important)
        .addSegment(' text with ', normal)
        .addSegment('code', code)
        .addSegment(' formatting.', normal);

      expect(() => {
        worksheet.writeRichString(0, 0, title);
        worksheet.writeRichString(2, 0, paragraph);
        worksheet.setColumnWidth(0, 50);
        workbook.save('test-richtext.xlsx');
      }).not.toThrow();
    });

    it('should write multiple rich text cells', () => {
      const bold = new Format().setBold();
      const normal = new Format();

      for (let row = 0; row < 5; row++) {
        const richText = new RichText()
          .addSegment('Row ', normal)
          .addSegment(`${row}`, bold)
          .addSegment(' content', normal);

        worksheet.writeRichString(row, 0, richText);
      }

      expect(() => {
        workbook.save('test-richtext.xlsx');
      }).not.toThrow();
    });
  });

  describe('Edge Cases', () => {
    it('should reject empty segments', () => {
      const format = new Format().setBold();
      const richText = new RichText().addSegment('', format);

      expect(() => {
        worksheet.writeRichString(0, 0, richText);
      }).toThrow('cannot be blank');
    });

    it('should handle long text with formatting', () => {
      const normal = new Format();
      const bold = new Format().setBold();

      const longText = 'A'.repeat(100);
      const richText = new RichText()
        .addSegment(longText, normal)
        .addSegment(' IMPORTANT', bold);

      expect(() => {
        worksheet.writeRichString(0, 0, richText);
        workbook.save('test-richtext.xlsx');
      }).not.toThrow();
    });
  });
});

describe('Combined Protection and Rich Text', () => {
  it('should use both features together', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Combined');

    // Create rich text header
    const bold = new Format().setBold().setFontColor('blue');
    const normal = new Format();
    const header = new RichText()
      .addSegment('Protected ', bold)
      .addSegment('Document', normal);

    // Write header
    worksheet.writeRichString(0, 0, header);

    // Add editable field
    const unlocked = new Format().setUnlocked().setBackgroundColor('yellow');
    worksheet.write(2, 0, 'Enter data here:', unlocked);

    // Protect with options
    const options = new ProtectionOptions()
      .setSelectUnlockedCells(true)
      .setFormatCells(false);

    expect(() => {
      worksheet.protectWithOptions(options);
      workbook.save('test-protection.xlsx');
    }).not.toThrow();

    // Cleanup
    try {
      if (fs.existsSync('test-protection.xlsx')) {
        fs.unlinkSync('test-protection.xlsx');
      }
    } catch (err) {}
  });
});
