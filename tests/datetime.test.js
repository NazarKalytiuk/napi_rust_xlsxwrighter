const { Workbook, Format } = require('../wrapper.js');
const fs = require('fs');

describe('DateTime Functions', () => {
  let workbook, worksheet, dateFormat, datetimeFormat;

  beforeEach(() => {
    workbook = new Workbook();
    worksheet = workbook.addWorksheet('DateTimes');
    dateFormat = new Format().setNumFormat('yyyy-mm-dd');
    datetimeFormat = new Format().setNumFormat('yyyy-mm-dd hh:mm:ss');
  });

  afterEach(() => {
    try {
      if (fs.existsSync('test-datetime.xlsx')) {
        fs.unlinkSync('test-datetime.xlsx');
      }
    } catch (err) {
      // Ignore cleanup errors
    }
  });

  describe('writeDatetimeFromTimestamp', () => {
    it('should write datetime from Unix timestamp', () => {
      expect(() => {
        worksheet.writeDatetimeFromTimestamp(0, 0, 1706184600, dateFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
      expect(fs.existsSync('test-datetime.xlsx')).toBe(true);
    });

    it('should write datetime from timestamp without format', () => {
      expect(() => {
        worksheet.writeDatetimeFromTimestamp(0, 0, 1706184600);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });

    it('should handle current timestamp', () => {
      const now = Math.floor(Date.now() / 1000);
      expect(() => {
        worksheet.writeDatetimeFromTimestamp(0, 0, now, datetimeFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });
  });

  describe('writeDatetimeFromYmd', () => {
    it('should write date from year, month, day', () => {
      expect(() => {
        worksheet.writeDatetimeFromYmd(0, 0, 2023, 1, 25, dateFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
      expect(fs.existsSync('test-datetime.xlsx')).toBe(true);
    });

    it('should write date without format', () => {
      expect(() => {
        worksheet.writeDatetimeFromYmd(0, 0, 2023, 12, 31);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });

    it('should handle leap year date', () => {
      expect(() => {
        worksheet.writeDatetimeFromYmd(0, 0, 2024, 2, 29, dateFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });

    it('should handle minimum valid date', () => {
      expect(() => {
        worksheet.writeDatetimeFromYmd(0, 0, 1900, 1, 1, dateFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });
  });

  describe('writeDatetimeFromYmdHms', () => {
    it('should write datetime with time components', () => {
      expect(() => {
        worksheet.writeDatetimeFromYmdHms(0, 0, 2023, 1, 25, 12, 30, 45.5, datetimeFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
      expect(fs.existsSync('test-datetime.xlsx')).toBe(true);
    });

    it('should write datetime without format', () => {
      expect(() => {
        worksheet.writeDatetimeFromYmdHms(0, 0, 2023, 6, 15, 9, 0, 0);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });

    it('should handle midnight time', () => {
      expect(() => {
        worksheet.writeDatetimeFromYmdHms(0, 0, 2023, 1, 1, 0, 0, 0, datetimeFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });

    it('should handle end of day time', () => {
      expect(() => {
        worksheet.writeDatetimeFromYmdHms(0, 0, 2023, 12, 31, 23, 59, 59.999, datetimeFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });

    it('should handle fractional seconds', () => {
      expect(() => {
        worksheet.writeDatetimeFromYmdHms(0, 0, 2023, 6, 15, 14, 30, 45.123, datetimeFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });
  });

  describe('writeDatetimeFromString', () => {
    it('should write datetime from ISO string', () => {
      expect(() => {
        worksheet.writeDatetimeFromString(0, 0, '2023-01-25T12:30:00', datetimeFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
      expect(fs.existsSync('test-datetime.xlsx')).toBe(true);
    });

    it('should write date from ISO string without time', () => {
      expect(() => {
        worksheet.writeDatetimeFromString(0, 0, '2023-01-25', dateFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });

    it('should handle time-only string', () => {
      expect(() => {
        const timeFormat = new Format().setNumFormat('hh:mm:ss');
        worksheet.writeDatetimeFromString(0, 0, '14:30:00', timeFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });

    it('should write datetime without format', () => {
      expect(() => {
        worksheet.writeDatetimeFromString(0, 0, '2023-06-15T09:00:00');
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });
  });

  describe('writeDatetime (convenience method)', () => {
    it('should write datetime from JavaScript Date object', () => {
      const date = new Date('2023-01-25T12:30:00');
      expect(() => {
        worksheet.writeDatetime(0, 0, date, datetimeFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
      expect(fs.existsSync('test-datetime.xlsx')).toBe(true);
    });

    it('should write datetime from current Date', () => {
      expect(() => {
        worksheet.writeDatetime(0, 0, new Date(), datetimeFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });

    it('should write datetime from timestamp number', () => {
      expect(() => {
        worksheet.writeDatetime(0, 0, 1706184600, datetimeFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });

    it('should write datetime from ISO string', () => {
      expect(() => {
        worksheet.writeDatetime(0, 0, '2023-01-25T12:30:00', datetimeFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });

    it('should write datetime from object with date only', () => {
      expect(() => {
        worksheet.writeDatetime(0, 0, { year: 2023, month: 1, day: 25 }, dateFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });

    it('should write datetime from object with date and time', () => {
      expect(() => {
        worksheet.writeDatetime(0, 0, {
          year: 2023, month: 1, day: 25,
          hour: 12, min: 30, sec: 45
        }, datetimeFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });

    it('should write datetime from object with partial time', () => {
      expect(() => {
        worksheet.writeDatetime(0, 0, {
          year: 2023, month: 6, day: 15,
          hour: 14
        }, datetimeFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });

    it('should throw error for invalid input type', () => {
      expect(() => {
        worksheet.writeDatetime(0, 0, true, dateFormat); // boolean is invalid
      }).toThrow('Invalid datetime input');
    });

    it('should write datetime without format', () => {
      expect(() => {
        worksheet.writeDatetime(0, 0, new Date());
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });
  });

  describe('Multiple datetime formats', () => {
    it('should write datetimes with different formats', () => {
      const format1 = new Format().setNumFormat('dd/mm/yyyy');
      const format2 = new Format().setNumFormat('mm/dd/yyyy');
      const format3 = new Format().setNumFormat('yyyy-mm-dd hh:mm');
      const format4 = new Format().setNumFormat('ddd dd mmm yyyy');

      expect(() => {
        worksheet.writeDatetime(0, 0, '2023-01-25', format1);
        worksheet.writeDatetime(1, 0, '2023-01-25', format2);
        worksheet.writeDatetime(2, 0, '2023-01-25T12:30:00', format3);
        worksheet.writeDatetime(3, 0, '2023-01-25', format4);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });
  });

  describe('DateTime with styling', () => {
    it('should write datetime with bold and colored format', () => {
      const styledFormat = new Format()
        .setNumFormat('yyyy-mm-dd hh:mm')
        .setBold()
        .setFontColor('blue')
        .setBackgroundColor('yellow');

      expect(() => {
        worksheet.writeDatetime(0, 0, new Date(), styledFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });
  });

  describe('DateTime series', () => {
    it('should write a series of dates', () => {
      expect(() => {
        for (let i = 0; i < 10; i++) {
          worksheet.writeDatetimeFromYmd(i, 0, 2023, 1, i + 1, dateFormat);
        }
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });

    it('should write a series of datetimes with timestamps', () => {
      const startTime = 1706184600; // Some start timestamp
      expect(() => {
        for (let i = 0; i < 10; i++) {
          worksheet.writeDatetimeFromTimestamp(i, 0, startTime + (i * 3600), datetimeFormat);
        }
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });
  });

  describe('Edge cases', () => {
    it('should handle year 9999', () => {
      expect(() => {
        worksheet.writeDatetimeFromYmd(0, 0, 9999, 12, 31, dateFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });

    it('should write multiple datetimes in same row', () => {
      expect(() => {
        worksheet.writeDatetime(0, 0, '2023-01-01', dateFormat);
        worksheet.writeDatetime(0, 1, '2023-06-15', dateFormat);
        worksheet.writeDatetime(0, 2, '2023-12-31', dateFormat);
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });
  });
});
