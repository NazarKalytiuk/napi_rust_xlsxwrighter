import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { Workbook, Format } from '../wrapper.ts';
import fs from 'node:fs';
import { cleanupTestOutput, readCellStyle, readXlsxPart } from './test-helpers.mjs';

describe('DateTime Functions', () => {
  let workbook, worksheet, dateFormat, datetimeFormat;

  beforeEach(() => {
    workbook = new Workbook();
    worksheet = workbook.addWorksheet('DateTimes');
    dateFormat = new Format().setNumFormat('yyyy-mm-dd');
    datetimeFormat = new Format().setNumFormat('yyyy-mm-dd hh:mm:ss');
  });

  afterEach(() => cleanupTestOutput('test-datetime.xlsx'));

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
        worksheet.writeDatetime(
          0,
          0,
          {
            year: 2023,
            month: 1,
            day: 25,
            hour: 12,
            min: 30,
            sec: 45,
          },
          datetimeFormat,
        );
        workbook.save('test-datetime.xlsx');
      }).not.toThrow();
    });

    it('should write datetime from object with partial time', () => {
      expect(() => {
        worksheet.writeDatetime(
          0,
          0,
          {
            year: 2023,
            month: 6,
            day: 15,
            hour: 14,
          },
          datetimeFormat,
        );
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
          worksheet.writeDatetimeFromTimestamp(i, 0, startTime + i * 3600, datetimeFormat);
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

/** The serial date value written to A1. */
function serialOf(datetime) {
  const workbook = new Workbook();
  const format = new Format().setNumFormat('yyyy-mm-dd hh:mm:ss');
  workbook.addWorksheet().writeDatetime(0, 0, datetime, format);
  const sheet = readXlsxPart(workbook.saveToBuffer(), 'xl/worksheets/sheet1.xml');
  return Number(/<c r="A1"[^>]*><v>([^<]+)<\/v>/.exec(sheet)[1]);
}

describe('writeDatetime object values in XLSX output', () => {
  // 2024-01-01 is Excel serial day 45292; a day has 1440 minutes and 86400 seconds.
  it.each([
    [{ year: 2024, month: 1, day: 1 }, 45292],
    [{ year: 2024, month: 1, day: 1, hour: 6 }, 45292 + 360 / 1440],
    [{ year: 2024, month: 1, day: 1, min: 30 }, 45292 + 30 / 1440],
    [{ year: 2024, month: 1, day: 1, sec: 45 }, 45292 + 45 / 86400],
  ])('writes %o as serial %d', (datetime, serial) => {
    expect(serialOf(datetime)).toBeCloseTo(serial, 9);
  });
});

/** The `<c>` element of a cell in the first worksheet. */
function cellXml(xlsx, cell) {
  const sheet = readXlsxPart(xlsx, 'xl/worksheets/sheet1.xml');
  return new RegExp(`<c r="${cell}"[^>]*>.*?</c>`).exec(sheet)?.[0];
}

/** The number format id of a cell's style. */
function numFmtId(xlsx, cell) {
  return /numFmtId="(\d+)"/.exec(readCellStyle(xlsx, cell).xf)[1];
}

describe('Dates as cell values', () => {
  it('writes a Date as its UTC date with the built-in date format', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.write(0, 0, new Date('2024-01-15'));

    const xlsx = workbook.saveToBuffer();
    expect(cellXml(xlsx, 'A1')).toBe('<c r="A1" s="1"><v>45306.0</v></c>');
    expect(numFmtId(xlsx, 'A1')).toBe('14');
  });

  it('writes a Date with a time, to the millisecond, with the built-in date-time format', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.write(0, 0, new Date('2024-01-15T10:30:45.250Z'));
    worksheet.writeDatetime(1, 0, new Date('2024-01-15T12:00:00.999Z'));

    const xlsx = workbook.saveToBuffer();
    expect(cellXml(xlsx, 'A1')).toBe('<c r="A1" s="1"><v>45306.43802372685</v></c>');
    expect(cellXml(xlsx, 'A2')).toBe('<c r="A2" s="1"><v>45306.5000115625</v></c>');
    expect(numFmtId(xlsx, 'A1')).toBe('22');
  });

  it('writes Dates among other values in writeRow and writeRows', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.writeRow(0, 0, ['due', new Date('2024-02-29'), 3]);
    worksheet.writeRows(1, 0, [[new Date('2024-03-01T08:00:00Z'), null]]);

    const xlsx = workbook.saveToBuffer();
    expect(cellXml(xlsx, 'B1')).toBe('<c r="B1" s="1"><v>45351.0</v></c>');
    expect(cellXml(xlsx, 'C1')).toBe('<c r="C1"><v>3.0</v></c>');
    expect(cellXml(xlsx, 'A2')).toBe('<c r="A2" s="2"><v>45352.333333333336</v></c>');
    expect(numFmtId(xlsx, 'A2')).toBe('22');
  });

  it('adds the date format to a format without a number format of its own', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();
    const bold = new Format().setBold();

    worksheet.writeRow(0, 0, ['label', new Date('2024-01-15')], bold);

    const xlsx = workbook.saveToBuffer();
    expect(numFmtId(xlsx, 'A1')).toBe('0');
    expect(numFmtId(xlsx, 'B1')).toBe('14');
    expect(readCellStyle(xlsx, 'B1').font).toContain('<b/>');
  });

  it('keeps the number format of the given format', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.write(0, 0, new Date('2024-01-15'), new Format().setNumFormat('yyyy-mm-dd'));

    const xlsx = workbook.saveToBuffer();
    expect(numFmtId(xlsx, 'A1')).toBe('164');
    expect(readXlsxPart(xlsx, 'xl/styles.xml')).toContain(
      '<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>',
    );
  });

  it('gives the datetime writers the built-in format for their value', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.writeDatetimeFromYmd(0, 0, 2023, 1, 25);
    worksheet.writeDatetimeFromYmdHms(1, 0, 2023, 1, 25, 12, 30, 0);
    worksheet.writeDatetimeFromString(2, 0, '12:30:00');
    worksheet.writeDatetimeFromTimestamp(3, 0, 1706184600);

    const xlsx = workbook.saveToBuffer();
    expect(numFmtId(xlsx, 'A1')).toBe('14');
    expect(numFmtId(xlsx, 'A2')).toBe('22');
    expect(numFmtId(xlsx, 'A3')).toBe('21');
    expect(numFmtId(xlsx, 'A4')).toBe('22');
  });

  it("counts days across Excel's 1900 leap year", () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.writeRow(0, 0, [
      new Date('1900-01-01'),
      new Date('1900-02-28'),
      new Date('1900-03-01'),
    ]);

    const xlsx = workbook.saveToBuffer();
    expect(cellXml(xlsx, 'A1')).toBe('<c r="A1" s="1"><v>1.0</v></c>');
    expect(cellXml(xlsx, 'B1')).toBe('<c r="B1" s="1"><v>59.0</v></c>');
    expect(cellXml(xlsx, 'C1')).toBe('<c r="C1" s="1"><v>61.0</v></c>');
  });

  it("rejects Dates outside Excel's 1900-9999 range", () => {
    const worksheet = new Workbook().addWorksheet();
    const message =
      "Date range error: 'Date outside Excel's range of 1900-01-01 to 9999-12-31 (UTC)'";

    expect(() => worksheet.write(0, 0, new Date('1899-12-31T23:59:59.999Z'))).toThrow(
      `Failed to write: ${message}`,
    );
    expect(() => worksheet.writeRow(1, 0, [new Date('+010000-01-01T00:00:00Z')])).toThrow(
      `Failed to write at row 1, col 0: ${message}`,
    );
    expect(() => worksheet.write(2, 0, new Date('9999-12-31T23:59:59.999Z'))).not.toThrow();
  });
});
