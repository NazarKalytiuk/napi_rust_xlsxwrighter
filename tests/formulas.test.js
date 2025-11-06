/**
 * Tests for Formula and formula writing methods
 */

const fs = require('fs');
const path = require('path');
const { Workbook, Formula, Format } = require('../wrapper');

describe('Formulas', () => {
  const testOutputDir = path.join(__dirname, 'output');

  beforeAll(() => {
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
  });

  afterAll(() => {
    if (fs.existsSync(testOutputDir)) {
      fs.rmSync(testOutputDir, { recursive: true, force: true });
    }
  });

  describe('Formula Creation', () => {
    test('should create a new formula', () => {
      const formula = new Formula('=A1+B1');
      expect(formula).toBeInstanceOf(Formula);
      expect(formula._native).toBeDefined();
    });

    test('should set result for formula', () => {
      const formula = new Formula('=2+2');
      const result = formula.setResult('4');
      expect(result).toBe(formula); // Returns this for chaining
    });

    test('should support method chaining', () => {
      const formula = new Formula('=10*5').setResult('50');
      expect(formula).toBeInstanceOf(Formula);
    });
  });

  describe('Write Formula', () => {
    test('should write formula to cell', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Formulas');

      worksheet.write(0, 0, 10);
      worksheet.write(0, 1, 20);
      worksheet.writeFormula(0, 2, new Formula('=A1+B1'));

      const outputPath = path.join(testOutputDir, 'formula-basic.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should write formula with format', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Formulas');

      const format = new Format().setBold().setFontColor('red');

      worksheet.write(0, 0, 10);
      worksheet.write(0, 1, 20);
      worksheet.writeFormula(0, 2, new Formula('=A1+B1'), format);

      const outputPath = path.join(testOutputDir, 'formula-formatted.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should write SUM formula', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Sum');

      // Write data
      for (let i = 0; i < 5; i++) {
        worksheet.write(i, 0, (i + 1) * 10);
      }

      // Write SUM formula
      worksheet.writeFormula(5, 0, new Formula('=SUM(A1:A5)'));

      const outputPath = path.join(testOutputDir, 'formula-sum.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should write multiple formula types', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Various');

      worksheet.write(0, 0, 100);
      worksheet.write(1, 0, 50);

      // Different formula types
      worksheet.writeFormula(0, 1, new Formula('=A1*2')); // Multiplication
      worksheet.writeFormula(1, 1, new Formula('=A2/2')); // Division
      worksheet.writeFormula(2, 1, new Formula('=AVERAGE(A1:A2)')); // Average
      worksheet.writeFormula(3, 1, new Formula('=MAX(A1:A2)')); // Max
      worksheet.writeFormula(4, 1, new Formula('=MIN(A1:A2)')); // Min

      const outputPath = path.join(testOutputDir, 'formula-various.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should write formula with preset result', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Result');

      const formula = new Formula('=10*10').setResult('100');
      worksheet.writeFormula(0, 0, formula);

      const outputPath = path.join(testOutputDir, 'formula-result.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Write Array Formula', () => {
    test('should write array formula', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Array');

      // Write data
      for (let i = 0; i < 5; i++) {
        worksheet.write(i, 0, (i + 1) * 10);
        worksheet.write(i, 1, 2);
      }

      // Array formula to multiply arrays
      worksheet.writeArrayFormula(0, 2, 4, 2, new Formula('=A1:A5*B1:B5'));

      const outputPath = path.join(testOutputDir, 'array-formula.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should write array formula with format', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Array Formatted');

      const format = new Format().setNumFormat('#,##0.00');

      worksheet.write(0, 0, 100);
      worksheet.write(1, 0, 200);
      worksheet.write(0, 1, 1.5);
      worksheet.write(1, 1, 2.5);

      worksheet.writeArrayFormula(0, 2, 1, 2, new Formula('=A1:A2*B1:B2'), format);

      const outputPath = path.join(testOutputDir, 'array-formula-formatted.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Write Dynamic Formula', () => {
    test('should write dynamic array formula', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Dynamic');

      // Write data
      worksheet.write(0, 0, 'Name');
      worksheet.write(0, 1, 'Value');
      worksheet.write(1, 0, 'Item 1');
      worksheet.write(1, 1, 10);
      worksheet.write(2, 0, 'Item 2');
      worksheet.write(2, 1, 20);
      worksheet.write(3, 0, 'Item 3');
      worksheet.write(3, 1, 30);

      // Dynamic formula (note: requires Excel 365)
      // SORT function will dynamically spill results
      worksheet.writeDynamicFormula(5, 0, new Formula('=SORT(A2:B4,2,-1)'));

      const outputPath = path.join(testOutputDir, 'dynamic-formula.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should write dynamic formula with format', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Dynamic Formatted');

      const format = new Format().setBold();

      // UNIQUE function (Excel 365)
      worksheet.write(0, 0, 'Apple');
      worksheet.write(1, 0, 'Banana');
      worksheet.write(2, 0, 'Apple');
      worksheet.write(3, 0, 'Cherry');

      worksheet.writeDynamicFormula(0, 2, new Formula('=UNIQUE(A1:A4)'), format);

      const outputPath = path.join(testOutputDir, 'dynamic-formula-formatted.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Complex Formula Scenarios', () => {
    test('should create spreadsheet with multiple formula types', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Complex');

      // Headers
      const headerFormat = new Format().setBold().setBackgroundColor('#D3D3D3');
      worksheet.write(0, 0, 'Item', headerFormat);
      worksheet.write(0, 1, 'Quantity', headerFormat);
      worksheet.write(0, 2, 'Price', headerFormat);
      worksheet.write(0, 3, 'Total', headerFormat);

      // Data
      const items = [
        ['Apples', 10, 1.5],
        ['Bananas', 5, 0.75],
        ['Oranges', 8, 2.0],
      ];

      items.forEach((item, i) => {
        const row = i + 1;
        worksheet.write(row, 0, item[0]);
        worksheet.write(row, 1, item[1]);
        worksheet.write(row, 2, item[2]);
        worksheet.writeFormula(row, 3, new Formula(`=B${row+1}*C${row+1}`));
      });

      // Totals
      worksheet.write(4, 2, 'Total:', new Format().setBold());
      worksheet.writeFormula(4, 3, new Formula('=SUM(D2:D4)'), new Format().setBold());

      const outputPath = path.join(testOutputDir, 'complex-formulas.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should write formulas with cell references', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('References');

      // Absolute references
      worksheet.write(0, 0, 10);
      worksheet.writeFormula(0, 1, new Formula('=$A$1*2'));
      worksheet.writeFormula(1, 1, new Formula('=$A$1*3'));
      worksheet.writeFormula(2, 1, new Formula('=$A$1*4'));

      // Relative references
      worksheet.write(4, 0, 5);
      worksheet.write(5, 0, 10);
      worksheet.write(6, 0, 15);
      worksheet.writeFormula(4, 1, new Formula('=A5*2'));
      worksheet.writeFormula(5, 1, new Formula('=A6*2'));
      worksheet.writeFormula(6, 1, new Formula('=A7*2'));

      const outputPath = path.join(testOutputDir, 'formula-references.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Formula to Buffer', () => {
    test('should save formulas to buffer', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Buffer Test');

      worksheet.write(0, 0, 10);
      worksheet.write(0, 1, 20);
      worksheet.writeFormula(0, 2, new Formula('=A1+B1'));

      const buffer = workbook.saveToBuffer();
      expect(buffer).toBeInstanceOf(Buffer);
      expect(buffer.length).toBeGreaterThan(0);
    });
  });
});
