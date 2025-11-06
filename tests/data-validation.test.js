/**
 * Tests for DataValidation
 */

const fs = require('fs');
const path = require('path');
const { Workbook, DataValidation, Format } = require('../wrapper');
const { cleanupTestOutput } = require('./test-helpers');

describe('Data Validation', () => {
  const testOutputDir = path.join(__dirname, 'output');

  beforeAll(() => {
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
  });

  afterAll(() => cleanupTestOutput(testOutputDir));

  describe('DataValidation Creation', () => {
    test('should create a new data validation', () => {
      const validation = new DataValidation();
      expect(validation).toBeInstanceOf(DataValidation);
      expect(validation._native).toBeDefined();
    });
  });

  describe('List Validation', () => {
    test('should create dropdown with string list', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('List');

      worksheet.write(0, 0, 'Select Status:');

      const validation = new DataValidation()
        .allowListStrings(['Pending', 'In Progress', 'Completed']);

      worksheet.addDataValidation(0, 1, 0, 1, validation);

      const outputPath = path.join(testOutputDir, 'validation-list.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should create dropdown with input and error messages', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('List with Messages');

      const validation = new DataValidation()
        .allowListStrings(['Option 1', 'Option 2', 'Option 3'])
        .setInputTitle('Select Option')
        .setInputMessage('Choose one of the available options')
        .setErrorTitle('Invalid Selection')
        .setErrorMessage('Please select from the dropdown list')
        .setErrorStyle('stop');

      worksheet.addDataValidation(0, 0, 0, 0, validation);

      const outputPath = path.join(testOutputDir, 'validation-list-messages.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should create dropdown from cell range formula', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('List Formula');

      // Write list values
      worksheet.write(5, 0, 'Apple');
      worksheet.write(6, 0, 'Banana');
      worksheet.write(7, 0, 'Orange');

      // Reference the list
      const validation = new DataValidation()
        .allowListFormula('=$A$6:$A$8');

      worksheet.addDataValidation(0, 0, 0, 0, validation);

      const outputPath = path.join(testOutputDir, 'validation-list-formula.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Number Validation', () => {
    test('should validate whole number between range', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Whole Number Range');

      worksheet.write(0, 0, 'Enter Age (18-100):');

      const validation = new DataValidation()
        .allowWholeNumberBetween(18, 100)
        .setInputTitle('Age')
        .setInputMessage('Enter your age between 18 and 100');

      worksheet.addDataValidation(0, 1, 0, 1, validation);

      const outputPath = path.join(testOutputDir, 'validation-whole-number.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should validate whole number with rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Whole Number Rule');

      worksheet.write(0, 0, 'Enter value > 10:');

      const validation = new DataValidation()
        .allowWholeNumber('greaterThan', 10)
        .setErrorTitle('Too Small')
        .setErrorMessage('Value must be greater than 10');

      worksheet.addDataValidation(0, 1, 0, 1, validation);

      const outputPath = path.join(testOutputDir, 'validation-whole-rule.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should validate decimal number between range', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Decimal Range');

      worksheet.write(0, 0, 'Enter Price ($0.01 - $999.99):');

      const validation = new DataValidation()
        .allowDecimalNumberBetween(0.01, 999.99)
        .setInputTitle('Price')
        .setInputMessage('Enter a price between $0.01 and $999.99');

      worksheet.addDataValidation(0, 1, 0, 1, validation);

      const outputPath = path.join(testOutputDir, 'validation-decimal.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should validate decimal number with rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Decimal Rule');

      worksheet.write(0, 0, 'Enter value >= 0:');

      const validation = new DataValidation()
        .allowDecimalNumber('greaterThanOrEqual', 0)
        .setErrorStyle('warning')
        .setErrorTitle('Negative Value')
        .setErrorMessage('Value should not be negative');

      worksheet.addDataValidation(0, 1, 0, 1, validation);

      const outputPath = path.join(testOutputDir, 'validation-decimal-rule.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Custom Validation', () => {
    test('should validate with custom formula', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Custom');

      worksheet.write(0, 0, 'Enter Even Number:');

      const validation = new DataValidation()
        .allowCustom('=MOD(B1,2)=0')
        .setInputTitle('Even Number')
        .setInputMessage('Enter an even number')
        .setErrorTitle('Not Even')
        .setErrorMessage('The number must be even');

      worksheet.addDataValidation(0, 1, 0, 1, validation);

      const outputPath = path.join(testOutputDir, 'validation-custom.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should allow any value with info message', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Any Value');

      const validation = new DataValidation()
        .allowAnyValue()
        .setInputTitle('Information')
        .setInputMessage('You can enter any value here');

      worksheet.addDataValidation(0, 0, 0, 0, validation);

      const outputPath = path.join(testOutputDir, 'validation-any.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Validation Configuration', () => {
    test('should configure ignore blank', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Ignore Blank');

      const validation = new DataValidation()
        .allowWholeNumberBetween(1, 100)
        .ignoreBlank(false); // Don't allow blank

      worksheet.addDataValidation(0, 0, 0, 0, validation);

      const outputPath = path.join(testOutputDir, 'validation-no-blank.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should configure dropdown visibility', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('No Dropdown');

      const validation = new DataValidation()
        .allowListStrings(['A', 'B', 'C'])
        .showDropdown(false); // Hide dropdown arrow

      worksheet.addDataValidation(0, 0, 0, 0, validation);

      const outputPath = path.join(testOutputDir, 'validation-no-dropdown.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should configure message visibility', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('No Messages');

      const validation = new DataValidation()
        .allowWholeNumberBetween(1, 10)
        .showInputMessage(false)
        .showErrorMessage(false);

      worksheet.addDataValidation(0, 0, 0, 0, validation);

      const outputPath = path.join(testOutputDir, 'validation-no-messages.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Error Styles', () => {
    test('should use stop error style', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Stop');

      const validation = new DataValidation()
        .allowWholeNumberBetween(1, 10)
        .setErrorStyle('stop')
        .setErrorMessage('You cannot proceed with invalid data');

      worksheet.addDataValidation(0, 0, 0, 0, validation);

      const outputPath = path.join(testOutputDir, 'validation-error-stop.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should use warning error style', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Warning');

      const validation = new DataValidation()
        .allowWholeNumberBetween(1, 10)
        .setErrorStyle('warning')
        .setErrorMessage('Data is outside recommended range, but you can continue');

      worksheet.addDataValidation(0, 0, 0, 0, validation);

      const outputPath = path.join(testOutputDir, 'validation-error-warning.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should use information error style', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Information');

      const validation = new DataValidation()
        .allowWholeNumberBetween(1, 10)
        .setErrorStyle('information')
        .setErrorMessage('Note: Value is outside the normal range');

      worksheet.addDataValidation(0, 0, 0, 0, validation);

      const outputPath = path.join(testOutputDir, 'validation-error-info.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Multiple Validations', () => {
    test('should apply validation to multiple cells', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Multiple Cells');

      const validation = new DataValidation()
        .allowListStrings(['Red', 'Green', 'Blue']);

      // Apply to a range
      worksheet.addDataValidation(0, 0, 5, 0, validation);

      const outputPath = path.join(testOutputDir, 'validation-multiple.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should apply different validations to different ranges', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Different Validations');

      worksheet.write(0, 0, 'Status:');
      const statusValidation = new DataValidation()
        .allowListStrings(['Active', 'Inactive']);
      worksheet.addDataValidation(0, 1, 0, 1, statusValidation);

      worksheet.write(1, 0, 'Priority:');
      const priorityValidation = new DataValidation()
        .allowWholeNumberBetween(1, 5);
      worksheet.addDataValidation(1, 1, 1, 1, priorityValidation);

      worksheet.write(2, 0, 'Category:');
      const categoryValidation = new DataValidation()
        .allowListStrings(['Bug', 'Feature', 'Task']);
      worksheet.addDataValidation(2, 1, 2, 1, categoryValidation);

      const outputPath = path.join(testOutputDir, 'validation-different.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Method Chaining', () => {
    test('should support full method chaining', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Chained');

      const validation = new DataValidation()
        .allowWholeNumberBetween(1, 100)
        .ignoreBlank(true)
        .showDropdown(true)
        .showInputMessage(true)
        .showErrorMessage(true)
        .setInputTitle('Score')
        .setInputMessage('Enter a score between 1 and 100')
        .setErrorTitle('Invalid Score')
        .setErrorMessage('Score must be between 1 and 100')
        .setErrorStyle('stop');

      expect(validation).toBeInstanceOf(DataValidation);

      worksheet.addDataValidation(0, 0, 0, 0, validation);

      const outputPath = path.join(testOutputDir, 'validation-chained.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Validation to Buffer', () => {
    test('should save validation to buffer', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Buffer');

      const validation = new DataValidation()
        .allowListStrings(['A', 'B', 'C']);

      worksheet.addDataValidation(0, 0, 0, 0, validation);

      const buffer = workbook.saveToBuffer();
      expect(buffer).toBeInstanceOf(Buffer);
      expect(buffer.length).toBeGreaterThan(0);
    });
  });
});
