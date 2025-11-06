/**
 * Tests for Workbook functionality
 */

const fs = require('fs');
const path = require('path');
const { Workbook } = require('../wrapper');
const { cleanupTestOutput } = require('./test-helpers');

describe('Workbook', () => {
  const testOutputDir = path.join(__dirname, 'output');

  beforeAll(() => {
    // Create output directory for test files
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
  });

  afterAll(() => cleanupTestOutput(testOutputDir));

  test('should create a new workbook', () => {
    const workbook = new Workbook();
    expect(workbook).toBeInstanceOf(Workbook);
    expect(workbook._native).toBeDefined();
  });

  test('should add a worksheet without a name', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();
    expect(worksheet).toBeDefined();
    expect(worksheet._native).toBeDefined();
  });

  test('should add a worksheet with a name', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('TestSheet');
    expect(worksheet).toBeDefined();
  });

  test('should add multiple worksheets', () => {
    const workbook = new Workbook();
    const sheet1 = workbook.addWorksheet('Sheet1');
    const sheet2 = workbook.addWorksheet('Sheet2');
    const sheet3 = workbook.addWorksheet('Sheet3');

    expect(sheet1).toBeDefined();
    expect(sheet2).toBeDefined();
    expect(sheet3).toBeDefined();
  });

  test('should save workbook to file', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();
    worksheet.write(0, 0, 'Test');

    const filename = path.join(testOutputDir, 'test-save.xlsx');
    workbook.save(filename);

    expect(fs.existsSync(filename)).toBe(true);
    expect(fs.statSync(filename).size).toBeGreaterThan(0);
  });

  test('should save workbook to buffer', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();
    worksheet.write(0, 0, 'Test');

    const buffer = workbook.saveToBuffer();

    expect(buffer).toBeInstanceOf(Buffer);
    expect(buffer.length).toBeGreaterThan(0);
  });

  test('should throw error for invalid worksheet name', () => {
    const workbook = new Workbook();
    // Excel worksheet names can't be longer than 31 characters
    expect(() => {
      workbook.addWorksheet('A'.repeat(32));
    }).toThrow();
  });
});
