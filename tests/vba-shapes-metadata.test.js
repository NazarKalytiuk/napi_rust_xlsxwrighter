/**
 * Tests for VBA, Buttons, Shapes, and Metadata features
 */

const fs = require('fs');
const path = require('path');
const { Workbook, DocProperties, Button, Shape } = require('../wrapper');
const { cleanupTestOutput } = require('./test-helpers');

describe('VBA, Buttons, Shapes, and Metadata Features', () => {
  const testOutputDir = path.join(__dirname, 'output');

  beforeAll(() => {
    // Create output directory for test files
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
  });

  afterAll(() => cleanupTestOutput(testOutputDir));

  describe('Named Ranges', () => {
    test('should define a global named range', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Data');

      // Write some data
      worksheet.write(0, 0, 'Item');
      worksheet.write(1, 0, 'Apple');
      worksheet.write(2, 0, 'Banana');
      worksheet.write(3, 0, 'Cherry');

      // Define a global named range
      expect(() => {
        workbook.defineName('Items', '=Data!$A$2:$A$4');
      }).not.toThrow();

      workbook.save(path.join(testOutputDir, 'named_range_global.xlsx'));
    });

    test('should define a local named range', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');

      worksheet.write(0, 0, 100);
      worksheet.write(0, 1, 200);

      // Define a local named range (Sheet1!LocalName)
      expect(() => {
        workbook.defineName('Sheet1!LocalRange', '=Sheet1!$A$1:$B$1');
      }).not.toThrow();

      workbook.save(path.join(testOutputDir, 'named_range_local.xlsx'));
    });

    test('should define a formula name', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');

      worksheet.write(0, 0, 10);
      worksheet.write(0, 1, 20);

      // Define a formula
      expect(() => {
        workbook.defineName('TotalSum', '=SUM(Sheet1!$A$1:$B$1)');
      }).not.toThrow();

      workbook.save(path.join(testOutputDir, 'named_formula.xlsx'));
    });
  });

  describe('Document Properties', () => {
    test('should create DocProperties instance', () => {
      const props = new DocProperties();
      expect(props).toBeInstanceOf(DocProperties);
      expect(props._native).toBeDefined();
    });

    test('should set all document metadata properties', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');
      worksheet.write(0, 0, 'Test Data');

      const props = new DocProperties();

      // Test fluent API
      const result = props
        .setAuthor('John Doe')
        .setTitle('Test Document')
        .setSubject('Testing Metadata')
        .setComment('This is a test document')
        .setCompany('ACME Corporation')
        .setManager('Jane Smith')
        .setStatus('Draft')
        .setCategory('Reports')
        .setKeywords('test, metadata, excel')
        .setHyperlinkBase('https://example.com');

      // Verify fluent API returns this
      expect(result).toBe(props);

      // Set properties on workbook
      expect(() => {
        workbook.setProperties(props);
      }).not.toThrow();

      workbook.save(path.join(testOutputDir, 'document_properties.xlsx'));
    });

    test('should set partial document properties', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');

      const props = new DocProperties()
        .setAuthor('Alice')
        .setTitle('Quarterly Report');

      expect(() => {
        workbook.setProperties(props);
      }).not.toThrow();

      workbook.save(path.join(testOutputDir, 'partial_properties.xlsx'));
    });
  });

  describe('Read-Only Recommended', () => {
    test('should set workbook as read-only recommended', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');
      worksheet.write(0, 0, 'Sensitive Data');

      expect(() => {
        workbook.readOnlyRecommended();
      }).not.toThrow();

      workbook.save(path.join(testOutputDir, 'read_only_recommended.xlsx'));
    });
  });

  describe('Buttons', () => {
    test('should create Button instance', () => {
      const button = new Button();
      expect(button).toBeInstanceOf(Button);
      expect(button._native).toBeDefined();
    });

    test('should set button properties with fluent API', () => {
      const button = new Button();

      const result = button
        .setCaption('Click Me')
        .setMacro('MyMacro')
        .setWidth(100)
        .setHeight(30)
        .setAltText('Button for running macro');

      // Verify fluent API returns this
      expect(result).toBe(button);
    });

    test('should insert button to worksheet', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');

      const button = new Button();
      button
        .setCaption('Run Report')
        .setMacro('GenerateReport')
        .setWidth(120)
        .setHeight(35);

      expect(() => {
        worksheet.insertButton(2, 1, button);
      }).not.toThrow();

      workbook.save(path.join(testOutputDir, 'button.xlsx'));
    });

    test('should insert button with offset', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');

      const button = new Button();
      button.setCaption('Submit').setWidth(80).setHeight(25);

      expect(() => {
        worksheet.insertButtonWithOffset(1, 1, button, 10, 5);
      }).not.toThrow();

      workbook.save(path.join(testOutputDir, 'button_offset.xlsx'));
    });

    test('should set button object movement', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');

      const button1 = new Button();
      button1.setCaption('Move & Size').setObjectMovement('moveAndSize');
      worksheet.insertButton(0, 0, button1);

      const button2 = new Button();
      button2.setCaption('Move Only').setObjectMovement('moveOnly');
      worksheet.insertButton(2, 0, button2);

      const button3 = new Button();
      button3.setCaption('Fixed').setObjectMovement('none');
      worksheet.insertButton(4, 0, button3);

      workbook.save(path.join(testOutputDir, 'button_movement.xlsx'));
    });

    test('should create multiple buttons', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Dashboard');

      const buttons = [
        { caption: 'Refresh Data', macro: 'RefreshData', row: 1, col: 0 },
        { caption: 'Export CSV', macro: 'ExportCSV', row: 3, col: 0 },
        { caption: 'Print Report', macro: 'PrintReport', row: 5, col: 0 },
      ];

      buttons.forEach(btn => {
        const button = new Button();
        button.setCaption(btn.caption).setMacro(btn.macro);
        worksheet.insertButton(btn.row, btn.col, button);
      });

      workbook.save(path.join(testOutputDir, 'multiple_buttons.xlsx'));
    });
  });

  describe('Shapes (Textboxes)', () => {
    test('should create Shape instance using textbox factory', () => {
      const shape = Shape.textbox();
      expect(shape).toBeInstanceOf(Shape);
      expect(shape._native).toBeDefined();
    });

    test('should set shape properties with fluent API', () => {
      const shape = Shape.textbox();

      const result = shape
        .setText('Important Note')
        .setWidth(200)
        .setHeight(100)
        .setAltText('Note about data');

      // Verify fluent API returns this
      expect(result).toBe(shape);
    });

    test('should insert textbox to worksheet', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');

      worksheet.write(0, 0, 'Data');
      worksheet.write(1, 0, 100);

      const shape = Shape.textbox();
      shape
        .setText('This data was collected on 2025-01-01')
        .setWidth(250)
        .setHeight(60);

      expect(() => {
        worksheet.insertShape(3, 0, shape);
      }).not.toThrow();

      workbook.save(path.join(testOutputDir, 'textbox.xlsx'));
    });

    test('should insert textbox with offset', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');

      const shape = Shape.textbox();
      shape.setText('Note: Review carefully').setWidth(180).setHeight(50);

      expect(() => {
        worksheet.insertShapeWithOffset(1, 2, shape, 15, 10);
      }).not.toThrow();

      workbook.save(path.join(testOutputDir, 'textbox_offset.xlsx'));
    });

    test('should set textbox object movement', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');

      const shape1 = Shape.textbox();
      shape1.setText('Move & Size with cells').setObjectMovement('moveAndSize');
      worksheet.insertShape(0, 0, shape1);

      const shape2 = Shape.textbox();
      shape2.setText('Move but don\'t resize').setObjectMovement('moveOnly');
      worksheet.insertShape(3, 0, shape2);

      const shape3 = Shape.textbox();
      shape3.setText('Stay fixed').setObjectMovement('none');
      worksheet.insertShape(6, 0, shape3);

      workbook.save(path.join(testOutputDir, 'textbox_movement.xlsx'));
    });

    test('should create annotation with textbox', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Report');

      // Add some data
      worksheet.write(0, 0, 'Quarter');
      worksheet.write(0, 1, 'Sales');
      worksheet.write(1, 0, 'Q1');
      worksheet.write(1, 1, 15000);
      worksheet.write(2, 0, 'Q2');
      worksheet.write(2, 1, 23000);

      // Add annotation
      const annotation = Shape.textbox();
      annotation
        .setText('Note: Q2 sales increased by 53% due to new product launch')
        .setWidth(300)
        .setHeight(80)
        .setAltText('Sales growth annotation');

      worksheet.insertShape(4, 0, annotation);

      workbook.save(path.join(testOutputDir, 'textbox_annotation.xlsx'));
    });
  });

  describe('Combined Features', () => {
    test('should use all new features together', () => {
      const workbook = new Workbook();

      // Set document properties
      const props = new DocProperties();
      props
        .setAuthor('Test Automation')
        .setTitle('Complete Feature Test')
        .setCompany('Testing Inc')
        .setStatus('Final');
      workbook.setProperties(props);

      // Define named ranges
      const worksheet = workbook.addWorksheet('Data');
      worksheet.write(0, 0, 'Revenue');
      worksheet.write(1, 0, 1000);
      worksheet.write(2, 0, 2000);
      worksheet.write(3, 0, 3000);

      workbook.defineName('Revenue', '=Data!$A$2:$A$4');

      // Add button
      const button = new Button();
      button.setCaption('Calculate Total').setMacro('CalculateTotal');
      worksheet.insertButton(5, 0, button);

      // Add textbox annotation
      const note = Shape.textbox();
      note.setText('Click the button to calculate total revenue');
      worksheet.insertShape(5, 2, note);

      // Set read-only recommended
      workbook.readOnlyRecommended();

      workbook.save(path.join(testOutputDir, 'all_features_combined.xlsx'));
    });
  });
});
