/**
 * Tests for Notes (Comments) functionality
 */

import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { Workbook, Note } from '../wrapper.ts';
import { cleanupTestOutput, createTestOutputDir, readXlsxPart } from './test-helpers.mjs';

describe('Notes (Comments)', () => {
  let testOutputDir;

  beforeAll(() => {
    testOutputDir = createTestOutputDir();
  });

  afterAll(() => cleanupTestOutput(testOutputDir));

  describe('Note Creation', () => {
    test('should create a new note', () => {
      const note = new Note('Test note');
      expect(note).toBeInstanceOf(Note);
      expect(note._native).toBeDefined();
    });

    test('should create note with text', () => {
      expect(new Note('This is a test note')).toBeInstanceOf(Note);
    });
  });

  describe('Note Properties', () => {
    test('should set author', () => {
      const note = new Note('Test');
      const result = note.setAuthor('John Doe');
      expect(result).toBe(note); // Check chaining
    });

    test('should reject author name that is too long', () => {
      const note = new Note('Test');
      const longName = 'a'.repeat(60); // More than 52 characters
      expect(() => {
        note.setAuthor(longName);
      }).toThrow('Author name must be less than 52 characters');
    });

    test('should set width', () => {
      const note = new Note('Test');
      expect(() => {
        note.setWidth(200);
      }).not.toThrow();
    });

    test('should set height', () => {
      const note = new Note('Test');
      expect(() => {
        note.setHeight(100);
      }).not.toThrow();
    });

    test('should set background color with named color', () => {
      const note = new Note('Test');
      expect(() => {
        note.setBackgroundColor('yellow');
      }).not.toThrow();
    });

    test('should set background color with hex code', () => {
      const note = new Note('Test');
      expect(() => {
        note.setBackgroundColor('#FFFFE1');
      }).not.toThrow();
    });

    test('should set visible', () => {
      const note = new Note('Test');
      expect(() => {
        note.setVisible(true);
      }).not.toThrow();
    });

    test('should set alt text', () => {
      const note = new Note('Test');
      expect(() => {
        note.setAltText('Accessibility description');
      }).not.toThrow();
    });
  });

  describe('Note Method Chaining', () => {
    test('should support chaining all note methods', () => {
      const note = new Note('Important note');
      const result = note
        .setAuthor('Jane Smith')
        .setWidth(250)
        .setHeight(120)
        .setBackgroundColor('#FFB6C1')
        .setVisible(true)
        .setAltText('Important information');

      expect(result).toBe(note);
    });
  });

  describe('Worksheet Integration', () => {
    test('should insert note to worksheet', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Notes');

      const note = new Note('Test note');
      note.setAuthor('Test Author');

      expect(() => {
        worksheet.insertNote(0, 0, note);
      }).not.toThrow();
    });

    test('should set default note author', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Notes');

      expect(() => {
        worksheet.setDefaultNoteAuthor('Default Author');
      }).not.toThrow();
    });

    test('should reject default author name that is too long', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Notes');

      const longName = 'a'.repeat(60);
      expect(() => {
        worksheet.setDefaultNoteAuthor(longName);
      }).toThrow('Author name must be less than 52 characters');
    });
  });

  describe('Integration Tests', () => {
    test('should create file with notes', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Task List');

      // Set default author
      worksheet.setDefaultNoteAuthor('Project Manager');

      // Write some data
      worksheet.write(0, 0, 'Task');
      worksheet.write(0, 1, 'Status');

      worksheet.write(1, 0, 'Design wireframes');
      worksheet.write(2, 0, 'Implement API');
      worksheet.write(3, 0, 'Write tests');

      // Add basic note
      const note1 = new Note('Deadline: Friday');
      note1.setAuthor('PM');
      worksheet.insertNote(1, 0, note1);

      // Add custom styled note
      const note2 = new Note('URGENT: Review this ASAP!');
      note2.setAuthor('Team Lead');
      note2.setWidth(250);
      note2.setHeight(100);
      note2.setBackgroundColor('#FFB6C1');
      note2.setVisible(true);
      worksheet.insertNote(2, 0, note2);

      // Add note with alt text
      const note3 = new Note('Remember to add edge cases');
      note3.setAuthor('QA Lead');
      note3.setAltText('Testing reminder note');
      worksheet.insertNote(3, 0, note3);

      const outputPath = path.join(testOutputDir, 'notes-test.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
      const stats = fs.statSync(outputPath);
      expect(stats.size).toBeGreaterThan(0);
    });

    test('should create file with multiple notes', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Multiple Notes');

      // Add notes to multiple cells
      for (let i = 0; i < 5; i++) {
        worksheet.write(i, 0, `Cell ${i}`);

        const note = new Note(`Note for cell ${i}`);
        note.setAuthor(`Author ${i}`);
        note.setWidth(150 + i * 20);
        note.setHeight(80 + i * 10);
        worksheet.insertNote(i, 0, note);
      }

      const outputPath = path.join(testOutputDir, 'multiple-notes-test.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should support notes with various colors', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Colored Notes');

      const colors = ['#FFFFE1', '#FFB6C1', '#E6E6FA', '#FFE4B5', '#98FB98'];

      colors.forEach((color, index) => {
        worksheet.write(index, 0, `Color ${index + 1}`);

        const note = new Note(`Note with color ${color}`);
        note.setBackgroundColor(color);
        worksheet.insertNote(index, 0, note);
      });

      const outputPath = path.join(testOutputDir, 'colored-notes-test.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });
});

/** The `comments1.xml` and `vmlDrawing1.vml` parts written for a note on A1. */
function noteParts(note) {
  const workbook = new Workbook();
  workbook.addWorksheet().insertNote(0, 0, note);
  const xlsx = workbook.saveToBuffer();
  return {
    comments: readXlsxPart(xlsx, 'xl/comments1.xml'),
    vml: readXlsxPart(xlsx, 'xl/drawings/vmlDrawing1.vml'),
  };
}

describe('Note XLSX output', () => {
  test('writes the font name and size of the note text', () => {
    const { comments } = noteParts(
      new Note('Check the part number').setFontName('Arial').setFontSize(10.5),
    );
    expect(comments).toContain(
      '<rPr><sz val="10.5"/><color indexed="81"/><rFont val="Arial"/><family val="2"/></rPr>',
    );
  });

  test('writes Tahoma 8 by default', () => {
    const { comments } = noteParts(new Note('Default font'));
    expect(comments).toContain('<sz val="8"/>');
    expect(comments).toContain('<rFont val="Tahoma"/>');
  });

  test('writes the background color into the VML shape', () => {
    const { vml } = noteParts(new Note('Colored').setBackgroundColor('#FFB6C1'));
    expect(vml).toContain('fillcolor="#ffb6c1"');
  });

  test('rejects a theme background color, which notes store as RGB', () => {
    expect(() => new Note('Theme').setBackgroundColor('theme:4:2')).toThrow(
      "Invalid color: theme:4:2. Theme colors aren't supported here; use a named or #RRGGBB color",
    );
  });
});
