// Type-level contract of the public API, enforced by `npm run typecheck`.
import { expectTypeOf } from 'vitest';
import {
  type CellValue,
  type Color,
  DataValidation,
  Format,
  Formula,
  Note,
  RichText,
  Shape,
  ShapeFont,
  ShapeFormat,
  ShapeLine,
  ShapePatternFill,
  ShapeSolidFill,
  ShapeText,
  Workbook,
  Worksheet,
  rustMemoryStats,
} from '@nazarkalytiuk/rusc-xlsx';

const format = new Format();
const worksheet = new Workbook().addWorksheet();

expectTypeOf(worksheet).toEqualTypeOf<Worksheet>();
expectTypeOf(format.setBold().setItalic()).toEqualTypeOf<Format>();

// Allocator diagnostics are a [currentBytes, peakBytes] pair.
expectTypeOf(rustMemoryStats).returns.toEqualTypeOf<[number, number]>();

// Cell values include Dates.
expectTypeOf(worksheet.write).toBeCallableWith(0, 0, new Date());
expectTypeOf(worksheet.writeRow).toBeCallableWith(0, 0, ['a', 1, true, new Date(), null]);
expectTypeOf<CellValue>().toEqualTypeOf<string | number | boolean | Date | null | undefined>();

// Formula writers take a Formula or a formula string.
expectTypeOf(worksheet.writeFormula).toBeCallableWith(0, 0, '=A1+B1');
expectTypeOf(worksheet.writeFormula).toBeCallableWith(0, 0, new Formula('=A1+B1'), format);
expectTypeOf(worksheet.writeArrayFormula).toBeCallableWith(0, 0, 1, 0, '=A1:A2*2');
expectTypeOf(worksheet.writeDynamicFormula).toBeCallableWith(0, 0, '=SORT(A1:A2)');

// Rich text segments may omit their format.
expectTypeOf(new RichText().addSegment('plain')).toEqualTypeOf<RichText>();

// Shapes only come from factories: `new Shape()` has no native handle.
expectTypeOf(Shape.textbox()).toEqualTypeOf<Shape>();
expectTypeOf<typeof Shape>().not.toExtend<abstract new (...args: never) => unknown>();

// Theme colors are colors of cell formats and worksheet tabs, but not of
// notes or conditional formats, which store colors as RGB.
expectTypeOf<'theme:4:2'>().toExtend<Color>();
expectTypeOf(format.setBackgroundColor('theme:4:2').setBorder('hair')).toEqualTypeOf<Format>();
expectTypeOf(worksheet.setTabColor).toBeCallableWith('theme:5:3');
// @ts-expect-error -- notes take a named color or a hex code
new Note('note').setBackgroundColor('theme:4:2');
// @ts-expect-error -- conditional formats take a named color or a hex code
worksheet.addConditionalFormatDataBar(0, 0, 9, 0, 'theme:4:0');

// New validation, note and view methods chain like their siblings.
expectTypeOf(
  new DataValidation().allowTextLength('lessThanOrEqual', 40),
).toEqualTypeOf<DataValidation>();
expectTypeOf(new Note('note').setFontName('Arial').setFontSize(10)).toEqualTypeOf<Note>();
expectTypeOf(worksheet.setScreenGridlines).returns.toEqualTypeOf<void>();

// Shape formatting setters chain, and the shape setters take each kind.
const shapeFormat = new ShapeFormat()
  .setSolidFill(new ShapeSolidFill().setColor('theme:4:1').setTransparency(20))
  .setLine(new ShapeLine().setColor('#000000').setWidth(0.75).setDashType('dash'));
expectTypeOf(shapeFormat).toEqualTypeOf<ShapeFormat>();
expectTypeOf(
  new ShapePatternFill().setPattern('dotted20Percent').setForegroundColor('gray'),
).toEqualTypeOf<ShapePatternFill>();
expectTypeOf(
  new ShapeFont().setName('Arial').setSize(8).setBold().setColor('red'),
).toEqualTypeOf<ShapeFont>();
expectTypeOf(
  new ShapeText().setHorizontalAlignment('center').setDirection('rotate270'),
).toEqualTypeOf<ShapeText>();
expectTypeOf(
  Shape.textbox().setFormat(shapeFormat).setFont(new ShapeFont()).setTextOptions(new ShapeText()),
).toEqualTypeOf<Shape>();
// @ts-expect-error -- shapes take a ShapeFormat, not a Format
Shape.textbox().setFormat(format);
// @ts-expect-error -- not a shape pattern
new ShapePatternFill().setPattern('dotted15Percent');
// @ts-expect-error -- shape text aligns with 'middle', not 'center'
new ShapeText().setVerticalAlignment('center');
// @ts-expect-error -- not a shape text direction
new ShapeText().setDirection('rotate180');

// String-literal unions reject unknown values.
// @ts-expect-error -- not a calculation mode
new Workbook().setCalculationMode('sometimes');
// @ts-expect-error -- not a named color or a hex code
format.setFontColor('invalid-color');
// @ts-expect-error -- not a border style
format.setBorder('wavy');
// @ts-expect-error -- arrays are not cell values
worksheet.write(0, 0, []);
// A missing value leaves the cell empty.
worksheet.write(0, 0, null);
worksheet.write(0, 0, undefined, format);
