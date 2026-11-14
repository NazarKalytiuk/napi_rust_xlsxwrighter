// Type-level contract of the public API, enforced by `npm run typecheck`.
import { expectTypeOf } from 'vitest';
import {
  Format,
  Formula,
  RichText,
  Shape,
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
