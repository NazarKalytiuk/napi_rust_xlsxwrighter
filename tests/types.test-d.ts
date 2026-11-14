// Type-level contract of the public API, enforced by `npm run typecheck`.
import { expectTypeOf } from 'vitest';
import {
  Format,
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
