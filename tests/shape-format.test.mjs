/**
 * Tests for the fill, line, font and text options of shapes (textboxes)
 */

import { describe, expect, test } from 'vitest';
import {
  Shape,
  ShapeFont,
  ShapeFormat,
  ShapeLine,
  ShapePatternFill,
  ShapeSolidFill,
  ShapeText,
  Workbook,
} from '../wrapper.ts';
import { readXlsxPart } from './test-helpers.mjs';

const worksheetModes = [
  ['standard', (workbook) => workbook.addWorksheet()],
  ['constant memory', (workbook) => workbook.addWorksheetWithConstantMemory()],
  ['low memory', (workbook) => workbook.addWorksheetWithLowMemory()],
];

/** The drawing XML of a workbook with `shape` inserted on a standard worksheet. */
function drawingXml(shape) {
  const workbook = new Workbook();
  workbook.addWorksheet().insertShape(1, 1, shape);
  return readXlsxPart(workbook.saveToBuffer(), 'xl/drawings/drawing1.xml');
}

/** The `<xdr:spPr>` element: the shape's geometry, fill and line. */
function shapeProperties(drawing) {
  return /<xdr:spPr>.*<\/xdr:spPr>/.exec(drawing)?.[0];
}

/** The `<a:ln>` element: the shape's border line. */
function lineXml(drawing) {
  return /<a:ln\b.*?<\/a:ln>/.exec(drawing)?.[0];
}

describe('Shape formatting XLSX output', () => {
  test.each(worksheetModes)(
    'writes a filled, lined, bold, rotated textbox on a %s worksheet',
    (_mode, addWorksheet) => {
      const workbook = new Workbook();
      const worksheet = addWorksheet(workbook);
      worksheet.write(0, 0, 'Form');
      worksheet.insertShape(
        1,
        1,
        Shape.textbox()
          .setText('Header')
          .setFormat(
            new ShapeFormat()
              .setSolidFill(new ShapeSolidFill().setColor('yellow'))
              .setLine(new ShapeLine().setColor('black').setWidth(0.75)),
          )
          .setFont(new ShapeFont().setName('Arial').setSize(8).setBold().setColor('red'))
          .setTextOptions(
            new ShapeText().setHorizontalAlignment('center').setDirection('rotate270'),
          ),
      );

      const drawing = readXlsxPart(workbook.saveToBuffer(), 'xl/drawings/drawing1.xml');
      expect(drawing).toContain('<a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill>');
      expect(drawing).toContain(
        '<a:ln w="9525" cmpd="sng"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
      );
      expect(drawing).toContain(
        '<a:rPr lang="en-US" sz="800" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:rPr>',
      );
      expect(drawing).toContain('<a:bodyPr vert="vert270" wrap="square" rtlCol="0" anchor="t"/>');
      expect(drawing).toContain('<a:p><a:pPr algn="ctr"/><a:r>');
      expect(drawing).toContain('<a:t>Header</a:t>');
    },
  );

  test('keeps the default fill, line and font without formatting', () => {
    const drawing = drawingXml(Shape.textbox().setText('Plain'));

    expect(shapeProperties(drawing)).toContain(
      '<a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:ln w="9525" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"><a:shade val="50000"/></a:schemeClr></a:solidFill></a:ln>',
    );
    expect(drawing).toContain('<a:rPr lang="en-US" sz="1100"/>');
  });

  test('writes theme colors as scheme colors', () => {
    const drawing = drawingXml(
      Shape.textbox()
        .setText('Theme')
        .setFormat(
          new ShapeFormat()
            .setSolidFill(new ShapeSolidFill().setColor('theme:4:0'))
            .setLine(new ShapeLine().setColor('theme:1:0')),
        )
        .setFont(new ShapeFont().setColor('theme:5:3')),
    );

    expect(shapeProperties(drawing)).toContain(
      '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>',
    );
    expect(lineXml(drawing)).toContain('<a:schemeClr val="tx1"/>');
    expect(drawing).toContain(
      '<a:rPr lang="en-US" sz="1100"><a:solidFill><a:schemeClr val="accent2"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr></a:solidFill></a:rPr>',
    );
  });

  test('writes a transparent solid fill', () => {
    const drawing = drawingXml(
      Shape.textbox().setFormat(
        new ShapeFormat().setSolidFill(
          new ShapeSolidFill().setColor('#4F81BD').setTransparency(25),
        ),
      ),
    );

    expect(shapeProperties(drawing)).toContain(
      '<a:solidFill><a:srgbClr val="4F81BD"><a:alpha val="75000"/></a:srgbClr></a:solidFill>',
    );
  });

  test('writes no fill and no line', () => {
    const drawing = drawingXml(
      Shape.textbox().setFormat(new ShapeFormat().setNoFill().setNoLine()),
    );

    expect(shapeProperties(drawing)).toContain(
      '<a:noFill/><a:ln w="9525" cmpd="sng"><a:noFill/></a:ln>',
    );
  });

  test.each([
    ['dotted5Percent', 'pct5'],
    ['dotted10Percent', 'pct10'],
    ['dotted20Percent', 'pct20'],
    ['dotted25Percent', 'pct25'],
    ['dotted30Percent', 'pct30'],
    ['dotted40Percent', 'pct40'],
    ['dotted50Percent', 'pct50'],
    ['dotted60Percent', 'pct60'],
    ['dotted70Percent', 'pct70'],
    ['dotted75Percent', 'pct75'],
    ['dotted80Percent', 'pct80'],
    ['dotted90Percent', 'pct90'],
    ['diagonalStripesLightDownwards', 'ltDnDiag'],
    ['diagonalStripesLightUpwards', 'ltUpDiag'],
    ['diagonalStripesDarkDownwards', 'dkDnDiag'],
    ['diagonalStripesDarkUpwards', 'dkUpDiag'],
    ['diagonalStripesWideDownwards', 'wdDnDiag'],
    ['diagonalStripesWideUpwards', 'wdUpDiag'],
    ['verticalStripesLight', 'ltVert'],
    ['horizontalStripesLight', 'ltHorz'],
    ['verticalStripesNarrow', 'narVert'],
    ['horizontalStripesNarrow', 'narHorz'],
    ['verticalStripesDark', 'dkVert'],
    ['horizontalStripesDark', 'dkHorz'],
    ['stripesBackslashes', 'dashDnDiag'],
    ['stripesForwardSlashes', 'dashUpDiag'],
    ['horizontalStripesAlternating', 'dashHorz'],
    ['verticalStripesAlternating', 'dashVert'],
    ['smallConfetti', 'smConfetti'],
    ['largeConfetti', 'lgConfetti'],
    ['zigzag', 'zigZag'],
    ['wave', 'wave'],
    ['diagonalBrick', 'diagBrick'],
    ['horizontalBrick', 'horzBrick'],
    ['weave', 'weave'],
    ['plaid', 'plaid'],
    ['divot', 'divot'],
    ['dottedGrid', 'dotGrid'],
    ['dottedDiamond', 'dotDmnd'],
    ['shingle', 'shingle'],
    ['trellis', 'trellis'],
    ['sphere', 'sphere'],
    ['smallGrid', 'smGrid'],
    ['largeGrid', 'lgGrid'],
    ['smallCheckerboard', 'smCheck'],
    ['largeCheckerboard', 'lgCheck'],
    ['outlinedDiamondGrid', 'openDmnd'],
    ['solidDiamondGrid', 'solidDmnd'],
  ])('writes the %s pattern fill as %s', (pattern, preset) => {
    const fill = new ShapePatternFill().setPattern(pattern).setForegroundColor('red');
    const drawing = drawingXml(Shape.textbox().setFormat(new ShapeFormat().setPatternFill(fill)));

    expect(shapeProperties(drawing)).toContain(
      `<a:pattFill prst="${preset}"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>`,
    );
  });

  test('writes the background color of a pattern fill', () => {
    const fill = new ShapePatternFill()
      .setPattern('dotted20Percent')
      .setForegroundColor('black')
      .setBackgroundColor('theme:0:1');
    const drawing = drawingXml(Shape.textbox().setFormat(new ShapeFormat().setPatternFill(fill)));

    expect(shapeProperties(drawing)).toContain(
      '<a:pattFill prst="pct20"><a:fgClr><a:srgbClr val="000000"/></a:fgClr><a:bgClr><a:schemeClr val="bg1"><a:lumMod val="95000"/></a:schemeClr></a:bgClr></a:pattFill>',
    );
  });

  test.each([
    ['roundDot', 'sysDot'],
    ['squareDot', 'sysDash'],
    ['dash', 'dash'],
    ['dashDot', 'dashDot'],
    ['longDash', 'lgDash'],
    ['longDashDot', 'lgDashDot'],
    ['longDashDotDot', 'lgDashDotDot'],
  ])('writes the %s line dash type as %s', (dashType, preset) => {
    const line = new ShapeLine().setColor('blue').setDashType(dashType);
    const drawing = drawingXml(Shape.textbox().setFormat(new ShapeFormat().setLine(line)));

    expect(lineXml(drawing)).toBe(
      `<a:ln w="9525" cmpd="sng"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:prstDash val="${preset}"/></a:ln>`,
    );
  });

  test('writes a solid dash type without a preset dash', () => {
    const line = new ShapeLine().setColor('blue').setDashType('solid');
    const drawing = drawingXml(Shape.textbox().setFormat(new ShapeFormat().setLine(line)));

    expect(lineXml(drawing)).toBe(
      '<a:ln w="9525" cmpd="sng"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>',
    );
  });

  test('rounds the line width to a quarter point', () => {
    const line = new ShapeLine().setColor('blue').setWidth(2.3);
    const drawing = drawingXml(Shape.textbox().setFormat(new ShapeFormat().setLine(line)));

    expect(lineXml(drawing)).toContain('<a:ln w="28575" cmpd="sng">');
  });

  test('writes a transparent line', () => {
    const line = new ShapeLine().setColor('blue').setTransparency(40);
    const drawing = drawingXml(Shape.textbox().setFormat(new ShapeFormat().setLine(line)));

    expect(lineXml(drawing)).toContain(
      '<a:srgbClr val="0000FF"><a:alpha val="60000"/></a:srgbClr>',
    );
  });

  test('hides a line and shows it again', () => {
    const hidden = drawingXml(
      Shape.textbox().setFormat(new ShapeFormat().setLine(new ShapeLine().setHidden(true))),
    );
    const shown = drawingXml(
      Shape.textbox().setFormat(
        new ShapeFormat().setLine(new ShapeLine().setHidden(true).setHidden(false)),
      ),
    );

    expect(lineXml(hidden)).toBe('<a:ln w="9525" cmpd="sng"><a:noFill/></a:ln>');
    expect(lineXml(shown)).toContain('<a:schemeClr val="lt1"><a:shade val="50000"/>');
  });

  test('writes italic, underlined text in a half point size', () => {
    const drawing = drawingXml(
      Shape.textbox()
        .setText('Note')
        .setFont(new ShapeFont().setItalic().setUnderline().setSize(10.5)),
    );

    expect(drawing).toContain('<a:rPr lang="en-US" sz="1050" i="1" u="sng"/>');
  });

  test('writes the font of every line of the text', () => {
    const drawing = drawingXml(
      Shape.textbox().setText('First\nSecond').setFont(new ShapeFont().setSize(6)),
    );

    expect(drawing.match(/<a:rPr lang="en-US" sz="600"\/>/g)).toHaveLength(2);
  });

  test.each([
    ['default', '<a:p><a:r>'],
    ['left', '<a:p><a:pPr algn="l"/><a:r>'],
    ['center', '<a:p><a:pPr algn="ctr"/><a:r>'],
    ['right', '<a:p><a:pPr algn="r"/><a:r>'],
  ])('writes the %s horizontal alignment', (alignment, paragraph) => {
    const options = new ShapeText().setHorizontalAlignment(alignment);
    const drawing = drawingXml(Shape.textbox().setText('Aligned').setTextOptions(options));

    expect(drawing).toContain(paragraph);
  });

  test.each([
    ['top', 'anchor="t"'],
    ['middle', 'anchor="ctr" anchorCtr="0"'],
    ['bottom', 'anchor="b" anchorCtr="0"'],
    ['topCentered', 'anchor="t" anchorCtr="1"'],
    ['middleCentered', 'anchor="ctr" anchorCtr="1"'],
    ['bottomCentered', 'anchor="b" anchorCtr="1"'],
  ])('writes the %s vertical alignment as %s', (alignment, anchor) => {
    const options = new ShapeText().setVerticalAlignment(alignment);
    const drawing = drawingXml(Shape.textbox().setText('Aligned').setTextOptions(options));

    expect(drawing).toContain(`<a:bodyPr wrap="square" rtlCol="0" ${anchor}/>`);
  });

  test.each([
    ['horizontal', '<a:bodyPr wrap="square"'],
    ['rotate90', '<a:bodyPr vert="vert" wrap="square"'],
    ['rotate270', '<a:bodyPr vert="vert270" wrap="square"'],
    ['stacked', '<a:bodyPr vert="wordArtVert" wrap="square"'],
    ['rotate90EastAsian', '<a:bodyPr vert="eaVert" wrap="square"'],
  ])('writes the %s text direction', (direction, bodyProperties) => {
    const options = new ShapeText().setDirection(direction);
    const drawing = drawingXml(Shape.textbox().setText('Turned').setTextOptions(options));

    expect(drawing).toContain(bodyProperties);
  });

  test('keeps a copy of the format, font and text options', () => {
    const fill = new ShapeSolidFill().setColor('yellow');
    const format = new ShapeFormat().setSolidFill(fill);
    const font = new ShapeFont().setSize(8);
    const options = new ShapeText().setDirection('rotate270');
    const shape = Shape.textbox()
      .setText('Copy')
      .setFormat(format)
      .setFont(font)
      .setTextOptions(options);

    fill.setColor('red');
    format.setSolidFill(fill);
    font.setSize(20);
    options.setDirection('horizontal');
    const drawing = drawingXml(shape);

    expect(shapeProperties(drawing)).toContain('<a:srgbClr val="FFFF00"/>');
    expect(drawing).toContain('sz="800"');
    expect(drawing).toContain('vert="vert270"');
  });

  test('formats many textboxes from shared settings', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();
    const format = new ShapeFormat()
      .setPatternFill(
        new ShapePatternFill().setPattern('dotted20Percent').setForegroundColor('gray'),
      )
      .setLine(new ShapeLine().setColor('black'));
    const font = new ShapeFont().setName('Arial').setSize(7);
    for (let col = 0; col < 280; col++) {
      worksheet.insertShape(
        0,
        col,
        Shape.textbox().setText(`Column ${col}`).setFormat(format).setFont(font),
      );
    }

    const drawing = readXlsxPart(workbook.saveToBuffer(), 'xl/drawings/drawing1.xml');
    expect(drawing.match(/<xdr:sp macro=""/g)).toHaveLength(280);
    expect(drawing.match(/<a:pattFill prst="pct20">/g)).toHaveLength(280);
    expect(drawing).toContain('<a:t>Column 279</a:t>');
  });
});

describe('Shape formatting validation', () => {
  test.each([
    ['solid fill', (color) => new ShapeSolidFill().setColor(color)],
    ['pattern foreground', (color) => new ShapePatternFill().setForegroundColor(color)],
    ['pattern background', (color) => new ShapePatternFill().setBackgroundColor(color)],
    ['line', (color) => new ShapeLine().setColor(color)],
    ['font', (color) => new ShapeFont().setColor(color)],
  ])('rejects an invalid %s color', (_target, setColor) => {
    expect(() => setColor('#12345')).toThrow('Invalid color: #12345');
    expect(() => setColor('theme:4:6')).toThrow(
      "Invalid color: theme:4:6. Theme colors are 'theme:<color>:<shade>' with a color of 0-9 and a shade of 0-5",
    );
  });

  test.each([101, -1, 12.5, Number.NaN])('rejects a transparency of %s', (transparency) => {
    const message = `Transparency must be an integer from 0 to 100, got ${transparency}`;
    expect(() => new ShapeSolidFill().setTransparency(transparency)).toThrow(message);
    expect(() => new ShapeLine().setTransparency(transparency)).toThrow(message);
  });

  test.each([-0.5, 1584.25, Number.NaN])('rejects a line width of %s', (width) => {
    expect(() => new ShapeLine().setWidth(width)).toThrow(
      `Shape line width must be between 0 and 1584 points, got ${width}`,
    );
  });

  test.each([0.5, 8.25, 4000.5, Number.NaN])('rejects a font size of %s', (size) => {
    expect(() => new ShapeFont().setSize(size)).toThrow(
      `Shape font size must be between 1 and 4000 points in steps of 0.5, got ${size}`,
    );
  });

  test('rejects unknown option names', () => {
    expect(() => new ShapePatternFill().setPattern('dotted15Percent')).toThrow(
      'Invalid shape pattern fill type: dotted15Percent',
    );
    expect(() => new ShapeLine().setDashType('dotted')).toThrow(
      "Invalid shape line dash type: dotted. Use 'solid', 'roundDot', 'squareDot', 'dash', 'dashDot', 'longDash', 'longDashDot' or 'longDashDotDot'",
    );
    expect(() => new ShapeText().setHorizontalAlignment('justify')).toThrow(
      "Invalid shape text horizontal alignment: justify. Use 'default', 'left', 'center' or 'right'",
    );
    expect(() => new ShapeText().setVerticalAlignment('center')).toThrow(
      "Invalid shape text vertical alignment: center. Use 'top', 'middle', 'bottom', 'topCentered', 'middleCentered' or 'bottomCentered'",
    );
    expect(() => new ShapeText().setDirection('rotate180')).toThrow(
      "Invalid shape text direction: rotate180. Use 'horizontal', 'rotate90', 'rotate270', 'stacked' or 'rotate90EastAsian'",
    );
  });
});
