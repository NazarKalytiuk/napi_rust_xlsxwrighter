/**
 * Rich Text example - demonstrates multi-format text within a single cell
 */

const { Workbook, Format, RichText } = require('@nazarkalytiuk/rusc-xlsx');

const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Rich Text');

worksheet.setColumnWidth(0, 50);

// --- Basic rich text: bold + normal ---
const rt1 = new RichText();
const boldFmt = new Format().setBold();
const normalFmt = new Format();

rt1.addSegment('Important: ', boldFmt);
rt1.addSegment('This text mixes bold and normal styles.', normalFmt);
worksheet.writeRichString(0, 0, rt1);

// --- Colored segments ---
const rt2 = new RichText();
const redFmt = new Format().setFontColor('red').setBold();
const greenFmt = new Format().setFontColor('green');
const blueFmt = new Format().setFontColor('blue').setItalic();

rt2.addSegment('Red ', redFmt);
rt2.addSegment('Green ', greenFmt);
rt2.addSegment('Blue', blueFmt);
worksheet.writeRichString(2, 0, rt2);

// --- Different font sizes ---
const rt3 = new RichText();
const smallFmt = new Format().setFontSize(8);
const mediumFmt = new Format().setFontSize(12);
const largeFmt = new Format().setFontSize(18).setBold();

rt3.addSegment('small ', smallFmt);
rt3.addSegment('medium ', mediumFmt);
rt3.addSegment('LARGE', largeFmt);
worksheet.setRowHeight(4, 25);
worksheet.writeRichString(4, 0, rt3);

// --- Superscript and subscript ---
const rt4 = new RichText();
const baseFmt = new Format().setFontSize(12);
const superFmt = new Format().setFontScript('superscript').setFontSize(8);
const subFmt = new Format().setFontScript('subscript').setFontSize(8);

rt4.addSegment('E = mc', baseFmt);
rt4.addSegment('2', superFmt);
rt4.addSegment('  H', baseFmt);
rt4.addSegment('2', subFmt);
rt4.addSegment('O', baseFmt);
worksheet.writeRichString(6, 0, rt4);

// --- Styled rich text with cell format ---
const rt5 = new RichText();
const headerFmt = new Format().setBold().setFontSize(14).setFontColor('#1F4E78');
const detailFmt = new Format().setFontSize(10).setFontColor('#555555');

rt5.addSegment('Product Launch', headerFmt);
rt5.addSegment(' - Scheduled for next quarter', detailFmt);

const cellFmt = new Format().setBackgroundColor('#E2EFDA').setBorder('thin');
worksheet.setRowHeight(8, 25);
worksheet.writeRichStringWithFormat(8, 0, rt5, cellFmt);

// --- Underline and strikethrough mix ---
const rt6 = new RichText();
const underlineFmt = new Format().setUnderline('single');
const strikeFmt = new Format().setStrikethrough();

rt6.addSegment('underlined', underlineFmt);
rt6.addSegment(' and ', normalFmt);
rt6.addSegment('strikethrough', strikeFmt);
rt6.addSegment(' in one cell', normalFmt);
worksheet.writeRichString(10, 0, rt6);

workbook.save('rich-text.xlsx');
console.log('Created rich-text.xlsx with 6 rich text examples');
