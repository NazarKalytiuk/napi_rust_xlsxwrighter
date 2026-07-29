// JavaScript wrappers for rusc-xlsx
const nativeBinding = require('./index.js');

// JavaScript wrapper for Format with fluent API
class Format {
  constructor() {
    this._native = new nativeBinding.RuscFormat()
  }

  setBold() {
    this._native.setBold()
    return this
  }

  setItalic() {
    this._native.setItalic()
    return this
  }

  setFontSize(size) {
    this._native.setFontSize(size)
    return this
  }

  setFontColor(color) {
    this._native.setFontColor(color)
    return this
  }

  setBackgroundColor(color) {
    this._native.setBackgroundColor(color)
    return this
  }

  setAlign(align) {
    this._native.setAlign(align)
    return this
  }

  setVerticalAlign(align) {
    this._native.setVerticalAlign(align)
    return this
  }

  setBorder(border) {
    this._native.setBorder(border)
    return this
  }

  setNumFormat(format) {
    this._native.setNumFormat(format)
    return this
  }

  setFontName(fontName) {
    this._native.setFontName(fontName)
    return this
  }

  setUnderline(underline) {
    this._native.setUnderline(underline)
    return this
  }

  setStrikethrough() {
    this._native.setStrikethrough()
    return this
  }

  setForegroundColor(color) {
    this._native.setForegroundColor(color)
    return this
  }

  setPattern(pattern) {
    this._native.setPattern(pattern)
    return this
  }

  setTextWrap() {
    this._native.setTextWrap()
    return this
  }

  setIndent(indent) {
    this._native.setIndent(indent)
    return this
  }

  setRotation(rotation) {
    this._native.setRotation(rotation)
    return this
  }

  setShrink() {
    this._native.setShrink()
    return this
  }

  setBorderTop(border) {
    this._native.setBorderTop(border)
    return this
  }

  setBorderBottom(border) {
    this._native.setBorderBottom(border)
    return this
  }

  setBorderLeft(border) {
    this._native.setBorderLeft(border)
    return this
  }

  setBorderRight(border) {
    this._native.setBorderRight(border)
    return this
  }

  setBorderColor(color) {
    this._native.setBorderColor(color)
    return this
  }

  setBorderTopColor(color) {
    this._native.setBorderTopColor(color)
    return this
  }

  setBorderBottomColor(color) {
    this._native.setBorderBottomColor(color)
    return this
  }

  setBorderLeftColor(color) {
    this._native.setBorderLeftColor(color)
    return this
  }

  setBorderRightColor(color) {
    this._native.setBorderRightColor(color)
    return this
  }

  setUnlocked() {
    this._native.setUnlocked()
    return this
  }

  setLocked() {
    this._native.setLocked()
    return this
  }

  setHidden() {
    this._native.setHidden()
    return this
  }

  setFontScript(script) {
    this._native.setFontScript(script)
    return this
  }

  setBorderDiagonal(border) {
    this._native.setBorderDiagonal(border)
    return this
  }

  setBorderDiagonalColor(color) {
    this._native.setBorderDiagonalColor(color)
    return this
  }

  setBorderDiagonalType(type) {
    this._native.setBorderDiagonalType(type)
    return this
  }

  setReadingDirection(direction) {
    this._native.setReadingDirection(direction)
    return this
  }

  setQuotePrefix() {
    this._native.setQuotePrefix()
    return this
  }
}

// JavaScript wrapper for Note
class Note {
  constructor(text) {
    this._native = new nativeBinding.RuscNote(text)
  }

  setAuthor(name) {
    this._native.setAuthor(name)
    return this
  }

  setWidth(width) {
    this._native.setWidth(width)
    return this
  }

  setHeight(height) {
    this._native.setHeight(height)
    return this
  }

  setBackgroundColor(color) {
    this._native.setBackgroundColor(color)
    return this
  }

  setVisible(visible) {
    this._native.setVisible(visible)
    return this
  }

  setAltText(altText) {
    this._native.setAltText(altText)
    return this
  }
}

// JavaScript wrapper for ProtectionOptions
class ProtectionOptions {
  constructor() {
    this._native = new nativeBinding.RuscProtectionOptions()
  }

  setSelectLockedCells(allow) {
    this._native.setSelectLockedCells(allow)
    return this
  }

  setSelectUnlockedCells(allow) {
    this._native.setSelectUnlockedCells(allow)
    return this
  }

  setFormatCells(allow) {
    this._native.setFormatCells(allow)
    return this
  }

  setFormatColumns(allow) {
    this._native.setFormatColumns(allow)
    return this
  }

  setFormatRows(allow) {
    this._native.setFormatRows(allow)
    return this
  }

  setInsertColumns(allow) {
    this._native.setInsertColumns(allow)
    return this
  }

  setInsertRows(allow) {
    this._native.setInsertRows(allow)
    return this
  }

  setInsertLinks(allow) {
    this._native.setInsertLinks(allow)
    return this
  }

  setDeleteColumns(allow) {
    this._native.setDeleteColumns(allow)
    return this
  }

  setDeleteRows(allow) {
    this._native.setDeleteRows(allow)
    return this
  }

  setSort(allow) {
    this._native.setSort(allow)
    return this
  }

  setUseAutofilter(allow) {
    this._native.setUseAutofilter(allow)
    return this
  }

  setUsePivotTables(allow) {
    this._native.setUsePivotTables(allow)
    return this
  }

  setEditScenarios(allow) {
    this._native.setEditScenarios(allow)
    return this
  }

  setEditObjects(allow) {
    this._native.setEditObjects(allow)
    return this
  }

  setContents(allow) {
    this._native.setContents(allow)
    return this
  }
}

// JavaScript wrapper for RichText
class RichText {
  constructor() {
    this._native = new nativeBinding.RuscRichText()
  }

  addSegment(text, format) {
    this._native.addSegment(text, format ? format._native : null)
    return this
  }
}

// JavaScript wrapper for Image
class Image {
  constructor(path) {
    this._native = nativeBinding.RuscImage.new(path)
  }

  static fromBuffer(buffer) {
    const img = Object.create(Image.prototype)
    img._native = nativeBinding.RuscImage.newFromBuffer(buffer)
    return img
  }

  setWidth(width) {
    this._native.setWidth(width)
    return this
  }

  setHeight(height) {
    this._native.setHeight(height)
    return this
  }

  setScaleWidth(scale) {
    this._native.setScaleWidth(scale)
    return this
  }

  setScaleHeight(scale) {
    this._native.setScaleHeight(scale)
    return this
  }

  setScaleToSize(width, height, keepAspectRatio = true) {
    this._native.setScaleToSize(width, height, keepAspectRatio)
    return this
  }

  setAltText(altText) {
    this._native.setAltText(altText)
    return this
  }

  setDecorative(enable) {
    this._native.setDecorative(enable)
    return this
  }

  setObjectMovement(option) {
    this._native.setObjectMovement(option)
    return this
  }

  setUrl(url) {
    this._native.setUrl(url)
    return this
  }
}

// JavaScript wrapper for Formula
class Formula {
  constructor(formula) {
    this._native = new nativeBinding.RuscFormula(formula)
  }

  setResult(result) {
    this._native.setResult(result)
    return this
  }
}

// JavaScript wrapper for DataValidation
class DataValidation {
  constructor() {
    this._native = new nativeBinding.RuscDataValidation()
  }

  allowWholeNumber(rule, value) {
    this._native.allowWholeNumber(rule, value)
    return this
  }

  allowWholeNumberBetween(min, max) {
    this._native.allowWholeNumberBetween(min, max)
    return this
  }

  allowDecimalNumber(rule, value) {
    this._native.allowDecimalNumber(rule, value)
    return this
  }

  allowDecimalNumberBetween(min, max) {
    this._native.allowDecimalNumberBetween(min, max)
    return this
  }

  allowListStrings(list) {
    this._native.allowListStrings(list)
    return this
  }

  allowListFormula(formula) {
    this._native.allowListFormula(formula)
    return this
  }

  allowCustom(formula) {
    this._native.allowCustom(formula)
    return this
  }

  allowAnyValue() {
    this._native.allowAnyValue()
    return this
  }

  ignoreBlank(enable) {
    this._native.ignoreBlank(enable)
    return this
  }

  showDropdown(enable) {
    this._native.showDropdown(enable)
    return this
  }

  showInputMessage(enable) {
    this._native.showInputMessage(enable)
    return this
  }

  showErrorMessage(enable) {
    this._native.showErrorMessage(enable)
    return this
  }

  setInputTitle(title) {
    this._native.setInputTitle(title)
    return this
  }

  setInputMessage(message) {
    this._native.setInputMessage(message)
    return this
  }

  setErrorTitle(title) {
    this._native.setErrorTitle(title)
    return this
  }

  setErrorMessage(message) {
    this._native.setErrorMessage(message)
    return this
  }

  setErrorStyle(style) {
    this._native.setErrorStyle(style)
    return this
  }
}

// JavaScript wrapper for Worksheet
class Worksheet {
  constructor(nativeWorksheet) {
    this._native = nativeWorksheet
  }

  write(row, col, value, format) {
    return this._native.write(row, col, value, format ? format._native : undefined)
  }

  writeRow(row, startCol, values, format) {
    return this._native.writeRow(row, startCol, values, format ? format._native : undefined)
  }

  writeRows(startRow, startCol, rows, format) {
    return this._native.writeRows(startRow, startCol, rows, format ? format._native : undefined)
  }

  writeString(row, col, value) {
    return this._native.writeString(row, col, value)
  }

  writeNumber(row, col, value) {
    return this._native.writeNumber(row, col, value)
  }

  writeBoolean(row, col, value) {
    return this._native.writeBoolean(row, col, value)
  }

  writeBlank(row, col, format) {
    return this._native.writeBlank(row, col, format._native)
  }

  writeDatetimeFromTimestamp(row, col, timestamp, format) {
    return this._native.writeDatetimeFromTimestamp(row, col, timestamp, format ? format._native : undefined)
  }

  writeDatetimeFromYmd(row, col, year, month, day, format) {
    return this._native.writeDatetimeFromYmd(row, col, year, month, day, format ? format._native : undefined)
  }

  writeDatetimeFromYmdHms(row, col, year, month, day, hour, min, sec, format) {
    return this._native.writeDatetimeFromYmdHms(row, col, year, month, day, hour, min, sec, format ? format._native : undefined)
  }

  writeDatetimeFromString(row, col, datetimeStr, format) {
    return this._native.writeDatetimeFromString(row, col, datetimeStr, format ? format._native : undefined)
  }

  writeDatetime(row, col, datetime, format) {
    // Convenience method that handles multiple input types
    if (datetime instanceof Date) {
      // JavaScript Date object
      const timestamp = Math.floor(datetime.getTime() / 1000)
      return this.writeDatetimeFromTimestamp(row, col, timestamp, format)
    } else if (typeof datetime === 'number') {
      // Unix timestamp
      return this.writeDatetimeFromTimestamp(row, col, datetime, format)
    } else if (typeof datetime === 'string') {
      // ISO string
      return this.writeDatetimeFromString(row, col, datetime, format)
    } else if (typeof datetime === 'object') {
      // Object with year, month, day, etc.
      if (datetime.hour !== undefined || datetime.min !== undefined || datetime.sec !== undefined) {
        return this.writeDatetimeFromYmdHms(
          row, col,
          datetime.year, datetime.month, datetime.day,
          datetime.hour || 0, datetime.min || 0, datetime.sec || 0,
          format
        )
      } else {
        return this.writeDatetimeFromYmd(row, col, datetime.year, datetime.month, datetime.day, format)
      }
    } else {
      throw new Error('Invalid datetime input: must be Date, number (timestamp), string (ISO), or object')
    }
  }

  writeWithFormat(row, col, value, format) {
    return this._native.writeWithFormat(row, col, value, format._native)
  }

  setColumnWidth(col, width) {
    return this._native.setColumnWidth(col, width)
  }

  setRowHeight(row, height) {
    return this._native.setRowHeight(row, height)
  }

  setDefaultRowHeight(height) {
    return this._native.setDefaultRowHeight(height)
  }

  setName(name) {
    return this._native.setName(name)
  }

  mergeRange(firstRow, firstCol, lastRow, lastCol, value, format) {
    return this._native.mergeRange(firstRow, firstCol, lastRow, lastCol, value, format ? format._native : undefined)
  }

  mergeRangeWithFormat(firstRow, firstCol, lastRow, lastCol, value, format) {
    return this._native.mergeRangeWithFormat(firstRow, firstCol, lastRow, lastCol, value, format._native)
  }

  insertNote(row, col, note) {
    return this._native.insertNote(row, col, note._native)
  }

  setDefaultNoteAuthor(name) {
    return this._native.setDefaultNoteAuthor(name)
  }

  writeFormula(row, col, formula, format) {
    return this._native.writeFormula(row, col, formula._native, format ? format._native : undefined)
  }

  writeDynamicFormula(row, col, formula, format) {
    return this._native.writeDynamicFormula(row, col, formula._native, format ? format._native : undefined)
  }

  writeArrayFormula(firstRow, firstCol, lastRow, lastCol, formula, format) {
    return this._native.writeArrayFormula(firstRow, firstCol, lastRow, lastCol, formula._native, format ? format._native : undefined)
  }

  addDataValidation(firstRow, firstCol, lastRow, lastCol, validation) {
    return this._native.addDataValidation(firstRow, firstCol, lastRow, lastCol, validation._native)
  }

  writeUrl(row, col, url, format) {
    return this._native.writeUrl(row, col, url, format ? format._native : undefined)
  }

  writeUrlWithText(row, col, url, text, format) {
    return this._native.writeUrlWithText(row, col, url, text, format ? format._native : undefined)
  }

  autofilter(firstRow, firstCol, lastRow, lastCol) {
    return this._native.autofilter(firstRow, firstCol, lastRow, lastCol)
  }

  setFreezePanes(row, col) {
    return this._native.setFreezePanes(row, col)
  }

  // Conditional Formatting
  addConditionalFormat2ColorScale(firstRow, firstCol, lastRow, lastCol, minColor, maxColor) {
    return this._native.addConditionalFormat2ColorScale(firstRow, firstCol, lastRow, lastCol, minColor, maxColor)
  }

  addConditionalFormat3ColorScale(firstRow, firstCol, lastRow, lastCol, minColor, midColor, maxColor) {
    return this._native.addConditionalFormat3ColorScale(firstRow, firstCol, lastRow, lastCol, minColor, midColor, maxColor)
  }

  addConditionalFormatDataBar(firstRow, firstCol, lastRow, lastCol, barColor) {
    return this._native.addConditionalFormatDataBar(firstRow, firstCol, lastRow, lastCol, barColor)
  }

  addConditionalFormatCell(firstRow, firstCol, lastRow, lastCol, rule, value, format) {
    return this._native.addConditionalFormatCell(firstRow, firstCol, lastRow, lastCol, rule, value, format._native)
  }

  addConditionalFormatText(firstRow, firstCol, lastRow, lastCol, rule, text, format) {
    return this._native.addConditionalFormatText(firstRow, firstCol, lastRow, lastCol, rule, text, format._native)
  }

  addConditionalFormatAverage(firstRow, firstCol, lastRow, lastCol, rule, format) {
    return this._native.addConditionalFormatAverage(firstRow, firstCol, lastRow, lastCol, rule, format._native)
  }

  addConditionalFormatTop(firstRow, firstCol, lastRow, lastCol, rule, value, format) {
    return this._native.addConditionalFormatTop(firstRow, firstCol, lastRow, lastCol, rule, value, format._native)
  }

  addConditionalFormatDuplicate(firstRow, firstCol, lastRow, lastCol, invert, format) {
    return this._native.addConditionalFormatDuplicate(firstRow, firstCol, lastRow, lastCol, invert, format._native)
  }

  addConditionalFormatBlank(firstRow, firstCol, lastRow, lastCol, invert, format) {
    return this._native.addConditionalFormatBlank(firstRow, firstCol, lastRow, lastCol, invert, format._native)
  }

  addConditionalFormatError(firstRow, firstCol, lastRow, lastCol, invert, format) {
    return this._native.addConditionalFormatError(firstRow, firstCol, lastRow, lastCol, invert, format._native)
  }

  addConditionalFormatDate(firstRow, firstCol, lastRow, lastCol, rule, format) {
    return this._native.addConditionalFormatDate(firstRow, firstCol, lastRow, lastCol, rule, format._native)
  }

  addConditionalFormatFormula(firstRow, firstCol, lastRow, lastCol, formula, format) {
    return this._native.addConditionalFormatFormula(firstRow, firstCol, lastRow, lastCol, formula, format._native)
  }

  addConditionalFormatIconSet(firstRow, firstCol, lastRow, lastCol, iconType) {
    return this._native.addConditionalFormatIconSet(firstRow, firstCol, lastRow, lastCol, iconType)
  }

  // Page Setup
  setMargins(left, right, top, bottom, header, footer) {
    return this._native.setMargins(left, right, top, bottom, header, footer)
  }

  setPortrait() {
    return this._native.setPortrait()
  }

  setLandscape() {
    return this._native.setLandscape()
  }

  setPaperSize(paperSize) {
    return this._native.setPaperSize(paperSize)
  }

  setHeader(header) {
    return this._native.setHeader(header)
  }

  setFooter(footer) {
    return this._native.setFooter(footer)
  }

  setPrintArea(firstRow, firstCol, lastRow, lastCol) {
    return this._native.setPrintArea(firstRow, firstCol, lastRow, lastCol)
  }

  setPrintGridlines(enable) {
    return this._native.setPrintGridlines(enable)
  }

  setPrintCenterHorizontally(enable) {
    return this._native.setPrintCenterHorizontally(enable)
  }

  setPrintCenterVertically(enable) {
    return this._native.setPrintCenterVertically(enable)
  }

  setPrintFitToPages(width, height) {
    return this._native.setPrintFitToPages(width, height)
  }

  setPrintScale(scale) {
    return this._native.setPrintScale(scale)
  }

  setPrintFirstPageNumber(pageNumber) {
    return this._native.setPrintFirstPageNumber(pageNumber)
  }

  setPrintBlackAndWhite(enable) {
    return this._native.setPrintBlackAndWhite(enable)
  }

  setPrintDraft(enable) {
    return this._native.setPrintDraft(enable)
  }

  setPrintHeadings(enable) {
    return this._native.setPrintHeadings(enable)
  }

  setRepeatRows(firstRow, lastRow) {
    return this._native.setRepeatRows(firstRow, lastRow)
  }

  setRepeatColumns(firstCol, lastCol) {
    return this._native.setRepeatColumns(firstCol, lastCol)
  }

  setHeaderFooterScaleWithDoc(enable) {
    return this._native.setHeaderFooterScaleWithDoc(enable)
  }

  setHeaderFooterAlignWithPage(enable) {
    return this._native.setHeaderFooterAlignWithPage(enable)
  }

  // Worksheet View & Display
  setTabColor(color) {
    return this._native.setTabColor(color)
  }

  setHidden(enable) {
    return this._native.setHidden(enable)
  }

  setZoom(zoom) {
    return this._native.setZoom(zoom)
  }

  setSelection(firstRow, firstCol, lastRow, lastCol) {
    return this._native.setSelection(firstRow, firstCol, lastRow, lastCol)
  }

  setTopLeftCell(row, col) {
    return this._native.setTopLeftCell(row, col)
  }

  // Worksheet Protection
  protect() {
    return this._native.protect()
  }

  protectWithPassword(password) {
    return this._native.protectWithPassword(password)
  }

  protectWithOptions(options) {
    return this._native.protectWithOptions(options._native)
  }

  protectWithPasswordAndOptions(password, options) {
    return this._native.protectWithPasswordAndOptions(password, options._native)
  }

  // Rich Text
  writeRichString(row, col, richText) {
    return this._native.writeRichString(row, col, richText._native)
  }

  writeRichStringWithFormat(row, col, richText, format) {
    return this._native.writeRichStringWithFormat(row, col, richText._native, format._native)
  }

  // Charts
  insertChart(row, col, chart) {
    return this._native.insertChart(row, col, chart._native)
  }

  insertChartWithOffset(row, col, chart, xOffset, yOffset) {
    return this._native.insertChartWithOffset(row, col, chart._native, xOffset, yOffset)
  }

  // Images
  insertImage(row, col, image) {
    return this._native.insertImage(row, col, image._native)
  }

  insertImageWithOffset(row, col, image, xOffset, yOffset) {
    return this._native.insertImageWithOffset(row, col, image._native, xOffset, yOffset)
  }

  insertImageFitToCell(row, col, image, keepAspectRatio = true) {
    return this._native.insertImageFitToCell(row, col, image._native, keepAspectRatio)
  }

  addTable(firstRow, firstCol, lastRow, lastCol, table) {
    return this._native.addTable(firstRow, firstCol, lastRow, lastCol, table._native)
  }

  addSparkline(row, col, sparkline) {
    return this._native.addSparkline(row, col, sparkline._native)
  }

  addSparklineGroup(firstRow, firstCol, lastRow, lastCol, sparkline) {
    return this._native.addSparklineGroup(firstRow, firstCol, lastRow, lastCol, sparkline._native)
  }

  insertButton(row, col, button) {
    return this._native.insertButton(row, col, button._native)
  }

  insertButtonWithOffset(row, col, button, xOffset, yOffset) {
    return this._native.insertButtonWithOffset(row, col, button._native, xOffset, yOffset)
  }

  insertShape(row, col, shape) {
    return this._native.insertShape(row, col, shape._native)
  }

  insertShapeWithOffset(row, col, shape, xOffset, yOffset) {
    return this._native.insertShapeWithOffset(row, col, shape._native, xOffset, yOffset)
  }
}

// JavaScript wrapper for Chart
class Chart {
  constructor(chartType) {
    this._native = new nativeBinding.RuscChart(chartType)
  }

  addSeries(options) {
    const {categories, values, name} = options
    return this._native.addSeries(categories || null, values, name || null)
  }

  setTitle(title) {
    this._native.setTitle(title)
    return this
  }

  setXAxisName(name) {
    this._native.setXAxisName(name)
    return this
  }

  setYAxisName(name) {
    this._native.setYAxisName(name)
    return this
  }

  setStyle(style) {
    this._native.setStyle(style)
    return this
  }

  setLegendHidden() {
    this._native.setLegendHidden()
    return this
  }
}

// JavaScript wrapper for Table
class Table {
  constructor() {
    this._native = nativeBinding.RuscTable.new()
  }

  setHeaderRow(show) {
    this._native.setHeaderRow(show)
    return this
  }

  setTotalRow(show) {
    this._native.setTotalRow(show)
    return this
  }

  setBandedRows(show) {
    this._native.setBandedRows(show)
    return this
  }

  setBandedColumns(show) {
    this._native.setBandedColumns(show)
    return this
  }

  setFirstColumn(show) {
    this._native.setFirstColumn(show)
    return this
  }

  setLastColumn(show) {
    this._native.setLastColumn(show)
    return this
  }

  setName(name) {
    this._native.setName(name)
    return this
  }

  setStyle(styleId) {
    this._native.setStyle(styleId)
    return this
  }

  setColumns(columns) {
    const nativeColumns = columns.map(col => col._native)
    this._native.setColumns(nativeColumns)
    return this
  }
}

// JavaScript wrapper for TableColumn
class TableColumn {
  constructor() {
    this._native = nativeBinding.RuscTableColumn.new()
  }

  setHeader(header) {
    this._native.setHeader(header)
    return this
  }

  setTotalFunction(func) {
    this._native.setTotalFunction(func)
    return this
  }

  setTotalLabel(label) {
    this._native.setTotalLabel(label)
    return this
  }

  setFormula(formula) {
    this._native.setFormula(formula)
    return this
  }
}

// JavaScript wrapper for Sparkline
class Sparkline {
  constructor() {
    this._native = nativeBinding.RuscSparkline.new()
  }

  setRange(range) {
    this._native.setRange(range)
    return this
  }

  setType(type) {
    this._native.setType(type)
    return this
  }

  showHighPoint(enable) {
    this._native.showHighPoint(enable)
    return this
  }

  showLowPoint(enable) {
    this._native.showLowPoint(enable)
    return this
  }

  showFirstPoint(enable) {
    this._native.showFirstPoint(enable)
    return this
  }

  showLastPoint(enable) {
    this._native.showLastPoint(enable)
    return this
  }

  showNegativePoints(enable) {
    this._native.showNegativePoints(enable)
    return this
  }

  showMarkers(enable) {
    this._native.showMarkers(enable)
    return this
  }

  showAxis(enable) {
    this._native.showAxis(enable)
    return this
  }

  showHiddenData(enable) {
    this._native.showHiddenData(enable)
    return this
  }

  setRightToLeft(enable) {
    this._native.setRightToLeft(enable)
    return this
  }

  setSparklineColor(color) {
    this._native.setSparklineColor(color)
    return this
  }

  setHighPointColor(color) {
    this._native.setHighPointColor(color)
    return this
  }

  setLowPointColor(color) {
    this._native.setLowPointColor(color)
    return this
  }

  setFirstPointColor(color) {
    this._native.setFirstPointColor(color)
    return this
  }

  setLastPointColor(color) {
    this._native.setLastPointColor(color)
    return this
  }

  setNegativePointsColor(color) {
    this._native.setNegativePointsColor(color)
    return this
  }

  setMarkersColor(color) {
    this._native.setMarkersColor(color)
    return this
  }

  setLineWeight(weight) {
    this._native.setLineWeight(weight)
    return this
  }

  setCustomMax(max) {
    this._native.setCustomMax(max)
    return this
  }

  setCustomMin(min) {
    this._native.setCustomMin(min)
    return this
  }

  setGroupMax(enable) {
    this._native.setGroupMax(enable)
    return this
  }

  setGroupMin(enable) {
    this._native.setGroupMin(enable)
    return this
  }

  setStyle(style) {
    this._native.setStyle(style)
    return this
  }

  showEmptyCellsAs(option) {
    this._native.showEmptyCellsAs(option)
    return this
  }
}

// JavaScript wrapper for DocProperties
class DocProperties {
  constructor() {
    this._native = nativeBinding.RuscDocProperties.new()
  }

  setAuthor(author) {
    this._native.setAuthor(author)
    return this
  }

  setTitle(title) {
    this._native.setTitle(title)
    return this
  }

  setSubject(subject) {
    this._native.setSubject(subject)
    return this
  }

  setComment(comment) {
    this._native.setComment(comment)
    return this
  }

  setCompany(company) {
    this._native.setCompany(company)
    return this
  }

  setManager(manager) {
    this._native.setManager(manager)
    return this
  }

  setStatus(status) {
    this._native.setStatus(status)
    return this
  }

  setCategory(category) {
    this._native.setCategory(category)
    return this
  }

  setKeywords(keywords) {
    this._native.setKeywords(keywords)
    return this
  }

  setHyperlinkBase(base) {
    this._native.setHyperlinkBase(base)
    return this
  }
}

// JavaScript wrapper for Button
class Button {
  constructor() {
    this._native = nativeBinding.RuscButton.new()
  }

  setCaption(caption) {
    this._native.setCaption(caption)
    return this
  }

  setMacro(macroName) {
    this._native.setMacro(macroName)
    return this
  }

  setWidth(width) {
    this._native.setWidth(width)
    return this
  }

  setHeight(height) {
    this._native.setHeight(height)
    return this
  }

  setAltText(altText) {
    this._native.setAltText(altText)
    return this
  }

  setObjectMovement(option) {
    this._native.setObjectMovement(option)
    return this
  }
}

// JavaScript wrapper for Shape
class Shape {
  static textbox() {
    const shape = Object.create(Shape.prototype)
    shape._native = nativeBinding.RuscShape.textbox()
    return shape
  }

  setText(text) {
    this._native.setText(text)
    return this
  }

  setWidth(width) {
    this._native.setWidth(width)
    return this
  }

  setHeight(height) {
    this._native.setHeight(height)
    return this
  }

  setAltText(altText) {
    this._native.setAltText(altText)
    return this
  }

  setObjectMovement(option) {
    this._native.setObjectMovement(option)
    return this
  }
}

// JavaScript wrapper for Workbook
class Workbook {
  constructor() {
    this._native = new nativeBinding.RuscWorkbook()
  }

  addWorksheet(name) {
    const nativeWorksheet = this._native.addWorksheet(name)
    return new Worksheet(nativeWorksheet)
  }

  addWorksheetWithConstantMemory(name) {
    const nativeWorksheet = this._native.addWorksheetWithConstantMemory(name)
    return new Worksheet(nativeWorksheet)
  }

  addWorksheetWithLowMemory(name) {
    const nativeWorksheet = this._native.addWorksheetWithLowMemory(name)
    return new Worksheet(nativeWorksheet)
  }

  setTempdir(dir) {
    return this._native.setTempdir(dir)
  }

  save(filename) {
    return this._native.save(filename)
  }

  saveToBuffer() {
    return this._native.saveToBuffer()
  }

  async saveToStream(writable) {
    const os = require('os')
    const path = require('path')
    const fs = require('fs')
    const crypto = require('crypto')
    const { pipeline } = require('stream/promises')

    const tmpFile = path.join(os.tmpdir(), `rusc-xlsx-${crypto.randomBytes(8).toString('hex')}.xlsx`)
    let operationError

    try {
      this._native.save(tmpFile)
      await pipeline(fs.createReadStream(tmpFile), writable)
    } catch (error) {
      operationError = error
    }

    try {
      await fs.promises.unlink(tmpFile)
    } catch (cleanupError) {
      if (cleanupError.code !== 'ENOENT') {
        if (operationError) {
          throw new AggregateError(
            [operationError, cleanupError],
            'Failed to save workbook to stream and remove the temporary file'
          )
        }
        throw cleanupError
      }
    }

    if (operationError) {
      throw operationError
    }
  }

  defineName(name, formula) {
    return this._native.defineName(name, formula)
  }

  readOnlyRecommended() {
    return this._native.readOnlyRecommended()
  }

  addVbaProject(path) {
    return this._native.addVbaProject(path)
  }

  setProperties(properties) {
    return this._native.setProperties(properties._native)
  }
}

// Export wrapper classes
module.exports = {
  Format,
  Note,
  Formula,
  DataValidation,
  Chart,
  Workbook,
  Worksheet,
  ProtectionOptions,
  RichText,
  Image,
  Table,
  TableColumn,
  Sparkline,
  DocProperties,
  Button,
  Shape,
  // Also export native bindings for advanced use
  RuscFormat: nativeBinding.RuscFormat,
  RuscNote: nativeBinding.RuscNote,
  RuscFormula: nativeBinding.RuscFormula,
  RuscDataValidation: nativeBinding.RuscDataValidation,
  RuscChart: nativeBinding.RuscChart,
  RuscWorkbook: nativeBinding.RuscWorkbook,
  RuscWorksheet: nativeBinding.RuscWorksheet,
  RuscProtectionOptions: nativeBinding.RuscProtectionOptions,
  RuscRichText: nativeBinding.RuscRichText,
  RuscImage: nativeBinding.RuscImage,
  RuscTable: nativeBinding.RuscTable,
  RuscTableColumn: nativeBinding.RuscTableColumn,
  RuscSparkline: nativeBinding.RuscSparkline,
  RuscDocProperties: nativeBinding.RuscDocProperties,
  RuscButton: nativeBinding.RuscButton,
  RuscShape: nativeBinding.RuscShape,
  // Memory tracking
  rustMemoryStats: nativeBinding.rustMemoryStats,
  rustMemoryResetPeak: nativeBinding.rustMemoryResetPeak,
}
