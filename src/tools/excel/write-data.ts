/**
 * excel_write_data - 写入表格数据
 */

import ExcelJS from 'exceljs'
import { excelWriteDataSchema } from '../../schemas/index.js'
import { workbooks } from './store.js'

export function excelWriteData(args: Record<string, unknown>) {
  const {
    workbookId, sheetName, headers, rows,
    startCell, autoFilter, freezeHeader, style
  } = excelWriteDataSchema.parse(args)

  const entry = workbooks.get(workbookId)
  if (!entry) {
    return { success: false, message: `工作簿 ${workbookId} 不存在` }
  }

  let sheet = entry.workbook.getWorksheet(sheetName)
  if (!sheet) {
    sheet = entry.workbook.addWorksheet(sheetName)
  }

  const themeStyle = entry.theme.excel

  // 解析起始单元格
  const startMatch = startCell.match(/^([A-Z]+)(\d+)$/)
  const startCol = startMatch ? startMatch[1].charCodeAt(0) - 64 : 1
  const startRow = startMatch ? parseInt(startMatch[2]) : 1

  // 写入表头
  const headerRow = sheet.getRow(startRow)
  headers.forEach((header, index) => {
    const cell = headerRow.getCell(startCol + index)
    cell.value = header
    cell.font = {
      bold: themeStyle.headerRow.bold,
      color: { argb: `FF${themeStyle.headerRow.fontColor}` }
    }
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: `FF${themeStyle.headerRow.fill}` }
    }
    cell.alignment = { horizontal: 'center', vertical: 'middle' }
    if (style !== 'minimal') {
      cell.border = {
        top: { style: 'thin', color: { argb: `FF${themeStyle.border.color}` } },
        bottom: { style: 'thin', color: { argb: `FF${themeStyle.border.color}` } },
        left: { style: 'thin', color: { argb: `FF${themeStyle.border.color}` } },
        right: { style: 'thin', color: { argb: `FF${themeStyle.border.color}` } }
      }
    }
  })
  headerRow.height = 25

  // 写入数据行
  rows.forEach((row, rowIndex) => {
    const dataRow = sheet.getRow(startRow + 1 + rowIndex)
    row.forEach((value, colIndex) => {
      const cell = dataRow.getCell(startCol + colIndex)
      cell.value = value as ExcelJS.CellValue

      if (style === 'striped' && rowIndex % 2 === 1) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: `FF${themeStyle.alternateRow.fill}` }
        }
      }

      if (style !== 'minimal') {
        cell.border = {
          top: { style: 'thin', color: { argb: `FF${themeStyle.border.color}` } },
          bottom: { style: 'thin', color: { argb: `FF${themeStyle.border.color}` } },
          left: { style: 'thin', color: { argb: `FF${themeStyle.border.color}` } },
          right: { style: 'thin', color: { argb: `FF${themeStyle.border.color}` } }
        }
      }
    })
  })

  // 自动调整列宽
  headers.forEach((_, index) => {
    const column = sheet.getColumn(startCol + index)
    column.width = 15
  })

  if (freezeHeader) {
    sheet.views = [{ state: 'frozen', ySplit: startRow }]
  }

  if (autoFilter) {
    const endCol = String.fromCharCode(64 + startCol + headers.length - 1)
    const endRow = startRow + rows.length
    sheet.autoFilter = `${startCell}:${endCol}${endRow}`
  }

  return {
    success: true,
    message: `已写入 ${headers.length} 列 ${rows.length} 行数据`
  }
}
