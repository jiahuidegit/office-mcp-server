/**
 * excel_add_chart - 添加图表数据
 */

import { excelAddChartSchema } from '../../schemas/index.js'
import { workbooks } from './store.js'

export function excelAddChart(args: Record<string, unknown>) {
  const { workbookId, sheetName, chartType, title, categories, series, position } =
    excelAddChartSchema.parse(args)

  const entry = workbooks.get(workbookId)
  if (!entry) {
    return { success: false, message: `工作簿 ${workbookId} 不存在` }
  }

  const sheet = entry.workbook.getWorksheet(sheetName)
  if (!sheet) {
    return { success: false, message: `工作表 ${sheetName} 不存在` }
  }

  // 写入图表数据到指定位置
  const posMatch = position.match(/^([A-Z]+)(\d+)$/)
  const posCol = posMatch ? posMatch[1].charCodeAt(0) - 64 : 6
  const posRow = posMatch ? parseInt(posMatch[2]) : 2

  if (title) {
    const titleCell = sheet.getCell(posRow, posCol)
    titleCell.value = title
    titleCell.font = { bold: true, size: 14 }
  }

  const dataStartRow = posRow + (title ? 2 : 1)

  // 表头行
  const headerRow = sheet.getRow(dataStartRow)
  headerRow.getCell(posCol).value = '分类'
  series.forEach((s, i) => {
    headerRow.getCell(posCol + 1 + i).value = s.name
  })

  // 数据行
  categories.forEach((cat, catIndex) => {
    const row = sheet.getRow(dataStartRow + 1 + catIndex)
    row.getCell(posCol).value = cat
    series.forEach((s, seriesIndex) => {
      row.getCell(posCol + 1 + seriesIndex).value = s.values[catIndex]
    })
  })

  return {
    success: true,
    message: `图表数据已写入（${chartType} 类型），请在 Excel 中选择数据创建图表`,
    note: 'exceljs 不支持直接创建图表，数据已准备好供手动创建图表使用'
  }
}
