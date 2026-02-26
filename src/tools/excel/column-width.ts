/**
 * excel_set_column_width - 设置列宽
 */

import { excelSetColumnWidthSchema } from '../../schemas/index.js'
import { workbooks } from './store.js'

export function excelSetColumnWidth(args: Record<string, unknown>) {
  const { workbookId, sheetName, columns } = excelSetColumnWidthSchema.parse(args)

  const entry = workbooks.get(workbookId)
  if (!entry) {
    return { success: false, message: `工作簿 ${workbookId} 不存在` }
  }

  const sheet = entry.workbook.getWorksheet(sheetName)
  if (!sheet) {
    return { success: false, message: `工作表 ${sheetName} 不存在` }
  }

  Object.entries(columns).forEach(([col, width]) => {
    const column = sheet.getColumn(col)
    column.width = width
  })

  return { success: true, message: `已设置 ${Object.keys(columns).length} 列的宽度` }
}
