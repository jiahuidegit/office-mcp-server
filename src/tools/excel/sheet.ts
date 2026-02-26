/**
 * excel_add_sheet - 添加工作表
 */

import { excelAddSheetSchema } from '../../schemas/index.js'
import { workbooks } from './store.js'

export function excelAddSheet(args: Record<string, unknown>) {
  const { workbookId, sheetName, tabColor } = excelAddSheetSchema.parse(args)

  const entry = workbooks.get(workbookId)
  if (!entry) {
    return { success: false, message: `工作簿 ${workbookId} 不存在` }
  }

  const sheet = entry.workbook.addWorksheet(sheetName)

  if (tabColor) {
    sheet.properties.tabColor = { argb: `FF${tabColor}` }
  }

  return { success: true, message: `已添加工作表: ${sheetName}` }
}
