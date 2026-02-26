/**
 * excel_merge_cells - 合并单元格
 */

import { excelMergeCellsSchema } from '../../schemas/index.js'
import { workbooks } from './store.js'

export function excelMergeCells(args: Record<string, unknown>) {
  const { workbookId, sheetName, range } = excelMergeCellsSchema.parse(args)

  const entry = workbooks.get(workbookId)
  if (!entry) {
    return { success: false, message: `工作簿 ${workbookId} 不存在` }
  }

  const sheet = entry.workbook.getWorksheet(sheetName)
  if (!sheet) {
    return { success: false, message: `工作表 ${sheetName} 不存在` }
  }

  sheet.mergeCells(range)

  return { success: true, message: `已合并单元格: ${range}` }
}
