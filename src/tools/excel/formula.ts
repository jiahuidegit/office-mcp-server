/**
 * excel_add_formula - 添加公式
 */

import { excelAddFormulaSchema } from '../../schemas/index.js'
import { workbooks } from './store.js'

export function excelAddFormula(args: Record<string, unknown>) {
  const { workbookId, sheetName, cell, formula } = excelAddFormulaSchema.parse(args)

  const entry = workbooks.get(workbookId)
  if (!entry) {
    return { success: false, message: `工作簿 ${workbookId} 不存在` }
  }

  const sheet = entry.workbook.getWorksheet(sheetName)
  if (!sheet) {
    return { success: false, message: `工作表 ${sheetName} 不存在` }
  }

  const cleanFormula = formula.startsWith('=') ? formula.slice(1) : formula
  const targetCell = sheet.getCell(cell)
  targetCell.value = { formula: cleanFormula }

  return { success: true, message: `已在 ${cell} 添加公式: =${cleanFormula}` }
}
