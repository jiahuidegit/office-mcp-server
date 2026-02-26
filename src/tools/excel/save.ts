/**
 * excel_save - 保存工作簿
 */

import * as fs from 'fs'
import * as path from 'path'
import { excelSaveSchema } from '../../schemas/index.js'
import { workbooks } from './store.js'

export async function excelSave(args: Record<string, unknown>) {
  const { workbookId } = excelSaveSchema.parse(args)

  const entry = workbooks.get(workbookId)
  if (!entry) {
    return { success: false, message: `工作簿 ${workbookId} 不存在` }
  }

  const dir = path.dirname(entry.outputPath)
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true })
  }

  await entry.workbook.xlsx.writeFile(entry.outputPath)
  workbooks.delete(workbookId)

  return {
    success: true,
    message: `工作簿已保存至: ${entry.outputPath}`,
    path: entry.outputPath
  }
}
