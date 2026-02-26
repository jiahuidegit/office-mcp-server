/**
 * Excel 工作簿内存存储
 */

import ExcelJS from 'exceljs'
import { Theme } from '../../styles/themes/index.js'

export interface WorkbookEntry {
  workbook: ExcelJS.Workbook
  outputPath: string
  theme: Theme
  title: string
  createdAt: number
}

export const workbooks = new Map<string, WorkbookEntry>()

/** TTL 清理：30 分钟过期 */
const WORKBOOK_TTL_MS = 30 * 60 * 1000

export function cleanExpiredWorkbooks(): void {
  const now = Date.now()
  for (const [id, entry] of workbooks) {
    if (now - entry.createdAt > WORKBOOK_TTL_MS) {
      workbooks.delete(id)
    }
  }
}
