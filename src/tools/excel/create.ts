/**
 * excel_create - 创建工作簿
 */

import ExcelJS from 'exceljs'
import { getTheme } from '../../styles/themes/index.js'
import { excelCreateSchema } from '../../schemas/index.js'
import { generateId } from '../../utils/index.js'
import { workbooks, cleanExpiredWorkbooks } from './store.js'

export function excelCreate(args: Record<string, unknown>) {
  const { title, template, theme, outputPath } = excelCreateSchema.parse(args)

  cleanExpiredWorkbooks()

  const id = generateId('excel')
  const themeConfig = getTheme(theme)
  const workbook = new ExcelJS.Workbook()

  workbook.creator = 'Office MCP Server'
  workbook.created = new Date()
  workbook.title = title

  workbooks.set(id, {
    workbook,
    outputPath,
    theme: themeConfig,
    title,
    createdAt: Date.now()
  })

  const templateNote = template && template !== 'blank'
    ? `（注意：模板 "${template}" 暂未实现，已使用空白文档）`
    : ''

  return {
    success: true,
    workbookId: id,
    message: `工作簿 "${title}" 创建成功，主题: ${themeConfig.displayName}${templateNote}`
  }
}
