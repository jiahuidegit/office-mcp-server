/**
 * word_create - 创建文档
 */

import { getTheme } from '../../styles/themes/index.js'
import { wordCreateSchema } from '../../schemas/index.js'
import { generateId } from '../../utils/index.js'
import { documents, cleanExpiredDocs } from './store.js'

export function wordCreate(args: Record<string, unknown>) {
  const { title, template, theme, author, outputPath } = wordCreateSchema.parse(args)

  cleanExpiredDocs()

  const id = generateId('word')
  const themeConfig = getTheme(theme)

  documents.set(id, {
    children: [],
    outputPath,
    theme: themeConfig,
    title,
    author,
    createdAt: Date.now()
  })

  const templateNote = template && template !== 'blank'
    ? `（注意：模板 "${template}" 暂未实现，已使用空白文档）`
    : ''

  return {
    success: true,
    docId: id,
    message: `文档 "${title}" 创建成功，主题: ${themeConfig.displayName}${templateNote}`
  }
}
