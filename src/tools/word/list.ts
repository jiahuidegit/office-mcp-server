/**
 * word_add_list - 添加列表
 */

import { Paragraph, TextRun, LineRuleType } from 'docx'
import { wordAddListSchema } from '../../schemas/index.js'
import { documents } from './store.js'

export function wordAddList(args: Record<string, unknown>) {
  const { docId, items, ordered } = wordAddListSchema.parse(args)

  const doc = documents.get(docId)
  if (!doc) {
    return { success: false, message: `文档 ${docId} 不存在` }
  }

  const listStyle = doc.theme.word.list
  const format = doc.theme.word.paragraph.format

  const lineSpacing = format.lineSpacingType === 'multiple'
    ? { line: Math.round(format.lineSpacing * 240), lineRule: LineRuleType.AUTO }
    : { line: Math.round(format.lineSpacing * 20), lineRule: LineRuleType.EXACT }

  items.forEach((item, index) => {
    const bullet = ordered ? `${index + 1}. ` : '• '
    const paragraph = new Paragraph({
      children: [
        new TextRun({
          text: bullet + item,
          size: listStyle.fontSize,
          font: listStyle.font
        })
      ],
      indent: { left: listStyle.indent },
      spacing: {
        ...lineSpacing,
        after: listStyle.spacing * 20
      }
    })
    doc.children.push(paragraph)
  })

  return { success: true, message: `已添加 ${items.length} 项${ordered ? '有序' : '无序'}列表` }
}
