/**
 * word_add_paragraph - 添加段落
 */

import { Paragraph, TextRun, ShadingType, BorderStyle, LineRuleType } from 'docx'
import { wordAddParagraphSchema } from '../../schemas/index.js'
import { documents } from './store.js'

export function wordAddParagraph(args: Record<string, unknown>) {
  const { docId, text, style, bold, italic } = wordAddParagraphSchema.parse(args)

  const doc = documents.get(docId)
  if (!doc) {
    return { success: false, message: `文档 ${docId} 不存在` }
  }

  const themeStyle = doc.theme.word.paragraph
  const format = themeStyle.format

  let shading
  let border
  let fontToUse = themeStyle.font
  let fontSizeToUse = themeStyle.fontSize
  let isItalic = italic

  if (style === 'quote') {
    const quoteStyle = doc.theme.word.quote
    shading = { type: ShadingType.SOLID, color: quoteStyle.background }
    border = { left: { style: BorderStyle.SINGLE, size: quoteStyle.borderWidth * 8, color: quoteStyle.borderLeft } }
    fontToUse = quoteStyle.font
    fontSizeToUse = quoteStyle.fontSize
    isItalic = quoteStyle.italic || italic
  } else if (style === 'note' || style === 'tip') {
    shading = { type: ShadingType.SOLID, color: 'e6f7ff' }
  } else if (style === 'warning') {
    shading = { type: ShadingType.SOLID, color: 'fff7e6' }
  }

  const lineSpacing = format.lineSpacingType === 'multiple'
    ? { line: Math.round(format.lineSpacing * 240), lineRule: LineRuleType.AUTO }
    : { line: Math.round(format.lineSpacing * 20), lineRule: LineRuleType.EXACT }

  const firstLineIndent = format.firstLineIndent > 0
    ? Math.round(format.firstLineIndent * themeStyle.fontSize * 10)
    : 0

  const paragraph = new Paragraph({
    children: [
      new TextRun({
        text,
        color: themeStyle.color,
        size: fontSizeToUse,
        font: fontToUse,
        bold,
        italics: isItalic
      })
    ],
    spacing: {
      ...lineSpacing,
      before: format.spaceBefore * 20,
      after: format.spaceAfter * 20
    },
    indent: {
      firstLine: firstLineIndent,
      ...(style === 'quote' ? { left: 720 } : {})
    },
    shading,
    border
  })

  doc.children.push(paragraph)

  return { success: true, message: '已添加段落' }
}
