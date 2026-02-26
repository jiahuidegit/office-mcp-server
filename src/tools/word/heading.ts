/**
 * word_add_heading - 添加标题
 */

import { Paragraph, TextRun, HeadingLevel, AlignmentType } from 'docx'
import { wordAddHeadingSchema } from '../../schemas/index.js'
import { documents } from './store.js'

export function wordAddHeading(args: Record<string, unknown>) {
  const { docId, text, level } = wordAddHeadingSchema.parse(args)

  const doc = documents.get(docId)
  if (!doc) {
    return { success: false, message: `文档 ${docId} 不存在` }
  }

  const headingLevelMap: Record<number, (typeof HeadingLevel)[keyof typeof HeadingLevel]> = {
    1: HeadingLevel.HEADING_1,
    2: HeadingLevel.HEADING_2,
    3: HeadingLevel.HEADING_3,
    4: HeadingLevel.HEADING_4,
    5: HeadingLevel.HEADING_5,
    6: HeadingLevel.HEADING_6
  }

  const styleKey = `heading${Math.min(level, 3)}` as 'heading1' | 'heading2' | 'heading3'
  const style = doc.theme.word[styleKey]

  const alignmentMap: Record<string, (typeof AlignmentType)[keyof typeof AlignmentType]> = {
    left: AlignmentType.LEFT,
    center: AlignmentType.CENTER,
    right: AlignmentType.RIGHT
  }
  const alignment = alignmentMap[style.alignment] || AlignmentType.LEFT

  const spaceBefore = style.spaceBefore * 20
  const spaceAfter = style.spaceAfter * 20

  const paragraph = new Paragraph({
    heading: headingLevelMap[level],
    alignment,
    children: [
      new TextRun({
        text,
        bold: style.bold,
        color: style.color,
        size: style.fontSize,
        font: style.font
      })
    ],
    spacing: { after: spaceAfter, before: spaceBefore }
  })

  doc.children.push(paragraph)

  return { success: true, message: `已添加 ${level} 级标题: ${text}` }
}
