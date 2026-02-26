/**
 * word_add_code - 添加代码块
 */

import { Paragraph, TextRun, ShadingType, BorderStyle, LineRuleType } from 'docx'
import { wordAddCodeSchema } from '../../schemas/index.js'
import { documents } from './store.js'

export function wordAddCode(args: Record<string, unknown>) {
  const { docId, code, language } = wordAddCodeSchema.parse(args)

  const doc = documents.get(docId)
  if (!doc) {
    return { success: false, message: `文档 ${docId} 不存在` }
  }

  const codeStyle = doc.theme.word.code

  // 语言标签
  if (language) {
    doc.children.push(
      new Paragraph({
        children: [
          new TextRun({ text: language, size: 18, color: '888888', font: codeStyle.font })
        ],
        spacing: { after: 0 }
      })
    )
  }

  // 代码内容（按行分割）
  const lines = code.split('\n')
  lines.forEach(line => {
    doc.children.push(
      new Paragraph({
        children: [
          new TextRun({ text: line || ' ', font: codeStyle.font, size: codeStyle.fontSize })
        ],
        shading: { type: ShadingType.SOLID, color: codeStyle.background },
        spacing: { after: 0, line: 280, lineRule: LineRuleType.AUTO },
        indent: { left: 360 },
        border: {
          left: { style: BorderStyle.SINGLE, size: 1, color: codeStyle.borderColor }
        }
      })
    )
  })

  // 代码块后空行
  doc.children.push(new Paragraph({ children: [], spacing: { after: 200 } }))

  return { success: true, message: `已添加代码块 (${language || '无语言标记'})` }
}
