/**
 * word_save - 保存文档
 */

import {
  Document, Paragraph, TextRun, Header, Footer,
  PageNumber, AlignmentType, Packer
} from 'docx'
import * as fs from 'fs'
import * as path from 'path'
import { wordSaveSchema } from '../../schemas/index.js'
import { documents } from './store.js'

export async function wordSave(args: Record<string, unknown>) {
  const { docId, addPageNumbers, addHeader, addFooter } = wordSaveSchema.parse(args)

  const doc = documents.get(docId)
  if (!doc) {
    return { success: false, message: `文档 ${docId} 不存在` }
  }

  const pageLayout = doc.theme.pageLayout

  const headers = addHeader
    ? {
        default: new Header({
          children: [
            new Paragraph({
              children: [new TextRun({ text: addHeader, color: '888888', size: 18 })],
              alignment: AlignmentType.RIGHT
            })
          ]
        })
      }
    : undefined

  const footers =
    addFooter || addPageNumbers
      ? {
          default: new Footer({
            children: [
              new Paragraph({
                children: [
                  ...(addFooter ? [new TextRun({ text: addFooter, color: '888888', size: 18 })] : []),
                  ...(addPageNumbers
                    ? [
                        new TextRun({ text: addFooter ? '  |  第 ' : '第 ', color: '888888', size: 18 }),
                        new TextRun({ children: [PageNumber.CURRENT], color: '888888', size: 18 }),
                        new TextRun({ text: ' 页', color: '888888', size: 18 })
                      ]
                    : [])
                ],
                alignment: AlignmentType.CENTER
              })
            ]
          })
        }
      : undefined

  const document = new Document({
    creator: doc.author || 'Office MCP Server',
    title: doc.title,
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: pageLayout.marginTop,
              bottom: pageLayout.marginBottom,
              left: pageLayout.marginLeft,
              right: pageLayout.marginRight
            }
          }
        },
        headers,
        footers,
        children: doc.children
      }
    ]
  })

  const dir = path.dirname(doc.outputPath)
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true })
  }

  const buffer = await Packer.toBuffer(document)
  fs.writeFileSync(doc.outputPath, buffer)

  documents.delete(docId)

  return {
    success: true,
    message: `文档已保存至: ${doc.outputPath}`,
    path: doc.outputPath
  }
}
