/**
 * word_add_table - 添加表格
 */

import {
  Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, AlignmentType, ShadingType
} from 'docx'
import { wordAddTableSchema } from '../../schemas/index.js'
import { documents } from './store.js'

export function wordAddTable(args: Record<string, unknown>) {
  const { docId, headers, rows, style } = wordAddTableSchema.parse(args)

  const doc = documents.get(docId)
  if (!doc) {
    return { success: false, message: `文档 ${docId} 不存在` }
  }

  const tableStyle = doc.theme.word.table
  const colCount = headers.length

  const headerRow = new TableRow({
    children: headers.map(
      header =>
        new TableCell({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: header,
                  bold: tableStyle.headerBold,
                  color: tableStyle.headerColor,
                  size: tableStyle.headerFontSize,
                  font: tableStyle.headerFont
                })
              ],
              alignment: AlignmentType.CENTER
            })
          ],
          shading: { type: ShadingType.SOLID, color: tableStyle.headerBg },
          width: { size: 100 / colCount, type: WidthType.PERCENTAGE }
        })
    )
  })

  const dataRows = rows.map(
    (row, rowIndex) =>
      new TableRow({
        children: row.map(
          cell =>
            new TableCell({
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: String(cell),
                      size: tableStyle.bodyFontSize,
                      font: tableStyle.bodyFont
                    })
                  ]
                })
              ],
              shading:
                style === 'striped' && rowIndex % 2 === 1
                  ? { type: ShadingType.SOLID, color: tableStyle.stripedBg }
                  : undefined,
              width: { size: 100 / colCount, type: WidthType.PERCENTAGE }
            })
        )
      })
  )

  const borderConfig = style === 'minimal'
    ? undefined
    : {
        top: { style: BorderStyle.SINGLE, size: tableStyle.borderWidth, color: tableStyle.borderColor },
        bottom: { style: BorderStyle.SINGLE, size: tableStyle.borderWidth, color: tableStyle.borderColor },
        left: { style: BorderStyle.SINGLE, size: tableStyle.borderWidth, color: tableStyle.borderColor },
        right: { style: BorderStyle.SINGLE, size: tableStyle.borderWidth, color: tableStyle.borderColor },
        insideHorizontal: { style: BorderStyle.SINGLE, size: tableStyle.borderWidth, color: tableStyle.borderColor },
        insideVertical: { style: BorderStyle.SINGLE, size: tableStyle.borderWidth, color: tableStyle.borderColor }
      }

  const table = new Table({
    rows: [headerRow, ...dataRows],
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: borderConfig
  })

  doc.children.push(table)
  doc.children.push(new Paragraph({ children: [] }))

  return { success: true, message: `已添加 ${headers.length} 列 ${rows.length} 行的表格` }
}
