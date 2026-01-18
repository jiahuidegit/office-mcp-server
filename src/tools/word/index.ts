/**
 * Word 工具模块 - 完整实现
 */

import { Tool } from '@modelcontextprotocol/sdk/types.js'
import {
  Document,
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
  Packer,
  ShadingType,
  Header,
  Footer,
  PageNumber,
  NumberFormat
} from 'docx'
import * as fs from 'fs'
import * as path from 'path'
import { getTheme, Theme } from '../../styles/themes/index.js'

// 文档存储（内存中维护）
interface DocEntry {
  children: Array<Paragraph | Table>
  outputPath: string
  theme: Theme
  title: string
  author?: string
}

const documents = new Map<string, DocEntry>()

function generateId(): string {
  return `word_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`
}

// ============ 工具定义 ============

export const wordTools: Tool[] = [
  {
    name: 'word_create',
    description: '创建一个新的 Word 文档',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string', description: '文档标题' },
        template: {
          type: 'string',
          enum: ['tech-doc', 'weekly-report', 'monthly-report', 'prd', 'meeting-notes', 'blank'],
          description: '使用的模板类型'
        },
        theme: {
          type: 'string',
          enum: ['alibaba', 'tencent', 'bytedance', 'default'],
          description: '主题风格'
        },
        author: { type: 'string', description: '作者' },
        outputPath: { type: 'string', description: '输出路径' }
      },
      required: ['title', 'outputPath']
    }
  },
  {
    name: 'word_add_heading',
    description: '添加标题（支持1-6级）',
    inputSchema: {
      type: 'object',
      properties: {
        docId: { type: 'string', description: '文档ID' },
        text: { type: 'string', description: '标题文本' },
        level: { type: 'number', enum: [1, 2, 3, 4, 5, 6], description: '标题级别' },
        numbering: { type: 'boolean', description: '是否自动编号' }
      },
      required: ['docId', 'text', 'level']
    }
  },
  {
    name: 'word_add_paragraph',
    description: '添加正文段落',
    inputSchema: {
      type: 'object',
      properties: {
        docId: { type: 'string', description: '文档ID' },
        text: { type: 'string', description: '段落文本' },
        style: {
          type: 'string',
          enum: ['normal', 'quote', 'note', 'warning', 'tip'],
          description: '段落样式'
        },
        bold: { type: 'boolean', description: '是否加粗' },
        italic: { type: 'boolean', description: '是否斜体' }
      },
      required: ['docId', 'text']
    }
  },
  {
    name: 'word_add_table',
    description: '添加表格',
    inputSchema: {
      type: 'object',
      properties: {
        docId: { type: 'string', description: '文档ID' },
        headers: { type: 'array', items: { type: 'string' }, description: '表头' },
        rows: { type: 'array', items: { type: 'array' }, description: '数据行' },
        style: {
          type: 'string',
          enum: ['striped', 'bordered', 'minimal', 'professional'],
          description: '表格样式'
        }
      },
      required: ['docId', 'headers', 'rows']
    }
  },
  {
    name: 'word_add_list',
    description: '添加有序/无序列表',
    inputSchema: {
      type: 'object',
      properties: {
        docId: { type: 'string', description: '文档ID' },
        items: { type: 'array', items: { type: 'string' }, description: '列表项' },
        ordered: { type: 'boolean', description: '是否有序列表' }
      },
      required: ['docId', 'items']
    }
  },
  {
    name: 'word_add_code',
    description: '添加代码块',
    inputSchema: {
      type: 'object',
      properties: {
        docId: { type: 'string', description: '文档ID' },
        code: { type: 'string', description: '代码内容' },
        language: { type: 'string', description: '编程语言' }
      },
      required: ['docId', 'code']
    }
  },
  {
    name: 'word_save',
    description: '保存文档',
    inputSchema: {
      type: 'object',
      properties: {
        docId: { type: 'string', description: '文档ID' },
        addPageNumbers: { type: 'boolean', description: '添加页码' },
        addHeader: { type: 'string', description: '页眉文字' },
        addFooter: { type: 'string', description: '页脚文字' }
      },
      required: ['docId']
    }
  }
]

// ============ 工具实现 ============

// 创建文档
function wordCreate(args: Record<string, unknown>) {
  const { title, theme = 'default', author, outputPath } = args as {
    title: string
    theme?: string
    author?: string
    outputPath: string
  }

  const id = generateId()
  const themeConfig = getTheme(theme)

  documents.set(id, {
    children: [],
    outputPath,
    theme: themeConfig,
    title,
    author
  })

  return {
    success: true,
    docId: id,
    message: `文档 "${title}" 创建成功，主题: ${themeConfig.displayName}`
  }
}

// 添加标题
function wordAddHeading(args: Record<string, unknown>) {
  const { docId, text, level } = args as {
    docId: string
    text: string
    level: number
  }

  const doc = documents.get(docId)
  if (!doc) {
    return { success: false, message: `文档 ${docId} 不存在` }
  }

  const headingLevelMap: Record<number, typeof HeadingLevel.HEADING_1> = {
    1: HeadingLevel.HEADING_1,
    2: HeadingLevel.HEADING_2,
    3: HeadingLevel.HEADING_3,
    4: HeadingLevel.HEADING_4,
    5: HeadingLevel.HEADING_5,
    6: HeadingLevel.HEADING_6
  }

  // 根据级别获取样式
  const styleKey = `heading${Math.min(level, 3)}` as 'heading1' | 'heading2' | 'heading3'
  const style = doc.theme.word[styleKey]

  const paragraph = new Paragraph({
    heading: headingLevelMap[level],
    children: [
      new TextRun({
        text,
        bold: style.bold,
        color: style.color,
        size: style.fontSize
      })
    ],
    spacing: { after: 200, before: level === 1 ? 400 : 200 }
  })

  doc.children.push(paragraph)

  return { success: true, message: `已添加 ${level} 级标题: ${text}` }
}

// 添加段落
function wordAddParagraph(args: Record<string, unknown>) {
  const { docId, text, style = 'normal', bold = false, italic = false } = args as {
    docId: string
    text: string
    style?: string
    bold?: boolean
    italic?: boolean
  }

  const doc = documents.get(docId)
  if (!doc) {
    return { success: false, message: `文档 ${docId} 不存在` }
  }

  const themeStyle = doc.theme.word.paragraph

  // 根据样式类型设置不同的外观
  let shading
  let border
  if (style === 'quote') {
    shading = {
      type: ShadingType.SOLID,
      color: doc.theme.word.quote.background
    }
    border = {
      left: { style: BorderStyle.SINGLE, size: 24, color: doc.theme.word.quote.borderLeft }
    }
  } else if (style === 'note' || style === 'tip') {
    shading = {
      type: ShadingType.SOLID,
      color: 'e6f7ff'
    }
  } else if (style === 'warning') {
    shading = {
      type: ShadingType.SOLID,
      color: 'fff7e6'
    }
  }

  const paragraph = new Paragraph({
    children: [
      new TextRun({
        text,
        color: themeStyle.color,
        size: themeStyle.fontSize,
        bold,
        italics: italic
      })
    ],
    spacing: { after: 200, line: themeStyle.lineSpacing * 240 },
    shading,
    border,
    indent: style === 'quote' ? { left: 720 } : undefined
  })

  doc.children.push(paragraph)

  return { success: true, message: '已添加段落' }
}

// 添加表格
function wordAddTable(args: Record<string, unknown>) {
  const { docId, headers, rows, style = 'professional' } = args as {
    docId: string
    headers: string[]
    rows: string[][]
    style?: string
  }

  const doc = documents.get(docId)
  if (!doc) {
    return { success: false, message: `文档 ${docId} 不存在` }
  }

  const tableStyle = doc.theme.word.table
  const colCount = headers.length

  // 创建表头行
  const headerRow = new TableRow({
    children: headers.map(
      header =>
        new TableCell({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: header,
                  bold: true,
                  color: tableStyle.headerColor,
                  size: 22
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

  // 创建数据行
  const dataRows = rows.map(
    (row, rowIndex) =>
      new TableRow({
        children: row.map(
          cell =>
            new TableCell({
              children: [
                new Paragraph({
                  children: [new TextRun({ text: String(cell), size: 20 })]
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

  const table = new Table({
    rows: [headerRow, ...dataRows],
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders:
      style === 'minimal'
        ? undefined
        : {
            top: { style: BorderStyle.SINGLE, size: 1, color: tableStyle.borderColor },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: tableStyle.borderColor },
            left: { style: BorderStyle.SINGLE, size: 1, color: tableStyle.borderColor },
            right: { style: BorderStyle.SINGLE, size: 1, color: tableStyle.borderColor },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: tableStyle.borderColor },
            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: tableStyle.borderColor }
          }
  })

  doc.children.push(table)
  // 添加表格后的空行
  doc.children.push(new Paragraph({ children: [] }))

  return { success: true, message: `已添加 ${headers.length} 列 ${rows.length} 行的表格` }
}

// 添加列表
function wordAddList(args: Record<string, unknown>) {
  const { docId, items, ordered = false } = args as {
    docId: string
    items: string[]
    ordered?: boolean
  }

  const doc = documents.get(docId)
  if (!doc) {
    return { success: false, message: `文档 ${docId} 不存在` }
  }

  const themeStyle = doc.theme.word.paragraph

  items.forEach((item, index) => {
    const bullet = ordered ? `${index + 1}. ` : '• '
    const paragraph = new Paragraph({
      children: [
        new TextRun({
          text: bullet + item,
          color: themeStyle.color,
          size: themeStyle.fontSize
        })
      ],
      indent: { left: 720 },
      spacing: { after: 120 }
    })
    doc.children.push(paragraph)
  })

  return { success: true, message: `已添加 ${items.length} 项${ordered ? '有序' : '无序'}列表` }
}

// 添加代码块
function wordAddCode(args: Record<string, unknown>) {
  const { docId, code, language = '' } = args as {
    docId: string
    code: string
    language?: string
  }

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
          new TextRun({
            text: language,
            size: 18,
            color: '888888',
            font: codeStyle.fontFamily
          })
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
          new TextRun({
            text: line || ' ',
            font: codeStyle.fontFamily,
            size: codeStyle.fontSize
          })
        ],
        shading: { type: ShadingType.SOLID, color: codeStyle.background },
        spacing: { after: 0, line: 280 },
        indent: { left: 360 }
      })
    )
  })

  // 代码块后空行
  doc.children.push(new Paragraph({ children: [], spacing: { after: 200 } }))

  return { success: true, message: `已添加代码块 (${language || '无语言标记'})` }
}

// 保存文档
async function wordSave(args: Record<string, unknown>) {
  const { docId, addPageNumbers = false, addHeader, addFooter } = args as {
    docId: string
    addPageNumbers?: boolean
    addHeader?: string
    addFooter?: string
  }

  const doc = documents.get(docId)
  if (!doc) {
    return { success: false, message: `文档 ${docId} 不存在` }
  }

  // 构建 header
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

  // 构建 footer
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
                        new TextRun({
                          children: [PageNumber.CURRENT],
                          color: '888888',
                          size: 18
                        }),
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

  // 创建最终文档
  const document = new Document({
    creator: doc.author || 'Office MCP Server',
    title: doc.title,
    sections: [
      {
        properties: {
          page: {
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
          }
        },
        headers,
        footers,
        children: doc.children
      }
    ]
  })

  // 确保目录存在
  const dir = path.dirname(doc.outputPath)
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true })
  }

  // 写入文件
  const buffer = await Packer.toBuffer(document)
  fs.writeFileSync(doc.outputPath, buffer)

  // 清理内存
  documents.delete(docId)

  return {
    success: true,
    message: `文档已保存至: ${doc.outputPath}`,
    path: doc.outputPath
  }
}

// ============ 工具调用分发 ============

export async function handleWordTool(name: string, args: Record<string, unknown>) {
  let result

  switch (name) {
    case 'word_create':
      result = wordCreate(args)
      break
    case 'word_add_heading':
      result = wordAddHeading(args)
      break
    case 'word_add_paragraph':
      result = wordAddParagraph(args)
      break
    case 'word_add_table':
      result = wordAddTable(args)
      break
    case 'word_add_list':
      result = wordAddList(args)
      break
    case 'word_add_code':
      result = wordAddCode(args)
      break
    case 'word_save':
      result = await wordSave(args)
      break
    default:
      result = { success: false, message: `未知的 Word 工具: ${name}` }
  }

  return {
    content: [{ type: 'text' as const, text: JSON.stringify(result, null, 2) }]
  }
}
