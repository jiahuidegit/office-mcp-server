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
  LineRuleType,
  ImageRun
} from 'docx'
import * as fs from 'fs'
import * as path from 'path'
import * as https from 'https'
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
          enum: ['government', 'academic', 'business', 'alibaba', 'tencent', 'bytedance', 'minimal', 'default'],
          description: '主题风格：government(党政公文)、academic(学术论文)、business(商务报告)、alibaba/tencent/bytedance(企业风格)、minimal(简约黑白)'
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
    name: 'word_add_diagram',
    description: '添加 Mermaid 流程图/架构图（自动转换为图片插入）',
    inputSchema: {
      type: 'object',
      properties: {
        docId: { type: 'string', description: '文档ID' },
        mermaid: { type: 'string', description: 'Mermaid 图表代码' },
        caption: { type: 'string', description: '图表标题（可选）' },
        width: { type: 'number', description: '图片宽度（厘米），默认15' },
        theme: {
          type: 'string',
          enum: ['professional', 'fresh', 'business', 'tech', 'warm', 'default'],
          description: '图表风格：professional(专业蓝)、fresh(清新绿)、business(商务灰)、tech(科技紫)、warm(暖橙色)、default(原生)'
        }
      },
      required: ['docId', 'mermaid']
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

  // 根据级别获取样式配置
  const styleKey = `heading${Math.min(level, 3)}` as 'heading1' | 'heading2' | 'heading3'
  const style = doc.theme.word[styleKey]

  // 对齐方式映射
  const alignmentMap: Record<string, typeof AlignmentType.LEFT> = {
    left: AlignmentType.LEFT,
    center: AlignmentType.CENTER,
    right: AlignmentType.RIGHT
  }
  const alignment = alignmentMap[style.alignment] || AlignmentType.LEFT

  // 段前段后间距（磅转 twips：1磅 = 20 twips）
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
  const format = themeStyle.format

  // 根据样式类型设置不同的外观
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

  // 计算行距
  // multiple: 倍数行距，240 = 单倍行距，所以 1.5 倍 = 360
  // exact: 固定值行距，单位是 twips，1磅 = 20 twips
  const lineSpacing = format.lineSpacingType === 'multiple'
    ? { line: Math.round(format.lineSpacing * 240), lineRule: LineRuleType.AUTO }
    : { line: Math.round(format.lineSpacing * 20), lineRule: LineRuleType.EXACT }

  // 首行缩进：字符数 * 字号（half-points）* 10（转为 twips）
  // 1个汉字 = 字号大小，例如小四(12pt) 的 2 字符 = 24pt = 480 twips
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

  // 创建表头行（使用主题配置的表头样式）
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

  // 创建数据行（使用主题配置的正文样式）
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

  // 边框样式
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

  const listStyle = doc.theme.word.list
  const format = doc.theme.word.paragraph.format

  // 计算行距
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
        after: listStyle.spacing * 20  // 磅转 twips
      }
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
            font: codeStyle.font
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
            font: codeStyle.font,
            size: codeStyle.fontSize
          })
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

// ============ Mermaid 图表支持 ============

// Mermaid 预设主题配置
const mermaidThemes: Record<string, string> = {
  // 专业蓝 - 适合技术文档
  professional: `%%{init: {
    'theme': 'base',
    'themeVariables': {
      'primaryColor': '#4a90d9',
      'primaryTextColor': '#ffffff',
      'primaryBorderColor': '#2d6cb5',
      'secondaryColor': '#e8f4fc',
      'tertiaryColor': '#f5f9fc',
      'lineColor': '#5c6bc0',
      'textColor': '#333333',
      'fontSize': '14px',
      'fontFamily': 'Microsoft YaHei, Arial, sans-serif'
    }
  }}%%`,

  // 清新绿 - 适合流程图
  fresh: `%%{init: {
    'theme': 'base',
    'themeVariables': {
      'primaryColor': '#52c41a',
      'primaryTextColor': '#ffffff',
      'primaryBorderColor': '#389e0d',
      'secondaryColor': '#f6ffed',
      'tertiaryColor': '#d9f7be',
      'lineColor': '#73d13d',
      'textColor': '#333333',
      'fontSize': '14px',
      'fontFamily': 'Microsoft YaHei, Arial, sans-serif'
    }
  }}%%`,

  // 商务灰 - 适合正式报告
  business: `%%{init: {
    'theme': 'base',
    'themeVariables': {
      'primaryColor': '#595959',
      'primaryTextColor': '#ffffff',
      'primaryBorderColor': '#434343',
      'secondaryColor': '#fafafa',
      'tertiaryColor': '#f0f0f0',
      'lineColor': '#8c8c8c',
      'textColor': '#262626',
      'fontSize': '14px',
      'fontFamily': 'Microsoft YaHei, Arial, sans-serif'
    }
  }}%%`,

  // 科技紫 - 适合产品文档
  tech: `%%{init: {
    'theme': 'base',
    'themeVariables': {
      'primaryColor': '#722ed1',
      'primaryTextColor': '#ffffff',
      'primaryBorderColor': '#531dab',
      'secondaryColor': '#f9f0ff',
      'tertiaryColor': '#efdbff',
      'lineColor': '#9254de',
      'textColor': '#333333',
      'fontSize': '14px',
      'fontFamily': 'Microsoft YaHei, Arial, sans-serif'
    }
  }}%%`,

  // 暖橙色 - 适合产品方案
  warm: `%%{init: {
    'theme': 'base',
    'themeVariables': {
      'primaryColor': '#fa8c16',
      'primaryTextColor': '#ffffff',
      'primaryBorderColor': '#d46b08',
      'secondaryColor': '#fff7e6',
      'tertiaryColor': '#ffe7ba',
      'lineColor': '#ffa940',
      'textColor': '#333333',
      'fontSize': '14px',
      'fontFamily': 'Microsoft YaHei, Arial, sans-serif'
    }
  }}%%`,

  // 默认 - Mermaid 原生
  default: ''
}

/**
 * 将 Mermaid 代码转换为图片 Buffer
 * 使用 kroki.io 在线服务（POST 方式）
 */
async function mermaidToImage(mermaidCode: string, theme: string = 'professional'): Promise<Buffer> {
  // 获取主题配置
  const themeConfig = mermaidThemes[theme] || mermaidThemes.professional

  // 清理代码并添加主题配置
  const cleanCode = mermaidCode.trim()
  const fullCode = themeConfig ? `${themeConfig}\n${cleanCode}` : cleanCode

  return new Promise((resolve, reject) => {
    const postData = fullCode

    const options = {
      hostname: 'kroki.io',
      port: 443,
      path: '/mermaid/png',
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain',
        'Content-Length': Buffer.byteLength(postData, 'utf-8')
      }
    }

    const req = https.request(options, (res) => {
      if (res.statusCode !== 200) {
        let errorBody = ''
        res.on('data', chunk => errorBody += chunk)
        res.on('end', () => {
          reject(new Error(`Mermaid 渲染失败: HTTP ${res.statusCode}, ${errorBody.slice(0, 100)}`))
        })
        return
      }

      const chunks: Buffer[] = []
      res.on('data', (chunk: Buffer) => chunks.push(chunk))
      res.on('end', () => resolve(Buffer.concat(chunks)))
      res.on('error', reject)
    })

    req.on('error', reject)
    req.write(postData)
    req.end()
  })
}

/**
 * 获取图片尺寸（简单的 PNG 解析）
 */
function getPngDimensions(buffer: Buffer): { width: number; height: number } {
  // PNG 文件头: 89 50 4E 47 0D 0A 1A 0A
  // IHDR chunk 在第 16-23 字节包含宽高
  if (buffer.length < 24 || buffer[0] !== 0x89 || buffer[1] !== 0x50) {
    // 默认尺寸
    return { width: 800, height: 400 }
  }

  const width = buffer.readUInt32BE(16)
  const height = buffer.readUInt32BE(20)

  return { width, height }
}

// 添加 Mermaid 图表
async function wordAddDiagram(args: Record<string, unknown>) {
  const {
    docId,
    mermaid: mermaidCode,
    caption,
    width: targetWidthCm = 15,
    theme = 'default'
  } = args as {
    docId: string
    mermaid: string
    caption?: string
    width?: number
    theme?: string
  }

  const doc = documents.get(docId)
  if (!doc) {
    return { success: false, message: `文档 ${docId} 不存在` }
  }

  try {
    // 1. 将 Mermaid 代码转换为图片
    const imageBuffer = await mermaidToImage(mermaidCode, theme)

    // 2. 获取图片原始尺寸
    const { width: origWidth, height: origHeight } = getPngDimensions(imageBuffer)

    // 3. 计算目标尺寸（保持宽高比）
    // docx 的 ImageRun transformation 单位是像素
    // 1厘米 ≈ 37.8 像素（96 DPI）
    const targetWidthPx = Math.round(targetWidthCm * 37.8)
    const scale = targetWidthPx / origWidth
    const targetHeightPx = Math.round(origHeight * scale)

    // 4. 创建图片段落
    const imageParagraph = new Paragraph({
      children: [
        new ImageRun({
          data: imageBuffer,
          transformation: {
            width: targetWidthPx,
            height: targetHeightPx
          },
          type: 'png'
        })
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 100 }
    })

    doc.children.push(imageParagraph)

    // 5. 添加图表标题（如果有）
    if (caption) {
      const captionParagraph = new Paragraph({
        children: [
          new TextRun({
            text: caption,
            size: 21,  // 五号字
            color: '666666',
            font: doc.theme.word.fonts.body
          })
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 }
      })
      doc.children.push(captionParagraph)
    }

    return { success: true, message: `已添加 Mermaid 图表${caption ? `: ${caption}` : ''}` }

  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error)
    return { success: false, message: `Mermaid 图表渲染失败: ${errMsg}` }
  }
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

  // 获取页面布局配置
  const pageLayout = doc.theme.pageLayout

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

  // 创建最终文档（应用主题页边距配置）
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
    case 'word_add_diagram':
      result = await wordAddDiagram(args)
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
