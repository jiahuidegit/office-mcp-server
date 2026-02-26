/**
 * Word 工具模块 - 工具定义与分发
 */

import { Tool } from '@modelcontextprotocol/sdk/types.js'
import { ZodError } from 'zod'
import { wordCreate } from './create.js'
import { wordAddHeading } from './heading.js'
import { wordAddParagraph } from './paragraph.js'
import { wordAddTable } from './table.js'
import { wordAddList } from './list.js'
import { wordAddCode } from './code.js'
import { wordAddDiagram } from './diagram.js'
import { wordSave } from './save.js'

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
          description: '图表风格'
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

// ============ 工具调用分发 ============

export async function handleWordTool(name: string, args: Record<string, unknown>) {
  try {
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
  } catch (error) {
    if (error instanceof ZodError) {
      const issues = error.issues.map(i => `${i.path.join('.')}: ${i.message}`).join('; ')
      return {
        content: [{ type: 'text' as const, text: JSON.stringify({
          success: false,
          message: `参数校验失败: ${issues}`
        }, null, 2) }]
      }
    }
    throw error
  }
}
