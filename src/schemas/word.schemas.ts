/**
 * Word 工具 Zod Schema 定义
 */

import { z } from 'zod'

export const wordCreateSchema = z.object({
  title: z.string().min(1, '标题不能为空'),
  template: z
    .enum(['tech-doc', 'weekly-report', 'monthly-report', 'prd', 'meeting-notes', 'blank'])
    .optional(),
  theme: z
    .enum(['government', 'academic', 'software', 'business', 'alibaba', 'tencent', 'bytedance', 'minimal', 'default'])
    .optional()
    .default('default'),
  author: z.string().optional(),
  outputPath: z.string().min(1, '输出路径不能为空')
})

export const wordAddHeadingSchema = z.object({
  docId: z.string().min(1, '文档ID不能为空'),
  text: z.string().min(1, '标题文本不能为空'),
  level: z.number().int().min(1).max(6),
  numbering: z.boolean().optional()
})

export const wordAddParagraphSchema = z.object({
  docId: z.string().min(1, '文档ID不能为空'),
  text: z.string(),
  style: z.enum(['normal', 'quote', 'note', 'warning', 'tip']).optional().default('normal'),
  bold: z.boolean().optional().default(false),
  italic: z.boolean().optional().default(false)
})

export const wordAddTableSchema = z.object({
  docId: z.string().min(1, '文档ID不能为空'),
  headers: z.array(z.string()).min(1, '表头不能为空'),
  rows: z.array(z.array(z.unknown())),
  style: z.enum(['striped', 'bordered', 'minimal', 'professional']).optional().default('professional')
})

export const wordAddListSchema = z.object({
  docId: z.string().min(1, '文档ID不能为空'),
  items: z.array(z.string()).min(1, '列表项不能为空'),
  ordered: z.boolean().optional().default(false)
})

export const wordAddCodeSchema = z.object({
  docId: z.string().min(1, '文档ID不能为空'),
  code: z.string(),
  language: z.string().optional().default('')
})

export const wordAddDiagramSchema = z.object({
  docId: z.string().min(1, '文档ID不能为空'),
  mermaid: z.string().min(1, 'Mermaid 代码不能为空'),
  caption: z.string().optional(),
  width: z.number().positive().optional().default(15),
  theme: z
    .enum(['professional', 'fresh', 'business', 'tech', 'warm', 'default'])
    .optional()
    .default('default')
})

export const wordSaveSchema = z.object({
  docId: z.string().min(1, '文档ID不能为空'),
  addPageNumbers: z.boolean().optional().default(false),
  addHeader: z.string().optional(),
  addFooter: z.string().optional()
})
