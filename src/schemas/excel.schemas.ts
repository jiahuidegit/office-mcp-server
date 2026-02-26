/**
 * Excel 工具 Zod Schema 定义
 */

import { z } from 'zod'

export const excelCreateSchema = z.object({
  title: z.string().min(1, '标题不能为空'),
  template: z
    .enum(['data-report', 'project-tracker', 'budget', 'blank'])
    .optional(),
  theme: z
    .enum(['alibaba', 'tencent', 'bytedance', 'default'])
    .optional()
    .default('default'),
  outputPath: z.string().min(1, '输出路径不能为空')
})

export const excelAddSheetSchema = z.object({
  workbookId: z.string().min(1, '工作簿ID不能为空'),
  sheetName: z.string().min(1, '工作表名称不能为空'),
  tabColor: z.string().optional()
})

export const excelWriteDataSchema = z.object({
  workbookId: z.string().min(1, '工作簿ID不能为空'),
  sheetName: z.string().min(1, '工作表名称不能为空'),
  headers: z.array(z.string()).min(1, '表头不能为空'),
  rows: z.array(z.array(z.unknown())),
  startCell: z
    .string()
    .regex(/^[A-Z]+\d+$/, '单元格格式错误，如 A1、B2')
    .optional()
    .default('A1'),
  autoFilter: z.boolean().optional().default(false),
  freezeHeader: z.boolean().optional().default(false),
  style: z
    .enum(['striped', 'bordered', 'professional', 'minimal'])
    .optional()
    .default('professional')
})

export const excelAddChartSchema = z.object({
  workbookId: z.string().min(1, '工作簿ID不能为空'),
  sheetName: z.string().min(1, '工作表名称不能为空'),
  chartType: z.enum(['bar', 'column', 'line', 'pie', 'doughnut', 'area']),
  title: z.string().optional(),
  categories: z.array(z.string()).min(1, '分类标签不能为空'),
  series: z.array(
    z.object({
      name: z.string(),
      values: z.array(z.number())
    })
  ).min(1, '数据系列不能为空'),
  position: z
    .string()
    .regex(/^[A-Z]+\d+$/, '位置格式错误，如 F2')
    .optional()
    .default('F2')
})

export const excelAddFormulaSchema = z.object({
  workbookId: z.string().min(1, '工作簿ID不能为空'),
  sheetName: z.string().min(1, '工作表名称不能为空'),
  cell: z.string().regex(/^[A-Z]+\d+$/, '单元格格式错误，如 A1'),
  formula: z.string().min(1, '公式不能为空')
})

export const excelSetColumnWidthSchema = z.object({
  workbookId: z.string().min(1, '工作簿ID不能为空'),
  sheetName: z.string().min(1, '工作表名称不能为空'),
  columns: z.record(z.string(), z.number().positive())
})

export const excelMergeCellsSchema = z.object({
  workbookId: z.string().min(1, '工作簿ID不能为空'),
  sheetName: z.string().min(1, '工作表名称不能为空'),
  range: z.string().min(1, '单元格范围不能为空')
})

export const excelSaveSchema = z.object({
  workbookId: z.string().min(1, '工作簿ID不能为空')
})
