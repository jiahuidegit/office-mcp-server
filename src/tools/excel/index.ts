/**
 * Excel 工具模块 - 工具定义与分发
 */

import { Tool } from '@modelcontextprotocol/sdk/types.js'
import { ZodError } from 'zod'
import { excelCreate } from './create.js'
import { excelAddSheet } from './sheet.js'
import { excelWriteData } from './write-data.js'
import { excelAddChart } from './chart.js'
import { excelAddFormula } from './formula.js'
import { excelSetColumnWidth } from './column-width.js'
import { excelMergeCells } from './merge-cells.js'
import { excelSave } from './save.js'

// ============ 工具定义 ============

export const excelTools: Tool[] = [
  {
    name: 'excel_create',
    description: '创建 Excel 工作簿',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string', description: '工作簿标题' },
        template: {
          type: 'string',
          enum: ['data-report', 'project-tracker', 'budget', 'blank'],
          description: '使用的模板'
        },
        theme: {
          type: 'string',
          enum: ['alibaba', 'tencent', 'bytedance', 'default'],
          description: '主题风格'
        },
        outputPath: { type: 'string', description: '输出路径' }
      },
      required: ['title', 'outputPath']
    }
  },
  {
    name: 'excel_add_sheet',
    description: '添加工作表',
    inputSchema: {
      type: 'object',
      properties: {
        workbookId: { type: 'string', description: '工作簿ID' },
        sheetName: { type: 'string', description: '工作表名称' },
        tabColor: { type: 'string', description: '标签颜色（如 FF0000）' }
      },
      required: ['workbookId', 'sheetName']
    }
  },
  {
    name: 'excel_write_data',
    description: '写入表格数据',
    inputSchema: {
      type: 'object',
      properties: {
        workbookId: { type: 'string', description: '工作簿ID' },
        sheetName: { type: 'string', description: '工作表名称' },
        headers: { type: 'array', items: { type: 'string' }, description: '表头' },
        rows: { type: 'array', items: { type: 'array' }, description: '数据行' },
        startCell: { type: 'string', description: '起始单元格，默认A1' },
        autoFilter: { type: 'boolean', description: '是否添加筛选' },
        freezeHeader: { type: 'boolean', description: '是否冻结首行' },
        style: {
          type: 'string',
          enum: ['striped', 'bordered', 'professional', 'minimal'],
          description: '表格样式'
        }
      },
      required: ['workbookId', 'sheetName', 'headers', 'rows']
    }
  },
  {
    name: 'excel_add_chart',
    description: '写入图表数据到指定位置（exceljs 不支持直接创建图表对象，数据写入后可在 Excel 中手动创建图表）',
    inputSchema: {
      type: 'object',
      properties: {
        workbookId: { type: 'string', description: '工作簿ID' },
        sheetName: { type: 'string', description: '工作表名称' },
        chartType: {
          type: 'string',
          enum: ['bar', 'column', 'line', 'pie', 'doughnut', 'area'],
          description: '图表类型'
        },
        title: { type: 'string', description: '图表标题' },
        categories: { type: 'array', items: { type: 'string' }, description: '分类标签' },
        series: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              name: { type: 'string' },
              values: { type: 'array', items: { type: 'number' } }
            }
          },
          description: '数据系列'
        },
        position: { type: 'string', description: '图表位置，如 F2' }
      },
      required: ['workbookId', 'sheetName', 'chartType', 'categories', 'series']
    }
  },
  {
    name: 'excel_add_formula',
    description: '添加公式',
    inputSchema: {
      type: 'object',
      properties: {
        workbookId: { type: 'string', description: '工作簿ID' },
        sheetName: { type: 'string', description: '工作表名称' },
        cell: { type: 'string', description: '单元格位置' },
        formula: { type: 'string', description: '公式，如 SUM(A1:A10)' }
      },
      required: ['workbookId', 'sheetName', 'cell', 'formula']
    }
  },
  {
    name: 'excel_set_column_width',
    description: '设置列宽',
    inputSchema: {
      type: 'object',
      properties: {
        workbookId: { type: 'string', description: '工作簿ID' },
        sheetName: { type: 'string', description: '工作表名称' },
        columns: { type: 'object', description: '列宽映射，如 { "A": 20, "B": 30 }' }
      },
      required: ['workbookId', 'sheetName', 'columns']
    }
  },
  {
    name: 'excel_merge_cells',
    description: '合并单元格',
    inputSchema: {
      type: 'object',
      properties: {
        workbookId: { type: 'string', description: '工作簿ID' },
        sheetName: { type: 'string', description: '工作表名称' },
        range: { type: 'string', description: '单元格范围，如 A1:C1' }
      },
      required: ['workbookId', 'sheetName', 'range']
    }
  },
  {
    name: 'excel_save',
    description: '保存工作簿',
    inputSchema: {
      type: 'object',
      properties: {
        workbookId: { type: 'string', description: '工作簿ID' }
      },
      required: ['workbookId']
    }
  }
]

// ============ 工具调用分发 ============

export async function handleExcelTool(name: string, args: Record<string, unknown>) {
  try {
    let result

    switch (name) {
      case 'excel_create':
        result = excelCreate(args)
        break
      case 'excel_add_sheet':
        result = excelAddSheet(args)
        break
      case 'excel_write_data':
        result = excelWriteData(args)
        break
      case 'excel_add_chart':
        result = excelAddChart(args)
        break
      case 'excel_add_formula':
        result = excelAddFormula(args)
        break
      case 'excel_set_column_width':
        result = excelSetColumnWidth(args)
        break
      case 'excel_merge_cells':
        result = excelMergeCells(args)
        break
      case 'excel_save':
        result = await excelSave(args)
        break
      default:
        result = { success: false, message: `未知的 Excel 工具: ${name}` }
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
