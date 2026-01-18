/**
 * Excel 工具模块 - 完整实现
 */

import { Tool } from '@modelcontextprotocol/sdk/types.js'
import ExcelJS from 'exceljs'
import * as fs from 'fs'
import * as path from 'path'
import { getTheme, Theme } from '../../styles/themes/index.js'

// 工作簿存储
interface WorkbookEntry {
  workbook: ExcelJS.Workbook
  outputPath: string
  theme: Theme
  title: string
}

const workbooks = new Map<string, WorkbookEntry>()

function generateId(): string {
  return `excel_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`
}

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
    description: '添加图表',
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
        columns: {
          type: 'object',
          description: '列宽映射，如 { "A": 20, "B": 30 }'
        }
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

// ============ 工具实现 ============

// 创建工作簿
function excelCreate(args: Record<string, unknown>) {
  const { title, theme = 'default', outputPath } = args as {
    title: string
    theme?: string
    outputPath: string
  }

  const id = generateId()
  const themeConfig = getTheme(theme)
  const workbook = new ExcelJS.Workbook()

  workbook.creator = 'Office MCP Server'
  workbook.created = new Date()
  workbook.title = title

  workbooks.set(id, {
    workbook,
    outputPath,
    theme: themeConfig,
    title
  })

  return {
    success: true,
    workbookId: id,
    message: `工作簿 "${title}" 创建成功，主题: ${themeConfig.displayName}`
  }
}

// 添加工作表
function excelAddSheet(args: Record<string, unknown>) {
  const { workbookId, sheetName, tabColor } = args as {
    workbookId: string
    sheetName: string
    tabColor?: string
  }

  const entry = workbooks.get(workbookId)
  if (!entry) {
    return { success: false, message: `工作簿 ${workbookId} 不存在` }
  }

  const sheet = entry.workbook.addWorksheet(sheetName)

  if (tabColor) {
    sheet.properties.tabColor = { argb: `FF${tabColor}` }
  }

  return { success: true, message: `已添加工作表: ${sheetName}` }
}

// 写入数据
function excelWriteData(args: Record<string, unknown>) {
  const {
    workbookId,
    sheetName,
    headers,
    rows,
    startCell = 'A1',
    autoFilter = false,
    freezeHeader = false,
    style = 'professional'
  } = args as {
    workbookId: string
    sheetName: string
    headers: string[]
    rows: unknown[][]
    startCell?: string
    autoFilter?: boolean
    freezeHeader?: boolean
    style?: string
  }

  const entry = workbooks.get(workbookId)
  if (!entry) {
    return { success: false, message: `工作簿 ${workbookId} 不存在` }
  }

  let sheet = entry.workbook.getWorksheet(sheetName)
  if (!sheet) {
    sheet = entry.workbook.addWorksheet(sheetName)
  }

  const themeStyle = entry.theme.excel

  // 解析起始单元格
  const startMatch = startCell.match(/^([A-Z]+)(\d+)$/)
  const startCol = startMatch ? startMatch[1].charCodeAt(0) - 64 : 1
  const startRow = startMatch ? parseInt(startMatch[2]) : 1

  // 写入表头
  const headerRow = sheet.getRow(startRow)
  headers.forEach((header, index) => {
    const cell = headerRow.getCell(startCol + index)
    cell.value = header
    cell.font = {
      bold: themeStyle.headerRow.bold,
      color: { argb: `FF${themeStyle.headerRow.fontColor}` }
    }
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: `FF${themeStyle.headerRow.fill}` }
    }
    cell.alignment = { horizontal: 'center', vertical: 'middle' }
    if (style !== 'minimal') {
      cell.border = {
        top: { style: 'thin', color: { argb: `FF${themeStyle.border.color}` } },
        bottom: { style: 'thin', color: { argb: `FF${themeStyle.border.color}` } },
        left: { style: 'thin', color: { argb: `FF${themeStyle.border.color}` } },
        right: { style: 'thin', color: { argb: `FF${themeStyle.border.color}` } }
      }
    }
  })
  headerRow.height = 25

  // 写入数据行
  rows.forEach((row, rowIndex) => {
    const dataRow = sheet.getRow(startRow + 1 + rowIndex)
    row.forEach((value, colIndex) => {
      const cell = dataRow.getCell(startCol + colIndex)
      cell.value = value as ExcelJS.CellValue

      // 斑马纹样式
      if (style === 'striped' && rowIndex % 2 === 1) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: `FF${themeStyle.alternateRow.fill}` }
        }
      }

      if (style !== 'minimal') {
        cell.border = {
          top: { style: 'thin', color: { argb: `FF${themeStyle.border.color}` } },
          bottom: { style: 'thin', color: { argb: `FF${themeStyle.border.color}` } },
          left: { style: 'thin', color: { argb: `FF${themeStyle.border.color}` } },
          right: { style: 'thin', color: { argb: `FF${themeStyle.border.color}` } }
        }
      }
    })
  })

  // 自动调整列宽
  headers.forEach((_, index) => {
    const column = sheet.getColumn(startCol + index)
    column.width = 15
  })

  // 冻结首行
  if (freezeHeader) {
    sheet.views = [{ state: 'frozen', ySplit: startRow }]
  }

  // 自动筛选
  if (autoFilter) {
    const endCol = String.fromCharCode(64 + startCol + headers.length - 1)
    const endRow = startRow + rows.length
    sheet.autoFilter = `${startCell}:${endCol}${endRow}`
  }

  return {
    success: true,
    message: `已写入 ${headers.length} 列 ${rows.length} 行数据`
  }
}

// 添加图表（注意：exceljs 对图表支持有限，这里提供基础实现）
function excelAddChart(args: Record<string, unknown>) {
  const { workbookId, sheetName, chartType, title, categories, series, position = 'F2' } = args as {
    workbookId: string
    sheetName: string
    chartType: string
    title?: string
    categories: string[]
    series: Array<{ name: string; values: number[] }>
    position?: string
  }

  const entry = workbooks.get(workbookId)
  if (!entry) {
    return { success: false, message: `工作簿 ${workbookId} 不存在` }
  }

  const sheet = entry.workbook.getWorksheet(sheetName)
  if (!sheet) {
    return { success: false, message: `工作表 ${sheetName} 不存在` }
  }

  // exceljs 图表支持有限，这里在单元格中写入图表数据作为替代方案
  // 实际图表需要在 Excel 中打开后手动创建，或使用其他库

  // 写入图表数据到指定位置
  const posMatch = position.match(/^([A-Z]+)(\d+)$/)
  const posCol = posMatch ? posMatch[1].charCodeAt(0) - 64 : 6
  const posRow = posMatch ? parseInt(posMatch[2]) : 2

  // 写入标题
  if (title) {
    const titleCell = sheet.getCell(posRow, posCol)
    titleCell.value = `📊 ${title}`
    titleCell.font = { bold: true, size: 14 }
  }

  // 写入分类和数据
  const dataStartRow = posRow + (title ? 2 : 1)

  // 表头行
  const headerRow = sheet.getRow(dataStartRow)
  headerRow.getCell(posCol).value = '分类'
  series.forEach((s, i) => {
    headerRow.getCell(posCol + 1 + i).value = s.name
  })

  // 数据行
  categories.forEach((cat, catIndex) => {
    const row = sheet.getRow(dataStartRow + 1 + catIndex)
    row.getCell(posCol).value = cat
    series.forEach((s, seriesIndex) => {
      row.getCell(posCol + 1 + seriesIndex).value = s.values[catIndex]
    })
  })

  return {
    success: true,
    message: `图表数据已写入（${chartType} 类型），请在 Excel 中选择数据创建图表`,
    note: 'exceljs 不支持直接创建图表，数据已准备好供手动创建图表使用'
  }
}

// 添加公式
function excelAddFormula(args: Record<string, unknown>) {
  const { workbookId, sheetName, cell, formula } = args as {
    workbookId: string
    sheetName: string
    cell: string
    formula: string
  }

  const entry = workbooks.get(workbookId)
  if (!entry) {
    return { success: false, message: `工作簿 ${workbookId} 不存在` }
  }

  const sheet = entry.workbook.getWorksheet(sheetName)
  if (!sheet) {
    return { success: false, message: `工作表 ${sheetName} 不存在` }
  }

  // 移除开头的 = 如果有的话
  const cleanFormula = formula.startsWith('=') ? formula.slice(1) : formula

  const targetCell = sheet.getCell(cell)
  targetCell.value = { formula: cleanFormula }

  return { success: true, message: `已在 ${cell} 添加公式: =${cleanFormula}` }
}

// 设置列宽
function excelSetColumnWidth(args: Record<string, unknown>) {
  const { workbookId, sheetName, columns } = args as {
    workbookId: string
    sheetName: string
    columns: Record<string, number>
  }

  const entry = workbooks.get(workbookId)
  if (!entry) {
    return { success: false, message: `工作簿 ${workbookId} 不存在` }
  }

  const sheet = entry.workbook.getWorksheet(sheetName)
  if (!sheet) {
    return { success: false, message: `工作表 ${sheetName} 不存在` }
  }

  Object.entries(columns).forEach(([col, width]) => {
    const column = sheet.getColumn(col)
    column.width = width
  })

  return { success: true, message: `已设置 ${Object.keys(columns).length} 列的宽度` }
}

// 合并单元格
function excelMergeCells(args: Record<string, unknown>) {
  const { workbookId, sheetName, range } = args as {
    workbookId: string
    sheetName: string
    range: string
  }

  const entry = workbooks.get(workbookId)
  if (!entry) {
    return { success: false, message: `工作簿 ${workbookId} 不存在` }
  }

  const sheet = entry.workbook.getWorksheet(sheetName)
  if (!sheet) {
    return { success: false, message: `工作表 ${sheetName} 不存在` }
  }

  sheet.mergeCells(range)

  return { success: true, message: `已合并单元格: ${range}` }
}

// 保存工作簿
async function excelSave(args: Record<string, unknown>) {
  const { workbookId } = args as { workbookId: string }

  const entry = workbooks.get(workbookId)
  if (!entry) {
    return { success: false, message: `工作簿 ${workbookId} 不存在` }
  }

  // 确保目录存在
  const dir = path.dirname(entry.outputPath)
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true })
  }

  // 写入文件
  await entry.workbook.xlsx.writeFile(entry.outputPath)

  // 清理内存
  workbooks.delete(workbookId)

  return {
    success: true,
    message: `工作簿已保存至: ${entry.outputPath}`,
    path: entry.outputPath
  }
}

// ============ 工具调用分发 ============

export async function handleExcelTool(name: string, args: Record<string, unknown>) {
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
}
