/**
 * Excel 工作簿完整流程集成测试
 */

import { describe, it, expect, afterEach } from 'vitest'
import * as fs from 'fs'
import * as path from 'path'
import * as os from 'os'
import { handleExcelTool } from '../../tools/excel/index.js'

const tmpDir = path.join(os.tmpdir(), 'office-mcp-test-excel')

function extractResult(response: { content: Array<{ text: string }> }) {
  return JSON.parse(response.content[0].text)
}

afterEach(() => {
  if (fs.existsSync(tmpDir)) {
    fs.rmSync(tmpDir, { recursive: true })
  }
})

describe('Excel 完整流程', () => {
  it('创建 -> 写入数据 -> 保存', async () => {
    const outputPath = path.join(tmpDir, 'test-flow.xlsx')

    // 1. 创建工作簿
    const createRes = extractResult(await handleExcelTool('excel_create', {
      title: '测试工作簿',
      outputPath,
      theme: 'default'
    }))
    expect(createRes.success).toBe(true)
    const workbookId = createRes.workbookId
    expect(workbookId).toMatch(/^excel_/)

    // 2. 添加工作表
    const sheetRes = extractResult(await handleExcelTool('excel_add_sheet', {
      workbookId, sheetName: '销售数据', tabColor: 'FF0000'
    }))
    expect(sheetRes.success).toBe(true)

    // 3. 写入数据
    const writeRes = extractResult(await handleExcelTool('excel_write_data', {
      workbookId,
      sheetName: '销售数据',
      headers: ['产品', 'Q1', 'Q2', 'Q3', 'Q4'],
      rows: [
        ['产品A', 100, 150, 200, 180],
        ['产品B', 80, 120, 160, 140]
      ],
      autoFilter: true,
      freezeHeader: true,
      style: 'striped'
    }))
    expect(writeRes.success).toBe(true)

    // 4. 保存
    const saveRes = extractResult(await handleExcelTool('excel_save', { workbookId }))
    expect(saveRes.success).toBe(true)
    expect(fs.existsSync(outputPath)).toBe(true)

    const stat = fs.statSync(outputPath)
    expect(stat.size).toBeGreaterThan(1024)
  })

  it('无效 workbookId 返回错误', async () => {
    const res = extractResult(await handleExcelTool('excel_add_sheet', {
      workbookId: 'not_exist', sheetName: '测试'
    }))
    expect(res.success).toBe(false)
    expect(res.message).toContain('不存在')
  })

  it('参数校验失败返回友好错误', async () => {
    const res = extractResult(await handleExcelTool('excel_create', {
      title: '测试'
      // 缺少 outputPath
    }))
    expect(res.success).toBe(false)
    expect(res.message).toContain('参数校验失败')
  })

  it('公式、图表数据、合并单元格、列宽', async () => {
    const outputPath = path.join(tmpDir, 'full-features.xlsx')

    const createRes = extractResult(await handleExcelTool('excel_create', {
      title: '全功能测试', outputPath
    }))
    const workbookId = createRes.workbookId

    // 添加工作表
    extractResult(await handleExcelTool('excel_add_sheet', {
      workbookId, sheetName: '数据表'
    }))

    // 写入数据
    extractResult(await handleExcelTool('excel_write_data', {
      workbookId, sheetName: '数据表',
      headers: ['项目', '金额'],
      rows: [['A', 100], ['B', 200], ['C', 300]]
    }))

    // 添加公式
    const formulaRes = extractResult(await handleExcelTool('excel_add_formula', {
      workbookId, sheetName: '数据表', cell: 'B5', formula: 'SUM(B2:B4)'
    }))
    expect(formulaRes.success).toBe(true)

    // 写入图表数据
    const chartRes = extractResult(await handleExcelTool('excel_add_chart', {
      workbookId, sheetName: '数据表',
      chartType: 'bar',
      categories: ['A', 'B', 'C'],
      series: [{ name: '金额', values: [100, 200, 300] }],
      position: 'D2'
    }))
    expect(chartRes.success).toBe(true)

    // 合并单元格
    const mergeRes = extractResult(await handleExcelTool('excel_merge_cells', {
      workbookId, sheetName: '数据表', range: 'D1:F1'
    }))
    expect(mergeRes.success).toBe(true)

    // 设置列宽
    const colRes = extractResult(await handleExcelTool('excel_set_column_width', {
      workbookId, sheetName: '数据表', columns: { A: 20, B: 15 }
    }))
    expect(colRes.success).toBe(true)

    // 保存
    const saveRes = extractResult(await handleExcelTool('excel_save', { workbookId }))
    expect(saveRes.success).toBe(true)
    expect(fs.existsSync(outputPath)).toBe(true)
  })

  it('各主题均可正常创建和保存', async () => {
    const themes = ['alibaba', 'tencent', 'bytedance', 'default']

    for (const theme of themes) {
      const outputPath = path.join(tmpDir, `theme-${theme}.xlsx`)
      const createRes = extractResult(await handleExcelTool('excel_create', {
        title: `${theme} 主题测试`, outputPath, theme
      }))
      expect(createRes.success).toBe(true)

      extractResult(await handleExcelTool('excel_add_sheet', {
        workbookId: createRes.workbookId, sheetName: '测试'
      }))

      extractResult(await handleExcelTool('excel_write_data', {
        workbookId: createRes.workbookId, sheetName: '测试',
        headers: ['A', 'B'], rows: [['1', '2']]
      }))

      const saveRes = extractResult(await handleExcelTool('excel_save', {
        workbookId: createRes.workbookId
      }))
      expect(saveRes.success).toBe(true)
      expect(fs.existsSync(outputPath)).toBe(true)
    }
  })
})
