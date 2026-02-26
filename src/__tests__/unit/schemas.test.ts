/**
 * Schema 校验测试
 */

import { describe, it, expect } from 'vitest'
import {
  wordCreateSchema,
  wordAddHeadingSchema,
  wordAddParagraphSchema,
  wordAddTableSchema,
  wordAddListSchema,
  wordAddCodeSchema,
  wordAddDiagramSchema,
  wordSaveSchema
} from '../../schemas/word.schemas.js'
import {
  excelCreateSchema,
  excelAddSheetSchema,
  excelWriteDataSchema,
  excelAddChartSchema,
  excelAddFormulaSchema,
  excelSetColumnWidthSchema,
  excelMergeCellsSchema,
  excelSaveSchema
} from '../../schemas/excel.schemas.js'

// ============ Word Schema 测试 ============

describe('Word Schemas', () => {
  describe('wordCreateSchema', () => {
    it('合法输入 - 最小参数', () => {
      const result = wordCreateSchema.parse({ title: '测试', outputPath: '/tmp/test.docx' })
      expect(result.title).toBe('测试')
      expect(result.theme).toBe('default')
    })

    it('合法输入 - 全部参数', () => {
      const result = wordCreateSchema.parse({
        title: '测试',
        template: 'tech-doc',
        theme: 'government',
        author: '张三',
        outputPath: '/tmp/test.docx'
      })
      expect(result.template).toBe('tech-doc')
      expect(result.theme).toBe('government')
    })

    it('非法输入 - 缺少 title', () => {
      expect(() => wordCreateSchema.parse({ outputPath: '/tmp/test.docx' })).toThrow()
    })

    it('非法输入 - 空 title', () => {
      expect(() => wordCreateSchema.parse({ title: '', outputPath: '/tmp/test.docx' })).toThrow()
    })

    it('非法输入 - 无效 theme', () => {
      expect(() => wordCreateSchema.parse({
        title: '测试', outputPath: '/tmp/test.docx', theme: 'invalid'
      })).toThrow()
    })

    it('非法输入 - 无效 template', () => {
      expect(() => wordCreateSchema.parse({
        title: '测试', outputPath: '/tmp/test.docx', template: 'nonexistent'
      })).toThrow()
    })
  })

  describe('wordAddHeadingSchema', () => {
    it('合法输入', () => {
      const result = wordAddHeadingSchema.parse({ docId: 'doc_1', text: '标题', level: 1 })
      expect(result.level).toBe(1)
    })

    it('非法输入 - level 超出范围', () => {
      expect(() => wordAddHeadingSchema.parse({ docId: 'doc_1', text: '标题', level: 7 })).toThrow()
    })

    it('非法输入 - level 为 0', () => {
      expect(() => wordAddHeadingSchema.parse({ docId: 'doc_1', text: '标题', level: 0 })).toThrow()
    })
  })

  describe('wordAddParagraphSchema', () => {
    it('合法输入 - 最小参数', () => {
      const result = wordAddParagraphSchema.parse({ docId: 'doc_1', text: '内容' })
      expect(result.style).toBe('normal')
      expect(result.bold).toBe(false)
    })

    it('非法输入 - 无效 style', () => {
      expect(() => wordAddParagraphSchema.parse({
        docId: 'doc_1', text: '内容', style: 'invalid'
      })).toThrow()
    })
  })

  describe('wordAddTableSchema', () => {
    it('合法输入', () => {
      const result = wordAddTableSchema.parse({
        docId: 'doc_1', headers: ['名称', '值'], rows: [['a', '1']]
      })
      expect(result.style).toBe('professional')
    })

    it('非法输入 - 空 headers', () => {
      expect(() => wordAddTableSchema.parse({ docId: 'doc_1', headers: [], rows: [] })).toThrow()
    })
  })

  describe('wordAddListSchema', () => {
    it('合法输入', () => {
      const result = wordAddListSchema.parse({ docId: 'doc_1', items: ['项目1'] })
      expect(result.ordered).toBe(false)
    })

    it('非法输入 - 空 items', () => {
      expect(() => wordAddListSchema.parse({ docId: 'doc_1', items: [] })).toThrow()
    })
  })

  describe('wordAddCodeSchema', () => {
    it('合法输入', () => {
      const result = wordAddCodeSchema.parse({ docId: 'doc_1', code: 'console.log()' })
      expect(result.language).toBe('')
    })
  })

  describe('wordAddDiagramSchema', () => {
    it('合法输入', () => {
      const result = wordAddDiagramSchema.parse({ docId: 'doc_1', mermaid: 'graph TD; A-->B' })
      expect(result.width).toBe(15)
      expect(result.theme).toBe('default')
    })

    it('非法输入 - 空 mermaid', () => {
      expect(() => wordAddDiagramSchema.parse({ docId: 'doc_1', mermaid: '' })).toThrow()
    })
  })

  describe('wordSaveSchema', () => {
    it('合法输入 - 最小参数', () => {
      const result = wordSaveSchema.parse({ docId: 'doc_1' })
      expect(result.addPageNumbers).toBe(false)
    })

    it('合法输入 - 全部参数', () => {
      const result = wordSaveSchema.parse({
        docId: 'doc_1', addPageNumbers: true, addHeader: '页眉', addFooter: '页脚'
      })
      expect(result.addPageNumbers).toBe(true)
    })
  })
})

// ============ Excel Schema 测试 ============

describe('Excel Schemas', () => {
  describe('excelCreateSchema', () => {
    it('合法输入 - 最小参数', () => {
      const result = excelCreateSchema.parse({ title: '测试', outputPath: '/tmp/test.xlsx' })
      expect(result.theme).toBe('default')
    })

    it('合法输入 - 带 template', () => {
      const result = excelCreateSchema.parse({
        title: '测试', outputPath: '/tmp/test.xlsx', template: 'data-report'
      })
      expect(result.template).toBe('data-report')
    })

    it('非法输入 - 缺少 outputPath', () => {
      expect(() => excelCreateSchema.parse({ title: '测试' })).toThrow()
    })

    it('非法输入 - 无效 template', () => {
      expect(() => excelCreateSchema.parse({
        title: '测试', outputPath: '/tmp/test.xlsx', template: 'invalid'
      })).toThrow()
    })
  })

  describe('excelAddSheetSchema', () => {
    it('合法输入', () => {
      const result = excelAddSheetSchema.parse({ workbookId: 'wb_1', sheetName: 'Sheet1' })
      expect(result.sheetName).toBe('Sheet1')
    })

    it('非法输入 - 空 sheetName', () => {
      expect(() => excelAddSheetSchema.parse({ workbookId: 'wb_1', sheetName: '' })).toThrow()
    })
  })

  describe('excelWriteDataSchema', () => {
    it('合法输入 - 最小参数', () => {
      const result = excelWriteDataSchema.parse({
        workbookId: 'wb_1', sheetName: 'Sheet1',
        headers: ['A', 'B'], rows: [[1, 2]]
      })
      expect(result.startCell).toBe('A1')
      expect(result.style).toBe('professional')
    })

    it('非法输入 - 无效 startCell', () => {
      expect(() => excelWriteDataSchema.parse({
        workbookId: 'wb_1', sheetName: 'Sheet1',
        headers: ['A'], rows: [], startCell: 'abc'
      })).toThrow()
    })
  })

  describe('excelAddChartSchema', () => {
    it('合法输入', () => {
      const result = excelAddChartSchema.parse({
        workbookId: 'wb_1', sheetName: 'Sheet1',
        chartType: 'bar', categories: ['Q1'],
        series: [{ name: '销售', values: [100] }]
      })
      expect(result.position).toBe('F2')
    })

    it('非法输入 - 无效 chartType', () => {
      expect(() => excelAddChartSchema.parse({
        workbookId: 'wb_1', sheetName: 'Sheet1',
        chartType: 'radar', categories: ['Q1'],
        series: [{ name: '销售', values: [100] }]
      })).toThrow()
    })

    it('非法输入 - 空 series', () => {
      expect(() => excelAddChartSchema.parse({
        workbookId: 'wb_1', sheetName: 'Sheet1',
        chartType: 'bar', categories: ['Q1'], series: []
      })).toThrow()
    })
  })

  describe('excelAddFormulaSchema', () => {
    it('合法输入', () => {
      const result = excelAddFormulaSchema.parse({
        workbookId: 'wb_1', sheetName: 'Sheet1', cell: 'A1', formula: 'SUM(B1:B10)'
      })
      expect(result.formula).toBe('SUM(B1:B10)')
    })

    it('非法输入 - 无效 cell 格式', () => {
      expect(() => excelAddFormulaSchema.parse({
        workbookId: 'wb_1', sheetName: 'Sheet1', cell: 'abc', formula: 'SUM(A1:A10)'
      })).toThrow()
    })
  })

  describe('excelSetColumnWidthSchema', () => {
    it('合法输入', () => {
      const result = excelSetColumnWidthSchema.parse({
        workbookId: 'wb_1', sheetName: 'Sheet1', columns: { A: 20, B: 30 }
      })
      expect(result.columns.A).toBe(20)
    })
  })

  describe('excelMergeCellsSchema', () => {
    it('合法输入', () => {
      const result = excelMergeCellsSchema.parse({
        workbookId: 'wb_1', sheetName: 'Sheet1', range: 'A1:C1'
      })
      expect(result.range).toBe('A1:C1')
    })

    it('非法输入 - 空 range', () => {
      expect(() => excelMergeCellsSchema.parse({
        workbookId: 'wb_1', sheetName: 'Sheet1', range: ''
      })).toThrow()
    })
  })

  describe('excelSaveSchema', () => {
    it('合法输入', () => {
      const result = excelSaveSchema.parse({ workbookId: 'wb_1' })
      expect(result.workbookId).toBe('wb_1')
    })

    it('非法输入 - 空 workbookId', () => {
      expect(() => excelSaveSchema.parse({ workbookId: '' })).toThrow()
    })
  })
})
