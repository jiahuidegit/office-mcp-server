/**
 * 基础测试 - 确保模块可以正常导入
 */

import { describe, it, expect } from 'vitest'
import { wordTools } from './tools/word/index.js'
import { excelTools } from './tools/excel/index.js'

describe('Office MCP Server', () => {
  describe('Word Tools', () => {
    it('should export word tools', () => {
      expect(wordTools).toBeDefined()
      expect(Array.isArray(wordTools)).toBe(true)
      expect(wordTools.length).toBeGreaterThan(0)
    })

    it('should have word_create tool', () => {
      const createTool = wordTools.find(t => t.name === 'word_create')
      expect(createTool).toBeDefined()
      expect(createTool?.description).toContain('Word')
    })
  })

  describe('Excel Tools', () => {
    it('should export excel tools', () => {
      expect(excelTools).toBeDefined()
      expect(Array.isArray(excelTools)).toBe(true)
      expect(excelTools.length).toBeGreaterThan(0)
    })

    it('should have excel_create tool', () => {
      const createTool = excelTools.find(t => t.name === 'excel_create')
      expect(createTool).toBeDefined()
      expect(createTool?.description).toContain('Excel')
    })
  })
})
