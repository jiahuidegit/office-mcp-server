/**
 * id 生成器单元测试
 */

import { describe, it, expect } from 'vitest'
import { generateId } from '../../utils/id.js'

describe('generateId', () => {
  it('包含前缀', () => {
    expect(generateId('word')).toMatch(/^word_/)
    expect(generateId('excel')).toMatch(/^excel_/)
  })

  it('每次生成不同 ID', () => {
    const ids = new Set(Array.from({ length: 100 }, () => generateId('test')))
    expect(ids.size).toBe(100)
  })

  it('包含时间戳和随机字符', () => {
    const id = generateId('doc')
    const parts = id.split('_')
    expect(parts.length).toBe(3)
    expect(Number(parts[1])).toBeGreaterThan(0)
    expect(parts[2].length).toBe(7)
  })
})
