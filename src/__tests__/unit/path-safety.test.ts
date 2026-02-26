/**
 * path-safety 单元测试
 */

import { describe, it, expect } from 'vitest'
import { validateOutputPath } from '../../utils/path-safety.js'

describe('validateOutputPath', () => {
  it('合法的 word 路径', () => {
    expect(validateOutputPath('/tmp/test.docx', 'word')).toEqual({ valid: true })
    expect(validateOutputPath('output/report.docx', 'word')).toEqual({ valid: true })
  })

  it('合法的 excel 路径', () => {
    expect(validateOutputPath('/tmp/data.xlsx', 'excel')).toEqual({ valid: true })
  })

  it('拒绝路径穿越', () => {
    const result = validateOutputPath('../etc/passwd.docx', 'word')
    expect(result.valid).toBe(false)
    expect(result.message).toContain('..')
  })

  it('拒绝错误扩展名', () => {
    const result = validateOutputPath('/tmp/test.txt', 'word')
    expect(result.valid).toBe(false)
    expect(result.message).toContain('扩展名')
  })

  it('拒绝无扩展名', () => {
    const result = validateOutputPath('/tmp/test', 'excel')
    expect(result.valid).toBe(false)
  })
})
