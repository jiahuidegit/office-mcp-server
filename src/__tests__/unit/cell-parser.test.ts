/**
 * cell-parser 单元测试
 */

import { describe, it, expect } from 'vitest'
import { columnLetterToNumber, columnNumberToLetter, parseCellAddress } from '../../utils/cell-parser.js'

describe('columnLetterToNumber', () => {
  it('单字母列', () => {
    expect(columnLetterToNumber('A')).toBe(1)
    expect(columnLetterToNumber('Z')).toBe(26)
  })

  it('多字母列', () => {
    expect(columnLetterToNumber('AA')).toBe(27)
    expect(columnLetterToNumber('AZ')).toBe(52)
    expect(columnLetterToNumber('BA')).toBe(53)
  })
})

describe('columnNumberToLetter', () => {
  it('单字母列', () => {
    expect(columnNumberToLetter(1)).toBe('A')
    expect(columnNumberToLetter(26)).toBe('Z')
  })

  it('多字母列', () => {
    expect(columnNumberToLetter(27)).toBe('AA')
    expect(columnNumberToLetter(52)).toBe('AZ')
    expect(columnNumberToLetter(53)).toBe('BA')
  })

  it('与 columnLetterToNumber 互逆', () => {
    for (let i = 1; i <= 100; i++) {
      expect(columnLetterToNumber(columnNumberToLetter(i))).toBe(i)
    }
  })
})

describe('parseCellAddress', () => {
  it('解析标准地址', () => {
    expect(parseCellAddress('A1')).toEqual({ col: 1, row: 1 })
    expect(parseCellAddress('B10')).toEqual({ col: 2, row: 10 })
    expect(parseCellAddress('AA100')).toEqual({ col: 27, row: 100 })
  })

  it('无效地址返回默认值', () => {
    expect(parseCellAddress('abc')).toEqual({ col: 1, row: 1 })
    expect(parseCellAddress('')).toEqual({ col: 1, row: 1 })
  })
})
