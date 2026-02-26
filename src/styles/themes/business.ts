/**
 * 商务报告风格
 * 标准中文文档格式：黑体标题 + 仿宋正文，黑色文字
 */

import { Theme } from './types.js'
import { createTheme } from './base.js'

export const businessTheme: Theme = createTheme({
  name: 'business',
  displayName: '商务报告',
  description: '适用于市场调研报告、商业计划书、工作汇报',
  colors: {
    primary: '000000', secondary: '333333', accent: '000000', background: 'ffffff',
    headerBg: 'f0f0f0', border: '000000', text: '000000', textSecondary: '333333'
  },
  word: {
    fonts: { title: '宋体', heading: '黑体', body: '仿宋', code: 'Consolas' },
    heading1: {
      font: '宋体', color: '000000', fontSize: 44,
      bold: true, alignment: 'center', spaceBefore: 24, spaceAfter: 24
    },
    heading2: {
      font: '黑体', color: '000000', fontSize: 32,
      bold: true, alignment: 'left', spaceBefore: 18, spaceAfter: 12
    },
    heading3: {
      font: '黑体', color: '000000', fontSize: 28,
      bold: true, alignment: 'left', spaceBefore: 12, spaceAfter: 6
    },
    paragraph: {
      font: '仿宋', color: '000000', fontSize: 28,
      format: { lineSpacing: 28, lineSpacingType: 'exact', firstLineIndent: 2, spaceBefore: 0, spaceAfter: 6 }
    },
    table: {
      headerBg: 'f0f0f0', headerColor: '000000', headerFont: '黑体', headerFontSize: 22,
      headerBold: true, bodyFont: '仿宋', bodyFontSize: 22, stripedBg: 'fafafa',
      borderColor: '000000', borderWidth: 1
    },
    quote: { borderLeft: '666666', borderWidth: 3, background: 'f5f5f5', font: '楷体', fontSize: 28, italic: false },
    code: { background: 'f5f5f5', font: 'Consolas', fontSize: 20, borderColor: 'cccccc' },
    list: { font: '仿宋', fontSize: 28, indent: 560, spacing: 4 }
  },
  excel: {
    fonts: { title: '黑体', heading: '黑体', body: '宋体', code: 'Consolas' },
    headerRow: { fill: 'f0f0f0', fontColor: '000000', font: '黑体', fontSize: 11, bold: true, alignment: 'center' },
    dataRow: { fill: 'ffffff', font: '宋体', fontSize: 10 },
    alternateRow: { fill: 'fafafa' },
    border: { color: '000000', style: 'thin' },
    chartColors: ['333333', '666666', '999999', 'bbbbbb', 'dddddd']
  }
})
