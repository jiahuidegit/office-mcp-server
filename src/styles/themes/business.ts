/**
 * 商务报告风格
 */

import { Theme } from './types.js'
import { createTheme } from './base.js'

export const businessTheme: Theme = createTheme({
  name: 'business',
  displayName: '商务报告',
  description: '适用于市场调研报告、商业计划书、工作汇报',
  colors: {
    primary: '1890ff', secondary: '13c2c2', accent: '722ed1', background: 'f5f5f5',
    headerBg: 'e6f7ff', border: '91d5ff', text: '262626', textSecondary: '8c8c8c'
  },
  word: {
    fonts: { title: '微软雅黑', heading: '微软雅黑', body: '微软雅黑', code: 'Consolas' },
    heading1: { font: '微软雅黑', color: '1890ff', fontSize: 36, bold: true, alignment: 'center', spaceBefore: 24, spaceAfter: 24 },
    heading2: { font: '微软雅黑', color: '262626', fontSize: 32, bold: true, alignment: 'left', spaceBefore: 18, spaceAfter: 12 },
    heading3: { font: '微软雅黑', color: '262626', fontSize: 28, bold: true, alignment: 'left', spaceBefore: 12, spaceAfter: 6 },
    paragraph: {
      font: '微软雅黑', color: '262626', fontSize: 24,
      format: { lineSpacing: 1.5, lineSpacingType: 'multiple', firstLineIndent: 0, spaceBefore: 0, spaceAfter: 8 }
    },
    table: {
      headerBg: 'e6f7ff', headerColor: '1890ff', headerFont: '微软雅黑', headerFontSize: 22,
      headerBold: true, bodyFont: '微软雅黑', bodyFontSize: 21, stripedBg: 'fafafa',
      borderColor: '91d5ff', borderWidth: 1
    },
    quote: { borderLeft: '1890ff', borderWidth: 4, background: 'e6f7ff', font: '微软雅黑', fontSize: 24, italic: false },
    code: { background: 'f5f5f5', font: 'Consolas', fontSize: 20, borderColor: 'd9d9d9' },
    list: { font: '微软雅黑', fontSize: 24, indent: 480, spacing: 6 }
  },
  excel: {
    fonts: { title: '微软雅黑', heading: '微软雅黑', body: '微软雅黑', code: 'Consolas' },
    headerRow: { fill: 'e6f7ff', fontColor: '1890ff', font: '微软雅黑', fontSize: 11, bold: true, alignment: 'center' },
    dataRow: { fill: 'ffffff', font: '微软雅黑', fontSize: 10 },
    alternateRow: { fill: 'fafafa' },
    border: { color: '91d5ff', style: 'thin' },
    chartColors: ['1890ff', '52c41a', 'faad14', 'f5222d', '722ed1']
  }
})
