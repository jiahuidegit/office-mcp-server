/**
 * 字节风格
 */

import { Theme } from './types.js'
import { createTheme } from './base.js'

export const bytedanceTheme: Theme = createTheme({
  name: 'bytedance',
  displayName: '字节风格',
  description: '字节黑色系，适用于科技、媒体行业',
  colors: {
    primary: 'fe2c55', secondary: '161823', accent: '25f4ee', background: 'f8f8f8',
    headerBg: '161823', border: 'e1e1e1', text: '161823', textSecondary: '86868b'
  },
  word: {
    fonts: { title: '抖音美好体', heading: '抖音美好体', body: '微软雅黑', code: 'Consolas' },
    heading1: { font: '抖音美好体', color: '161823', fontSize: 40, bold: true, alignment: 'left', spaceBefore: 24, spaceAfter: 18 },
    heading2: { font: '抖音美好体', color: '161823', fontSize: 32, bold: true, alignment: 'left', spaceBefore: 18, spaceAfter: 12 },
    heading3: { font: '微软雅黑', color: '161823', fontSize: 28, bold: true, alignment: 'left', spaceBefore: 12, spaceAfter: 6 },
    paragraph: {
      font: '微软雅黑', color: '161823', fontSize: 24,
      format: { lineSpacing: 1.8, lineSpacingType: 'multiple', firstLineIndent: 0, spaceBefore: 0, spaceAfter: 10 }
    },
    table: {
      headerBg: '161823', headerColor: 'ffffff', headerFont: '微软雅黑', headerFontSize: 22,
      headerBold: true, bodyFont: '微软雅黑', bodyFontSize: 21, stripedBg: 'f8f8f8',
      borderColor: 'e1e1e1', borderWidth: 1
    },
    quote: { borderLeft: 'fe2c55', borderWidth: 4, background: 'fff0f3', font: '微软雅黑', fontSize: 24, italic: false },
    code: { background: 'f8f8f8', font: 'Consolas', fontSize: 20, borderColor: 'e1e1e1' },
    list: { font: '微软雅黑', fontSize: 24, indent: 480, spacing: 8 }
  },
  excel: {
    fonts: { title: '微软雅黑', heading: '微软雅黑', body: '微软雅黑', code: 'Consolas' },
    headerRow: { fill: '161823', fontColor: 'ffffff', font: '微软雅黑', fontSize: 11, bold: true, alignment: 'center' },
    dataRow: { fill: 'ffffff', font: '微软雅黑', fontSize: 10 },
    alternateRow: { fill: 'f8f8f8' },
    border: { color: 'e1e1e1', style: 'thin' },
    chartColors: ['fe2c55', '25f4ee', '161823', 'ee1d52', '69c9d0']
  }
})
