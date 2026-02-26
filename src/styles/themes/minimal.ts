/**
 * 简约黑白风格
 */

import { Theme } from './types.js'
import { createTheme } from './base.js'

export const minimalTheme: Theme = createTheme({
  name: 'minimal',
  displayName: '简约黑白',
  description: '简洁的黑白配色，通用于各类正式文档',
  colors: {
    primary: '000000', secondary: '666666', accent: '000000', background: 'ffffff',
    headerBg: 'f5f5f5', border: 'cccccc', text: '000000', textSecondary: '666666'
  },
  word: {
    fonts: { title: '黑体', heading: '黑体', body: '宋体', code: 'Consolas' },
    heading1: { font: '黑体', color: '000000', fontSize: 36, bold: true, alignment: 'center', spaceBefore: 24, spaceAfter: 18 },
    heading2: { font: '黑体', color: '000000', fontSize: 30, bold: false, alignment: 'left', spaceBefore: 12, spaceAfter: 6 },
    heading3: { font: '黑体', color: '000000', fontSize: 26, bold: false, alignment: 'left', spaceBefore: 6, spaceAfter: 6 },
    paragraph: {
      font: '宋体', color: '000000', fontSize: 24,
      format: { lineSpacing: 1.5, lineSpacingType: 'multiple', firstLineIndent: 2, spaceBefore: 0, spaceAfter: 0 }
    },
    table: {
      headerBg: 'f5f5f5', headerColor: '000000', headerFont: '黑体', headerFontSize: 22,
      headerBold: true, bodyFont: '宋体', bodyFontSize: 21, stripedBg: 'fafafa',
      borderColor: '000000', borderWidth: 1
    },
    quote: { borderLeft: '000000', borderWidth: 3, background: 'f5f5f5', font: '楷体', fontSize: 24, italic: true },
    code: { background: 'f5f5f5', font: 'Consolas', fontSize: 20, borderColor: 'cccccc' },
    list: { font: '宋体', fontSize: 24, indent: 480, spacing: 6 }
  },
  excel: {
    fonts: { title: '黑体', heading: '黑体', body: '宋体', code: 'Consolas' },
    headerRow: { fill: 'f5f5f5', fontColor: '000000', font: '黑体', fontSize: 11, bold: true, alignment: 'center' },
    dataRow: { fill: 'ffffff', font: '宋体', fontSize: 10 },
    alternateRow: { fill: 'fafafa' },
    border: { color: '000000', style: 'thin' },
    chartColors: ['000000', '333333', '666666', '999999', 'cccccc']
  }
})
