/**
 * 软件文档风格 (GB/T 8567-2006)
 */

import { Theme } from './types.js'
import { createTheme } from './base.js'

export const softwareDocTheme: Theme = createTheme({
  name: 'software',
  displayName: '软件文档',
  description: '符合GB/T 8567-2006规范，适用于需求规格说明书、设计文档、技术方案',
  colors: {
    primary: '000000', secondary: '333333', accent: '1a73e8', background: 'ffffff',
    headerBg: 'f5f5f5', border: '000000', text: '000000', textSecondary: '666666'
  },
  word: {
    fonts: { title: '黑体', heading: '黑体', body: '宋体', code: 'Consolas' },
    heading1: { font: '黑体', color: '000000', fontSize: 44, bold: true, alignment: 'center', spaceBefore: 24, spaceAfter: 24 },
    heading2: { font: '黑体', color: '000000', fontSize: 32, bold: false, alignment: 'left', spaceBefore: 12, spaceAfter: 6 },
    heading3: { font: '黑体', color: '000000', fontSize: 28, bold: false, alignment: 'left', spaceBefore: 6, spaceAfter: 6 },
    paragraph: {
      font: '宋体', color: '000000', fontSize: 24,
      format: { lineSpacing: 1.5, lineSpacingType: 'multiple', firstLineIndent: 2, spaceBefore: 0, spaceAfter: 0 }
    },
    table: {
      headerBg: 'f5f5f5', headerColor: '000000', headerFont: '黑体', headerFontSize: 21,
      headerBold: true, bodyFont: '宋体', bodyFontSize: 21, stripedBg: 'fafafa',
      borderColor: '000000', borderWidth: 1
    },
    quote: { borderLeft: '1a73e8', borderWidth: 3, background: 'f5f5f5', font: '楷体', fontSize: 24, italic: false },
    code: { background: 'f5f5f5', font: 'Consolas', fontSize: 20, borderColor: 'cccccc' },
    list: { font: '宋体', fontSize: 24, indent: 480, spacing: 0 }
  },
  excel: {
    fonts: { title: '黑体', heading: '黑体', body: '宋体', code: 'Consolas' },
    headerRow: { fill: 'f5f5f5', fontColor: '000000', font: '黑体', fontSize: 10, bold: true, alignment: 'center' },
    dataRow: { fill: 'ffffff', font: '宋体', fontSize: 10 },
    alternateRow: { fill: 'fafafa' },
    border: { color: '000000', style: 'thin' },
    chartColors: ['1a73e8', '34a853', 'fbbc04', 'ea4335', '673ab7']
  }
})
