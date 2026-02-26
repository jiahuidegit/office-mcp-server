/**
 * 学术论文风格
 */

import { Theme } from './types.js'

export const academicTheme: Theme = {
  name: 'academic',
  displayName: '学术论文',
  description: '符合学术论文规范，适用于研究报告、毕业论文',
  colors: {
    primary: '000000', secondary: '333333', accent: '1a73e8', background: 'ffffff',
    headerBg: 'f8f9fa', border: 'dadce0', text: '000000', textSecondary: '5f6368'
  },
  pageLayout: {
    marginTop: Math.round(2.5 * 567), marginBottom: Math.round(2.5 * 567),
    marginLeft: Math.round(3.0 * 567), marginRight: Math.round(2.5 * 567)
  },
  word: {
    fonts: { title: '黑体', heading: '黑体', body: '宋体', code: 'Consolas' },
    heading1: { font: '黑体', color: '000000', fontSize: 30, bold: true, alignment: 'center', spaceBefore: 24, spaceAfter: 18 },
    heading2: { font: '黑体', color: '000000', fontSize: 28, bold: false, alignment: 'left', spaceBefore: 12, spaceAfter: 6 },
    heading3: { font: '黑体', color: '000000', fontSize: 24, bold: false, alignment: 'left', spaceBefore: 6, spaceAfter: 6 },
    paragraph: {
      font: '宋体', color: '000000', fontSize: 24,
      format: { lineSpacing: 1.5, lineSpacingType: 'multiple', firstLineIndent: 2, spaceBefore: 0, spaceAfter: 0 }
    },
    table: {
      headerBg: 'f8f9fa', headerColor: '000000', headerFont: '黑体', headerFontSize: 21,
      headerBold: true, bodyFont: '宋体', bodyFontSize: 21, stripedBg: 'f8f9fa',
      borderColor: '000000', borderWidth: 1
    },
    quote: { borderLeft: '1a73e8', borderWidth: 3, background: 'e8f0fe', font: '楷体', fontSize: 24, italic: true },
    code: { background: 'f8f9fa', font: 'Consolas', fontSize: 20, borderColor: 'dadce0' },
    list: { font: '宋体', fontSize: 24, indent: 480, spacing: 6 }
  },
  excel: {
    fonts: { title: '黑体', heading: '黑体', body: '宋体', code: 'Consolas' },
    headerRow: { fill: 'f8f9fa', fontColor: '000000', font: '黑体', fontSize: 10, bold: true, alignment: 'center' },
    dataRow: { fill: 'ffffff', font: '宋体', fontSize: 10 },
    alternateRow: { fill: 'f8f9fa' },
    border: { color: '000000', style: 'thin' },
    chartColors: ['1a73e8', '34a853', 'fbbc04', 'ea4335', '673ab7']
  }
}
