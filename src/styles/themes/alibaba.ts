/**
 * 阿里风格
 */

import { Theme } from './types.js'
import { createTheme } from './base.js'

export const alibabaTheme: Theme = createTheme({
  name: 'alibaba',
  displayName: '阿里风格',
  description: '阿里橙色系，适用于电商、零售行业',
  colors: {
    primary: 'ff6a00', secondary: '1677ff', accent: 'fa8c16', background: 'f5f5f5',
    headerBg: 'fff7e6', border: 'ffe7ba', text: '333333', textSecondary: '666666'
  },
  word: {
    fonts: { title: '阿里巴巴普惠体', heading: '阿里巴巴普惠体', body: '阿里巴巴普惠体', code: 'Consolas' },
    heading1: { font: '阿里巴巴普惠体', color: 'ff6a00', fontSize: 36, bold: true, alignment: 'left', spaceBefore: 24, spaceAfter: 18 },
    heading2: { font: '阿里巴巴普惠体', color: '1677ff', fontSize: 32, bold: true, alignment: 'left', spaceBefore: 18, spaceAfter: 12 },
    heading3: { font: '阿里巴巴普惠体', color: '333333', fontSize: 28, bold: true, alignment: 'left', spaceBefore: 12, spaceAfter: 6 },
    paragraph: {
      font: '阿里巴巴普惠体', color: '333333', fontSize: 24,
      format: { lineSpacing: 1.6, lineSpacingType: 'multiple', firstLineIndent: 0, spaceBefore: 0, spaceAfter: 8 }
    },
    table: {
      headerBg: 'fff7e6', headerColor: 'd46b08', headerFont: '阿里巴巴普惠体', headerFontSize: 22,
      headerBold: true, bodyFont: '阿里巴巴普惠体', bodyFontSize: 21, stripedBg: 'fafafa',
      borderColor: 'ffe7ba', borderWidth: 1
    },
    quote: { borderLeft: 'ff6a00', borderWidth: 4, background: 'fff7e6', font: '阿里巴巴普惠体', fontSize: 24, italic: false },
    code: { background: 'f5f5f5', font: 'Consolas', fontSize: 20, borderColor: 'e8e8e8' },
    list: { font: '阿里巴巴普惠体', fontSize: 24, indent: 480, spacing: 6 }
  },
  excel: {
    fonts: { title: '微软雅黑', heading: '微软雅黑', body: '微软雅黑', code: 'Consolas' },
    headerRow: { fill: 'fff7e6', fontColor: 'd46b08', font: '微软雅黑', fontSize: 11, bold: true, alignment: 'center' },
    dataRow: { fill: 'ffffff', font: '微软雅黑', fontSize: 10 },
    alternateRow: { fill: 'fafafa' },
    border: { color: 'ffe7ba', style: 'thin' },
    chartColors: ['ff6a00', '1677ff', '52c41a', 'faad14', '722ed1']
  }
})
