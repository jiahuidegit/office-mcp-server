/**
 * 腾讯风格
 */

import { Theme } from './types.js'
import { createTheme } from './base.js'

export const tencentTheme: Theme = createTheme({
  name: 'tencent',
  displayName: '腾讯风格',
  description: '腾讯绿色系，适用于社交、金融行业',
  colors: {
    primary: '07c160', secondary: '576b95', accent: '10aeff', background: 'f7f7f7',
    headerBg: 'f0f9eb', border: 'e5e5e5', text: '191919', textSecondary: '999999'
  },
  word: {
    fonts: { title: '微软雅黑', heading: '微软雅黑', body: '微软雅黑', code: 'Consolas' },
    heading1: { font: '微软雅黑', color: '07c160', fontSize: 36, bold: true, alignment: 'left', spaceBefore: 24, spaceAfter: 18 },
    heading2: { font: '微软雅黑', color: '191919', fontSize: 32, bold: true, alignment: 'left', spaceBefore: 18, spaceAfter: 12 },
    heading3: { font: '微软雅黑', color: '191919', fontSize: 28, bold: true, alignment: 'left', spaceBefore: 12, spaceAfter: 6 },
    paragraph: {
      font: '微软雅黑', color: '191919', fontSize: 24,
      format: { lineSpacing: 1.5, lineSpacingType: 'multiple', firstLineIndent: 0, spaceBefore: 0, spaceAfter: 8 }
    },
    table: {
      headerBg: 'f0f9eb', headerColor: '07c160', headerFont: '微软雅黑', headerFontSize: 22,
      headerBold: true, bodyFont: '微软雅黑', bodyFontSize: 21, stripedBg: 'f7f7f7',
      borderColor: 'e5e5e5', borderWidth: 1
    },
    quote: { borderLeft: '07c160', borderWidth: 4, background: 'f0f9eb', font: '微软雅黑', fontSize: 24, italic: false },
    code: { background: 'f7f7f7', font: 'Consolas', fontSize: 20, borderColor: 'e5e5e5' },
    list: { font: '微软雅黑', fontSize: 24, indent: 480, spacing: 6 }
  },
  excel: {
    fonts: { title: '微软雅黑', heading: '微软雅黑', body: '微软雅黑', code: 'Consolas' },
    headerRow: { fill: 'f0f9eb', fontColor: '07c160', font: '微软雅黑', fontSize: 11, bold: true, alignment: 'center' },
    dataRow: { fill: 'ffffff', font: '微软雅黑', fontSize: 10 },
    alternateRow: { fill: 'f7f7f7' },
    border: { color: 'e5e5e5', style: 'thin' },
    chartColors: ['07c160', '576b95', 'fa5151', 'ffc300', '1485ee']
  }
})
