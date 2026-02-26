/**
 * 党政公文风格 (GB/T 9704-2012)
 */

import { Theme } from './types.js'
import { cmToTwips } from './base.js'

export const governmentTheme: Theme = {
  name: 'government',
  displayName: '党政公文',
  description: '符合GB/T 9704-2012标准的党政机关公文格式',
  colors: {
    primary: 'cc0000',
    secondary: '000000',
    accent: 'cc0000',
    background: 'ffffff',
    headerBg: 'f0f0f0',
    border: '000000',
    text: '000000',
    textSecondary: '333333'
  },
  pageLayout: {
    marginTop: cmToTwips(3.7),
    marginBottom: cmToTwips(3.5),
    marginLeft: cmToTwips(2.8),
    marginRight: cmToTwips(2.6)
  },
  word: {
    fonts: { title: '方正小标宋简体', heading: '黑体', body: '仿宋', code: 'Consolas' },
    heading1: {
      font: '方正小标宋简体', color: '000000', fontSize: 44,
      bold: false, alignment: 'center', spaceBefore: 0, spaceAfter: 28
    },
    heading2: {
      font: '黑体', color: '000000', fontSize: 32,
      bold: false, alignment: 'left', spaceBefore: 0, spaceAfter: 0
    },
    heading3: {
      font: '楷体', color: '000000', fontSize: 32,
      bold: true, alignment: 'left', spaceBefore: 0, spaceAfter: 0
    },
    paragraph: {
      font: '仿宋', color: '000000', fontSize: 32,
      format: { lineSpacing: 28, lineSpacingType: 'exact', firstLineIndent: 2, spaceBefore: 0, spaceAfter: 0 }
    },
    table: {
      headerBg: 'f0f0f0', headerColor: '000000', headerFont: '黑体', headerFontSize: 24,
      headerBold: true, bodyFont: '仿宋', bodyFontSize: 24, stripedBg: 'fafafa',
      borderColor: '000000', borderWidth: 1
    },
    quote: { borderLeft: 'cc0000', borderWidth: 3, background: 'fff5f5', font: '楷体', fontSize: 32, italic: false },
    code: { background: 'f5f5f5', font: 'Consolas', fontSize: 24, borderColor: 'cccccc' },
    list: { font: '仿宋', fontSize: 32, indent: 640, spacing: 0 }
  },
  excel: {
    fonts: { title: '黑体', heading: '黑体', body: '宋体', code: 'Consolas' },
    headerRow: { fill: 'f0f0f0', fontColor: '000000', font: '黑体', fontSize: 11, bold: true, alignment: 'center' },
    dataRow: { fill: 'ffffff', font: '宋体', fontSize: 10 },
    alternateRow: { fill: 'fafafa' },
    border: { color: '000000', style: 'thin' },
    chartColors: ['cc0000', '1a1a1a', '666666', '999999', 'cccccc']
  }
}
