/**
 * 主题样式定义
 */

export interface ThemeColors {
  primary: string
  secondary: string
  background: string
  headerBg: string
  border: string
  text: string
  textSecondary: string
}

export interface WordStyles {
  heading1: { color: string; fontSize: number; bold: boolean }
  heading2: { color: string; fontSize: number; bold: boolean }
  heading3: { color: string; fontSize: number; bold: boolean }
  paragraph: { color: string; fontSize: number; lineSpacing: number }
  table: {
    headerBg: string
    headerColor: string
    stripedBg: string
    borderColor: string
  }
  quote: { borderLeft: string; background: string }
  code: { background: string; fontFamily: string; fontSize: number }
}

export interface ExcelStyles {
  headerRow: { fill: string; fontColor: string; bold: boolean }
  dataRow: { fill: string }
  alternateRow: { fill: string }
  border: { color: string; style: string }
  chartColors: string[]
}

export interface Theme {
  name: string
  displayName: string
  colors: ThemeColors
  word: WordStyles
  excel: ExcelStyles
}

// 阿里风格
const alibabaTheme: Theme = {
  name: 'alibaba',
  displayName: '阿里风格',
  colors: {
    primary: 'ff6a00',
    secondary: '1677ff',
    background: 'f5f5f5',
    headerBg: 'fafafa',
    border: 'e8e8e8',
    text: '333333',
    textSecondary: '666666'
  },
  word: {
    heading1: { color: 'ff6a00', fontSize: 44, bold: true },
    heading2: { color: '1677ff', fontSize: 32, bold: true },
    heading3: { color: '333333', fontSize: 28, bold: true },
    paragraph: { color: '333333', fontSize: 22, lineSpacing: 1.6 },
    table: {
      headerBg: 'fff7e6',
      headerColor: 'd46b08',
      stripedBg: 'fafafa',
      borderColor: 'ffe7ba'
    },
    quote: { borderLeft: 'ff6a00', background: 'fff7e6' },
    code: { background: 'f5f5f5', fontFamily: 'Consolas', fontSize: 20 }
  },
  excel: {
    headerRow: { fill: 'fff7e6', fontColor: 'd46b08', bold: true },
    dataRow: { fill: 'ffffff' },
    alternateRow: { fill: 'fafafa' },
    border: { color: 'ffe7ba', style: 'thin' },
    chartColors: ['ff6a00', '1677ff', '52c41a', 'faad14', '722ed1']
  }
}

// 腾讯风格
const tencentTheme: Theme = {
  name: 'tencent',
  displayName: '腾讯风格',
  colors: {
    primary: '07c160',
    secondary: '576b95',
    background: 'ededed',
    headerBg: 'f7f7f7',
    border: 'e5e5e5',
    text: '191919',
    textSecondary: '999999'
  },
  word: {
    heading1: { color: '07c160', fontSize: 44, bold: true },
    heading2: { color: '191919', fontSize: 32, bold: true },
    heading3: { color: '191919', fontSize: 28, bold: true },
    paragraph: { color: '191919', fontSize: 22, lineSpacing: 1.5 },
    table: {
      headerBg: 'f0f9eb',
      headerColor: '07c160',
      stripedBg: 'f7f7f7',
      borderColor: 'e5e5e5'
    },
    quote: { borderLeft: '07c160', background: 'f0f9eb' },
    code: { background: 'f7f7f7', fontFamily: 'Consolas', fontSize: 20 }
  },
  excel: {
    headerRow: { fill: 'f0f9eb', fontColor: '07c160', bold: true },
    dataRow: { fill: 'ffffff' },
    alternateRow: { fill: 'f7f7f7' },
    border: { color: 'e5e5e5', style: 'thin' },
    chartColors: ['07c160', '576b95', 'fa5151', 'ffc300', '1485ee']
  }
}

// 字节风格
const bytedanceTheme: Theme = {
  name: 'bytedance',
  displayName: '字节风格',
  colors: {
    primary: 'fe2c55',
    secondary: '161823',
    background: 'f8f8f8',
    headerBg: 'ffffff',
    border: 'e1e1e1',
    text: '161823',
    textSecondary: '86868b'
  },
  word: {
    heading1: { color: '161823', fontSize: 48, bold: true },
    heading2: { color: '161823', fontSize: 36, bold: true },
    heading3: { color: '161823', fontSize: 28, bold: true },
    paragraph: { color: '161823', fontSize: 22, lineSpacing: 1.8 },
    table: {
      headerBg: '161823',
      headerColor: 'ffffff',
      stripedBg: 'f8f8f8',
      borderColor: 'e1e1e1'
    },
    quote: { borderLeft: 'fe2c55', background: 'f8f8f8' },
    code: { background: 'f8f8f8', fontFamily: 'Consolas', fontSize: 20 }
  },
  excel: {
    headerRow: { fill: '161823', fontColor: 'ffffff', bold: true },
    dataRow: { fill: 'ffffff' },
    alternateRow: { fill: 'f8f8f8' },
    border: { color: 'e1e1e1', style: 'thin' },
    chartColors: ['fe2c55', '25f4ee', '161823', 'ee1d52', '69c9d0']
  }
}

// 默认专业风格
const defaultTheme: Theme = {
  name: 'default',
  displayName: '专业风格',
  colors: {
    primary: '1890ff',
    secondary: '722ed1',
    background: 'f0f2f5',
    headerBg: 'fafafa',
    border: 'd9d9d9',
    text: '262626',
    textSecondary: '8c8c8c'
  },
  word: {
    heading1: { color: '1890ff', fontSize: 44, bold: true },
    heading2: { color: '262626', fontSize: 32, bold: true },
    heading3: { color: '262626', fontSize: 28, bold: true },
    paragraph: { color: '262626', fontSize: 22, lineSpacing: 1.5 },
    table: {
      headerBg: 'e6f7ff',
      headerColor: '1890ff',
      stripedBg: 'fafafa',
      borderColor: '91d5ff'
    },
    quote: { borderLeft: '1890ff', background: 'e6f7ff' },
    code: { background: 'f5f5f5', fontFamily: 'Consolas', fontSize: 20 }
  },
  excel: {
    headerRow: { fill: 'e6f7ff', fontColor: '1890ff', bold: true },
    dataRow: { fill: 'ffffff' },
    alternateRow: { fill: 'fafafa' },
    border: { color: '91d5ff', style: 'thin' },
    chartColors: ['1890ff', '52c41a', 'faad14', 'f5222d', '722ed1']
  }
}

// 主题映射
const themes: Record<string, Theme> = {
  alibaba: alibabaTheme,
  tencent: tencentTheme,
  bytedance: bytedanceTheme,
  default: defaultTheme
}

export function getTheme(name: string): Theme {
  return themes[name] || defaultTheme
}

export { alibabaTheme, tencentTheme, bytedanceTheme, defaultTheme }
