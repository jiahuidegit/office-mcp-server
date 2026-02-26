/**
 * 主题类型定义
 *
 * 字号对照表（磅值 pt，Word中使用半磅 half-point）：
 * - 初号: 42pt (84)  小初: 36pt (72)
 * - 一号: 26pt (52)  小一: 24pt (48)
 * - 二号: 22pt (44)  小二: 18pt (36)
 * - 三号: 16pt (32)  小三: 15pt (30)
 * - 四号: 14pt (28)  小四: 12pt (24)
 * - 五号: 10.5pt (21) 小五: 9pt (18)
 */

export interface ThemeColors {
  primary: string
  secondary: string
  accent: string
  background: string
  headerBg: string
  border: string
  text: string
  textSecondary: string
}

export interface FontConfig {
  title: string
  heading: string
  body: string
  code: string
}

/** 页面布局配置（单位：twips，1cm = 567 twips） */
export interface PageLayout {
  marginTop: number
  marginBottom: number
  marginLeft: number
  marginRight: number
}

export interface ParagraphFormat {
  lineSpacing: number
  lineSpacingType: 'multiple' | 'exact'
  firstLineIndent: number
  spaceBefore: number
  spaceAfter: number
}

export interface HeadingStyle {
  font: string
  color: string
  fontSize: number
  bold: boolean
  alignment: 'left' | 'center' | 'right'
  spaceBefore: number
  spaceAfter: number
}

export interface TableStyle {
  headerBg: string
  headerColor: string
  headerFont: string
  headerFontSize: number
  headerBold: boolean
  bodyFont: string
  bodyFontSize: number
  stripedBg: string
  borderColor: string
  borderWidth: number
}

export interface WordStyles {
  fonts: FontConfig
  heading1: HeadingStyle
  heading2: HeadingStyle
  heading3: HeadingStyle
  paragraph: {
    font: string
    color: string
    fontSize: number
    format: ParagraphFormat
  }
  table: TableStyle
  quote: {
    borderLeft: string
    borderWidth: number
    background: string
    font: string
    fontSize: number
    italic: boolean
  }
  code: {
    background: string
    font: string
    fontSize: number
    borderColor: string
  }
  list: {
    font: string
    fontSize: number
    indent: number
    spacing: number
  }
}

export interface ExcelStyles {
  fonts: FontConfig
  headerRow: {
    fill: string
    fontColor: string
    font: string
    fontSize: number
    bold: boolean
    alignment: 'left' | 'center' | 'right'
  }
  dataRow: {
    fill: string
    font: string
    fontSize: number
  }
  alternateRow: { fill: string }
  border: { color: string; style: string }
  chartColors: string[]
}

export interface Theme {
  name: string
  displayName: string
  description: string
  colors: ThemeColors
  pageLayout: PageLayout
  word: WordStyles
  excel: ExcelStyles
}
