/**
 * 主题样式定义
 * 基于中国文档格式规范设计
 *
 * 参考标准：
 * - GB/T 9704-2012 党政机关公文格式
 * - GB/T 8567-2006 计算机软件文档编制规范
 * - 学术论文格式规范
 * - 企业商务文档规范
 *
 * 字号对照表（磅值 pt，Word中使用半磅 half-point）：
 * - 初号: 42pt (84 half-points)
 * - 小初: 36pt (72 half-points)
 * - 一号: 26pt (52 half-points)
 * - 小一: 24pt (48 half-points)
 * - 二号: 22pt (44 half-points)
 * - 小二: 18pt (36 half-points)
 * - 三号: 16pt (32 half-points)
 * - 小三: 15pt (30 half-points)
 * - 四号: 14pt (28 half-points)
 * - 小四: 12pt (24 half-points)
 * - 五号: 10.5pt (21 half-points)
 * - 小五: 9pt (18 half-points)
 */

// ============ 类型定义 ============

export interface ThemeColors {
  primary: string      // 主色
  secondary: string    // 次色
  accent: string       // 强调色
  background: string   // 背景色
  headerBg: string     // 表头背景
  border: string       // 边框色
  text: string         // 正文文字
  textSecondary: string // 次要文字
}

// 字体配置
export interface FontConfig {
  title: string        // 大标题字体
  heading: string      // 章节标题字体
  body: string         // 正文字体
  code: string         // 代码字体
}

// 页面布局配置（单位：twips，1cm = 567 twips）
export interface PageLayout {
  marginTop: number    // 上边距
  marginBottom: number // 下边距
  marginLeft: number   // 左边距
  marginRight: number  // 右边距
}

// 段落格式配置
export interface ParagraphFormat {
  lineSpacing: number      // 行距（倍数，如1.5）或固定值
  lineSpacingType: 'multiple' | 'exact'  // multiple=倍数, exact=固定值(磅)
  firstLineIndent: number  // 首行缩进（字符数，如2）
  spaceBefore: number      // 段前间距（磅）
  spaceAfter: number       // 段后间距（磅）
}

// 标题样式配置
export interface HeadingStyle {
  font: string
  color: string
  fontSize: number     // half-points
  bold: boolean
  alignment: 'left' | 'center' | 'right'
  spaceBefore: number  // 磅
  spaceAfter: number   // 磅
}

// 表格样式配置
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

// Word 样式配置
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
    indent: number      // 缩进（twips）
    spacing: number     // 项目间距（磅）
  }
}

// Excel 样式配置
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

// 完整主题定义
export interface Theme {
  name: string
  displayName: string
  description: string
  colors: ThemeColors
  pageLayout: PageLayout
  word: WordStyles
  excel: ExcelStyles
}

// ============ 工具函数 ============

// 厘米转 twips (1cm = 567 twips)
const cmToTwips = (cm: number) => Math.round(cm * 567)

// ============ 主题定义 ============

/**
 * 党政公文风格 (GB/T 9704-2012)
 * 严格遵循国家标准
 */
const governmentTheme: Theme = {
  name: 'government',
  displayName: '党政公文',
  description: '符合GB/T 9704-2012标准的党政机关公文格式',
  colors: {
    primary: 'cc0000',    // 公文红
    secondary: '000000',
    accent: 'cc0000',
    background: 'ffffff',
    headerBg: 'f0f0f0',
    border: '000000',
    text: '000000',
    textSecondary: '333333'
  },
  // 页边距：上3.7cm 下3.5cm 左2.8cm 右2.6cm
  pageLayout: {
    marginTop: cmToTwips(3.7),
    marginBottom: cmToTwips(3.5),
    marginLeft: cmToTwips(2.8),
    marginRight: cmToTwips(2.6)
  },
  word: {
    fonts: {
      title: '方正小标宋简体',  // 标题用小标宋
      heading: '黑体',          // 一级标题用黑体
      body: '仿宋',             // 正文用仿宋
      code: 'Consolas'
    },
    // 文件标题：二号小标宋，居中
    heading1: {
      font: '方正小标宋简体',
      color: '000000',
      fontSize: 44,  // 二号 22pt
      bold: false,   // 小标宋本身就是粗体效果
      alignment: 'center',
      spaceBefore: 0,
      spaceAfter: 28  // 段后一行
    },
    // 一级标题：三号黑体
    heading2: {
      font: '黑体',
      color: '000000',
      fontSize: 32,  // 三号 16pt
      bold: false,
      alignment: 'left',
      spaceBefore: 0,
      spaceAfter: 0
    },
    // 二级标题：三号楷体加粗
    heading3: {
      font: '楷体',
      color: '000000',
      fontSize: 32,  // 三号 16pt
      bold: true,
      alignment: 'left',
      spaceBefore: 0,
      spaceAfter: 0
    },
    // 正文：三号仿宋，固定行距28磅，首行缩进2字符
    paragraph: {
      font: '仿宋',
      color: '000000',
      fontSize: 32,  // 三号 16pt
      format: {
        lineSpacing: 28,
        lineSpacingType: 'exact',
        firstLineIndent: 2,
        spaceBefore: 0,
        spaceAfter: 0
      }
    },
    table: {
      headerBg: 'f0f0f0',
      headerColor: '000000',
      headerFont: '黑体',
      headerFontSize: 24,  // 小四
      headerBold: true,
      bodyFont: '仿宋',
      bodyFontSize: 24,
      stripedBg: 'fafafa',
      borderColor: '000000',
      borderWidth: 1
    },
    quote: {
      borderLeft: 'cc0000',
      borderWidth: 3,
      background: 'fff5f5',
      font: '楷体',
      fontSize: 32,
      italic: false
    },
    code: {
      background: 'f5f5f5',
      font: 'Consolas',
      fontSize: 24,
      borderColor: 'cccccc'
    },
    list: {
      font: '仿宋',
      fontSize: 32,
      indent: 640,    // 约2字符
      spacing: 0
    }
  },
  excel: {
    fonts: {
      title: '黑体',
      heading: '黑体',
      body: '宋体',
      code: 'Consolas'
    },
    headerRow: {
      fill: 'f0f0f0',
      fontColor: '000000',
      font: '黑体',
      fontSize: 11,
      bold: true,
      alignment: 'center'
    },
    dataRow: {
      fill: 'ffffff',
      font: '宋体',
      fontSize: 10
    },
    alternateRow: { fill: 'fafafa' },
    border: { color: '000000', style: 'thin' },
    chartColors: ['cc0000', '1a1a1a', '666666', '999999', 'cccccc']
  }
}

/**
 * 学术论文风格
 * 适用于学术论文、研究报告、毕业论文
 */
const academicTheme: Theme = {
  name: 'academic',
  displayName: '学术论文',
  description: '符合学术论文规范，适用于研究报告、毕业论文',
  colors: {
    primary: '000000',
    secondary: '333333',
    accent: '1a73e8',
    background: 'ffffff',
    headerBg: 'f8f9fa',
    border: 'dadce0',
    text: '000000',
    textSecondary: '5f6368'
  },
  // 页边距：上下2.5cm 左3cm 右2.5cm
  pageLayout: {
    marginTop: cmToTwips(2.5),
    marginBottom: cmToTwips(2.5),
    marginLeft: cmToTwips(3.0),
    marginRight: cmToTwips(2.5)
  },
  word: {
    fonts: {
      title: '黑体',
      heading: '黑体',
      body: '宋体',
      code: 'Consolas'
    },
    // 论文标题：小三黑体，居中
    heading1: {
      font: '黑体',
      color: '000000',
      fontSize: 30,  // 小三 15pt
      bold: true,
      alignment: 'center',
      spaceBefore: 24,
      spaceAfter: 18
    },
    // 一级标题：四号黑体
    heading2: {
      font: '黑体',
      color: '000000',
      fontSize: 28,  // 四号 14pt
      bold: false,
      alignment: 'left',
      spaceBefore: 12,
      spaceAfter: 6
    },
    // 二级标题：小四黑体
    heading3: {
      font: '黑体',
      color: '000000',
      fontSize: 24,  // 小四 12pt
      bold: false,
      alignment: 'left',
      spaceBefore: 6,
      spaceAfter: 6
    },
    // 正文：小四宋体，1.5倍行距，首行缩进2字符
    paragraph: {
      font: '宋体',
      color: '000000',
      fontSize: 24,  // 小四 12pt
      format: {
        lineSpacing: 1.5,
        lineSpacingType: 'multiple',
        firstLineIndent: 2,
        spaceBefore: 0,
        spaceAfter: 0
      }
    },
    table: {
      headerBg: 'f8f9fa',
      headerColor: '000000',
      headerFont: '黑体',
      headerFontSize: 21,  // 五号
      headerBold: true,
      bodyFont: '宋体',
      bodyFontSize: 21,
      stripedBg: 'f8f9fa',
      borderColor: '000000',
      borderWidth: 1
    },
    quote: {
      borderLeft: '1a73e8',
      borderWidth: 3,
      background: 'e8f0fe',
      font: '楷体',
      fontSize: 24,
      italic: true
    },
    code: {
      background: 'f8f9fa',
      font: 'Consolas',
      fontSize: 20,
      borderColor: 'dadce0'
    },
    list: {
      font: '宋体',
      fontSize: 24,
      indent: 480,
      spacing: 6
    }
  },
  excel: {
    fonts: {
      title: '黑体',
      heading: '黑体',
      body: '宋体',
      code: 'Consolas'
    },
    headerRow: {
      fill: 'f8f9fa',
      fontColor: '000000',
      font: '黑体',
      fontSize: 10,
      bold: true,
      alignment: 'center'
    },
    dataRow: {
      fill: 'ffffff',
      font: '宋体',
      fontSize: 10
    },
    alternateRow: { fill: 'f8f9fa' },
    border: { color: '000000', style: 'thin' },
    chartColors: ['1a73e8', '34a853', 'fbbc04', 'ea4335', '673ab7']
  }
}

/**
 * 软件文档风格 (参考 GB/T 8567-2006)
 * 适用于需求规格说明书、设计文档、技术方案
 *
 * 格式规范：
 * - 纸张：A4 (210mm × 297mm)
 * - 页边距：上下2.54cm 左右3.17cm
 * - 标题：黑体，文档标题二号居中，章节标题三号/四号左对齐
 * - 正文：小四宋体，1.5倍行距，首行缩进2字符
 * - 编号格式：1、1.1、1.1.1
 */
const softwareDocTheme: Theme = {
  name: 'software',
  displayName: '软件文档',
  description: '符合GB/T 8567-2006规范，适用于需求规格说明书、设计文档、技术方案',
  colors: {
    primary: '000000',     // 正式文档使用黑色
    secondary: '333333',
    accent: '1a73e8',      // 蓝色用于链接/强调
    background: 'ffffff',
    headerBg: 'f5f5f5',
    border: '000000',
    text: '000000',
    textSecondary: '666666'
  },
  // 页边距：上下2.54cm 左右3.17cm (国标要求)
  pageLayout: {
    marginTop: cmToTwips(2.54),
    marginBottom: cmToTwips(2.54),
    marginLeft: cmToTwips(3.17),
    marginRight: cmToTwips(3.17)
  },
  word: {
    fonts: {
      title: '黑体',      // 标题用黑体 (国标)
      heading: '黑体',    // 章节标题用黑体
      body: '宋体',       // 正文用宋体
      code: 'Consolas'
    },
    // 文档标题：二号黑体，居中
    heading1: {
      font: '黑体',
      color: '000000',
      fontSize: 44,  // 二号 22pt
      bold: true,
      alignment: 'center',
      spaceBefore: 24,
      spaceAfter: 24
    },
    // 一级标题：三号黑体
    heading2: {
      font: '黑体',
      color: '000000',
      fontSize: 32,  // 三号 16pt
      bold: false,
      alignment: 'left',
      spaceBefore: 12,
      spaceAfter: 6
    },
    // 二级标题：四号黑体
    heading3: {
      font: '黑体',
      color: '000000',
      fontSize: 28,  // 四号 14pt
      bold: false,
      alignment: 'left',
      spaceBefore: 6,
      spaceAfter: 6
    },
    // 正文：小四宋体，1.5倍行距，首行缩进2字符
    paragraph: {
      font: '宋体',
      color: '000000',
      fontSize: 24,  // 小四 12pt
      format: {
        lineSpacing: 1.5,
        lineSpacingType: 'multiple',
        firstLineIndent: 2,
        spaceBefore: 0,
        spaceAfter: 0
      }
    },
    table: {
      headerBg: 'f5f5f5',
      headerColor: '000000',
      headerFont: '黑体',
      headerFontSize: 21,  // 五号
      headerBold: true,
      bodyFont: '宋体',
      bodyFontSize: 21,
      stripedBg: 'fafafa',
      borderColor: '000000',
      borderWidth: 1
    },
    quote: {
      borderLeft: '1a73e8',
      borderWidth: 3,
      background: 'f5f5f5',
      font: '楷体',
      fontSize: 24,
      italic: false
    },
    code: {
      background: 'f5f5f5',
      font: 'Consolas',
      fontSize: 20,
      borderColor: 'cccccc'
    },
    list: {
      font: '宋体',
      fontSize: 24,
      indent: 480,
      spacing: 0
    }
  },
  excel: {
    fonts: {
      title: '黑体',
      heading: '黑体',
      body: '宋体',
      code: 'Consolas'
    },
    headerRow: {
      fill: 'f5f5f5',
      fontColor: '000000',
      font: '黑体',
      fontSize: 10,
      bold: true,
      alignment: 'center'
    },
    dataRow: {
      fill: 'ffffff',
      font: '宋体',
      fontSize: 10
    },
    alternateRow: { fill: 'fafafa' },
    border: { color: '000000', style: 'thin' },
    chartColors: ['1a73e8', '34a853', 'fbbc04', 'ea4335', '673ab7']
  }
}

/**
 * 商务报告风格
 * 适用于市场调研报告、商业计划书、工作汇报
 */
const businessTheme: Theme = {
  name: 'business',
  displayName: '商务报告',
  description: '适用于市场调研报告、商业计划书、工作汇报',
  colors: {
    primary: '1890ff',
    secondary: '13c2c2',
    accent: '722ed1',
    background: 'f5f5f5',
    headerBg: 'e6f7ff',
    border: '91d5ff',
    text: '262626',
    textSecondary: '8c8c8c'
  },
  pageLayout: {
    marginTop: cmToTwips(2.54),
    marginBottom: cmToTwips(2.54),
    marginLeft: cmToTwips(3.17),
    marginRight: cmToTwips(3.17)
  },
  word: {
    fonts: {
      title: '微软雅黑',
      heading: '微软雅黑',
      body: '微软雅黑',
      code: 'Consolas'
    },
    heading1: {
      font: '微软雅黑',
      color: '1890ff',
      fontSize: 36,  // 小二
      bold: true,
      alignment: 'center',
      spaceBefore: 24,
      spaceAfter: 24
    },
    heading2: {
      font: '微软雅黑',
      color: '262626',
      fontSize: 32,  // 三号
      bold: true,
      alignment: 'left',
      spaceBefore: 18,
      spaceAfter: 12
    },
    heading3: {
      font: '微软雅黑',
      color: '262626',
      fontSize: 28,  // 四号
      bold: true,
      alignment: 'left',
      spaceBefore: 12,
      spaceAfter: 6
    },
    paragraph: {
      font: '微软雅黑',
      color: '262626',
      fontSize: 24,
      format: {
        lineSpacing: 1.5,
        lineSpacingType: 'multiple',
        firstLineIndent: 0,  // 商务文档通常不缩进
        spaceBefore: 0,
        spaceAfter: 8
      }
    },
    table: {
      headerBg: 'e6f7ff',
      headerColor: '1890ff',
      headerFont: '微软雅黑',
      headerFontSize: 22,
      headerBold: true,
      bodyFont: '微软雅黑',
      bodyFontSize: 21,
      stripedBg: 'fafafa',
      borderColor: '91d5ff',
      borderWidth: 1
    },
    quote: {
      borderLeft: '1890ff',
      borderWidth: 4,
      background: 'e6f7ff',
      font: '微软雅黑',
      fontSize: 24,
      italic: false
    },
    code: {
      background: 'f5f5f5',
      font: 'Consolas',
      fontSize: 20,
      borderColor: 'd9d9d9'
    },
    list: {
      font: '微软雅黑',
      fontSize: 24,
      indent: 480,
      spacing: 6
    }
  },
  excel: {
    fonts: {
      title: '微软雅黑',
      heading: '微软雅黑',
      body: '微软雅黑',
      code: 'Consolas'
    },
    headerRow: {
      fill: 'e6f7ff',
      fontColor: '1890ff',
      font: '微软雅黑',
      fontSize: 11,
      bold: true,
      alignment: 'center'
    },
    dataRow: {
      fill: 'ffffff',
      font: '微软雅黑',
      fontSize: 10
    },
    alternateRow: { fill: 'fafafa' },
    border: { color: '91d5ff', style: 'thin' },
    chartColors: ['1890ff', '52c41a', 'faad14', 'f5222d', '722ed1']
  }
}

/**
 * 阿里风格
 */
const alibabaTheme: Theme = {
  name: 'alibaba',
  displayName: '阿里风格',
  description: '阿里橙色系，适用于电商、零售行业',
  colors: {
    primary: 'ff6a00',
    secondary: '1677ff',
    accent: 'fa8c16',
    background: 'f5f5f5',
    headerBg: 'fff7e6',
    border: 'ffe7ba',
    text: '333333',
    textSecondary: '666666'
  },
  pageLayout: {
    marginTop: cmToTwips(2.54),
    marginBottom: cmToTwips(2.54),
    marginLeft: cmToTwips(3.17),
    marginRight: cmToTwips(3.17)
  },
  word: {
    fonts: {
      title: '阿里巴巴普惠体',
      heading: '阿里巴巴普惠体',
      body: '阿里巴巴普惠体',
      code: 'Consolas'
    },
    heading1: {
      font: '阿里巴巴普惠体',
      color: 'ff6a00',
      fontSize: 36,
      bold: true,
      alignment: 'left',
      spaceBefore: 24,
      spaceAfter: 18
    },
    heading2: {
      font: '阿里巴巴普惠体',
      color: '1677ff',
      fontSize: 32,
      bold: true,
      alignment: 'left',
      spaceBefore: 18,
      spaceAfter: 12
    },
    heading3: {
      font: '阿里巴巴普惠体',
      color: '333333',
      fontSize: 28,
      bold: true,
      alignment: 'left',
      spaceBefore: 12,
      spaceAfter: 6
    },
    paragraph: {
      font: '阿里巴巴普惠体',
      color: '333333',
      fontSize: 24,
      format: {
        lineSpacing: 1.6,
        lineSpacingType: 'multiple',
        firstLineIndent: 0,
        spaceBefore: 0,
        spaceAfter: 8
      }
    },
    table: {
      headerBg: 'fff7e6',
      headerColor: 'd46b08',
      headerFont: '阿里巴巴普惠体',
      headerFontSize: 22,
      headerBold: true,
      bodyFont: '阿里巴巴普惠体',
      bodyFontSize: 21,
      stripedBg: 'fafafa',
      borderColor: 'ffe7ba',
      borderWidth: 1
    },
    quote: {
      borderLeft: 'ff6a00',
      borderWidth: 4,
      background: 'fff7e6',
      font: '阿里巴巴普惠体',
      fontSize: 24,
      italic: false
    },
    code: {
      background: 'f5f5f5',
      font: 'Consolas',
      fontSize: 20,
      borderColor: 'e8e8e8'
    },
    list: {
      font: '阿里巴巴普惠体',
      fontSize: 24,
      indent: 480,
      spacing: 6
    }
  },
  excel: {
    fonts: {
      title: '微软雅黑',
      heading: '微软雅黑',
      body: '微软雅黑',
      code: 'Consolas'
    },
    headerRow: {
      fill: 'fff7e6',
      fontColor: 'd46b08',
      font: '微软雅黑',
      fontSize: 11,
      bold: true,
      alignment: 'center'
    },
    dataRow: {
      fill: 'ffffff',
      font: '微软雅黑',
      fontSize: 10
    },
    alternateRow: { fill: 'fafafa' },
    border: { color: 'ffe7ba', style: 'thin' },
    chartColors: ['ff6a00', '1677ff', '52c41a', 'faad14', '722ed1']
  }
}

/**
 * 腾讯风格
 */
const tencentTheme: Theme = {
  name: 'tencent',
  displayName: '腾讯风格',
  description: '腾讯绿色系，适用于社交、金融行业',
  colors: {
    primary: '07c160',
    secondary: '576b95',
    accent: '10aeff',
    background: 'f7f7f7',
    headerBg: 'f0f9eb',
    border: 'e5e5e5',
    text: '191919',
    textSecondary: '999999'
  },
  pageLayout: {
    marginTop: cmToTwips(2.54),
    marginBottom: cmToTwips(2.54),
    marginLeft: cmToTwips(3.17),
    marginRight: cmToTwips(3.17)
  },
  word: {
    fonts: {
      title: '微软雅黑',
      heading: '微软雅黑',
      body: '微软雅黑',
      code: 'Consolas'
    },
    heading1: {
      font: '微软雅黑',
      color: '07c160',
      fontSize: 36,
      bold: true,
      alignment: 'left',
      spaceBefore: 24,
      spaceAfter: 18
    },
    heading2: {
      font: '微软雅黑',
      color: '191919',
      fontSize: 32,
      bold: true,
      alignment: 'left',
      spaceBefore: 18,
      spaceAfter: 12
    },
    heading3: {
      font: '微软雅黑',
      color: '191919',
      fontSize: 28,
      bold: true,
      alignment: 'left',
      spaceBefore: 12,
      spaceAfter: 6
    },
    paragraph: {
      font: '微软雅黑',
      color: '191919',
      fontSize: 24,
      format: {
        lineSpacing: 1.5,
        lineSpacingType: 'multiple',
        firstLineIndent: 0,
        spaceBefore: 0,
        spaceAfter: 8
      }
    },
    table: {
      headerBg: 'f0f9eb',
      headerColor: '07c160',
      headerFont: '微软雅黑',
      headerFontSize: 22,
      headerBold: true,
      bodyFont: '微软雅黑',
      bodyFontSize: 21,
      stripedBg: 'f7f7f7',
      borderColor: 'e5e5e5',
      borderWidth: 1
    },
    quote: {
      borderLeft: '07c160',
      borderWidth: 4,
      background: 'f0f9eb',
      font: '微软雅黑',
      fontSize: 24,
      italic: false
    },
    code: {
      background: 'f7f7f7',
      font: 'Consolas',
      fontSize: 20,
      borderColor: 'e5e5e5'
    },
    list: {
      font: '微软雅黑',
      fontSize: 24,
      indent: 480,
      spacing: 6
    }
  },
  excel: {
    fonts: {
      title: '微软雅黑',
      heading: '微软雅黑',
      body: '微软雅黑',
      code: 'Consolas'
    },
    headerRow: {
      fill: 'f0f9eb',
      fontColor: '07c160',
      font: '微软雅黑',
      fontSize: 11,
      bold: true,
      alignment: 'center'
    },
    dataRow: {
      fill: 'ffffff',
      font: '微软雅黑',
      fontSize: 10
    },
    alternateRow: { fill: 'f7f7f7' },
    border: { color: 'e5e5e5', style: 'thin' },
    chartColors: ['07c160', '576b95', 'fa5151', 'ffc300', '1485ee']
  }
}

/**
 * 字节风格
 */
const bytedanceTheme: Theme = {
  name: 'bytedance',
  displayName: '字节风格',
  description: '字节黑色系，适用于科技、媒体行业',
  colors: {
    primary: 'fe2c55',
    secondary: '161823',
    accent: '25f4ee',
    background: 'f8f8f8',
    headerBg: '161823',
    border: 'e1e1e1',
    text: '161823',
    textSecondary: '86868b'
  },
  pageLayout: {
    marginTop: cmToTwips(2.54),
    marginBottom: cmToTwips(2.54),
    marginLeft: cmToTwips(3.17),
    marginRight: cmToTwips(3.17)
  },
  word: {
    fonts: {
      title: '抖音美好体',
      heading: '抖音美好体',
      body: '微软雅黑',
      code: 'Consolas'
    },
    heading1: {
      font: '抖音美好体',
      color: '161823',
      fontSize: 40,
      bold: true,
      alignment: 'left',
      spaceBefore: 24,
      spaceAfter: 18
    },
    heading2: {
      font: '抖音美好体',
      color: '161823',
      fontSize: 32,
      bold: true,
      alignment: 'left',
      spaceBefore: 18,
      spaceAfter: 12
    },
    heading3: {
      font: '微软雅黑',
      color: '161823',
      fontSize: 28,
      bold: true,
      alignment: 'left',
      spaceBefore: 12,
      spaceAfter: 6
    },
    paragraph: {
      font: '微软雅黑',
      color: '161823',
      fontSize: 24,
      format: {
        lineSpacing: 1.8,
        lineSpacingType: 'multiple',
        firstLineIndent: 0,
        spaceBefore: 0,
        spaceAfter: 10
      }
    },
    table: {
      headerBg: '161823',
      headerColor: 'ffffff',
      headerFont: '微软雅黑',
      headerFontSize: 22,
      headerBold: true,
      bodyFont: '微软雅黑',
      bodyFontSize: 21,
      stripedBg: 'f8f8f8',
      borderColor: 'e1e1e1',
      borderWidth: 1
    },
    quote: {
      borderLeft: 'fe2c55',
      borderWidth: 4,
      background: 'fff0f3',
      font: '微软雅黑',
      fontSize: 24,
      italic: false
    },
    code: {
      background: 'f8f8f8',
      font: 'Consolas',
      fontSize: 20,
      borderColor: 'e1e1e1'
    },
    list: {
      font: '微软雅黑',
      fontSize: 24,
      indent: 480,
      spacing: 8
    }
  },
  excel: {
    fonts: {
      title: '微软雅黑',
      heading: '微软雅黑',
      body: '微软雅黑',
      code: 'Consolas'
    },
    headerRow: {
      fill: '161823',
      fontColor: 'ffffff',
      font: '微软雅黑',
      fontSize: 11,
      bold: true,
      alignment: 'center'
    },
    dataRow: {
      fill: 'ffffff',
      font: '微软雅黑',
      fontSize: 10
    },
    alternateRow: { fill: 'f8f8f8' },
    border: { color: 'e1e1e1', style: 'thin' },
    chartColors: ['fe2c55', '25f4ee', '161823', 'ee1d52', '69c9d0']
  }
}

/**
 * 简约黑白风格
 */
const minimalTheme: Theme = {
  name: 'minimal',
  displayName: '简约黑白',
  description: '简洁的黑白配色，通用于各类正式文档',
  colors: {
    primary: '000000',
    secondary: '666666',
    accent: '000000',
    background: 'ffffff',
    headerBg: 'f5f5f5',
    border: 'cccccc',
    text: '000000',
    textSecondary: '666666'
  },
  pageLayout: {
    marginTop: cmToTwips(2.54),
    marginBottom: cmToTwips(2.54),
    marginLeft: cmToTwips(3.17),
    marginRight: cmToTwips(3.17)
  },
  word: {
    fonts: {
      title: '黑体',
      heading: '黑体',
      body: '宋体',
      code: 'Consolas'
    },
    heading1: {
      font: '黑体',
      color: '000000',
      fontSize: 36,
      bold: true,
      alignment: 'center',
      spaceBefore: 24,
      spaceAfter: 18
    },
    heading2: {
      font: '黑体',
      color: '000000',
      fontSize: 30,
      bold: false,
      alignment: 'left',
      spaceBefore: 12,
      spaceAfter: 6
    },
    heading3: {
      font: '黑体',
      color: '000000',
      fontSize: 26,
      bold: false,
      alignment: 'left',
      spaceBefore: 6,
      spaceAfter: 6
    },
    paragraph: {
      font: '宋体',
      color: '000000',
      fontSize: 24,
      format: {
        lineSpacing: 1.5,
        lineSpacingType: 'multiple',
        firstLineIndent: 2,
        spaceBefore: 0,
        spaceAfter: 0
      }
    },
    table: {
      headerBg: 'f5f5f5',
      headerColor: '000000',
      headerFont: '黑体',
      headerFontSize: 22,
      headerBold: true,
      bodyFont: '宋体',
      bodyFontSize: 21,
      stripedBg: 'fafafa',
      borderColor: '000000',
      borderWidth: 1
    },
    quote: {
      borderLeft: '000000',
      borderWidth: 3,
      background: 'f5f5f5',
      font: '楷体',
      fontSize: 24,
      italic: true
    },
    code: {
      background: 'f5f5f5',
      font: 'Consolas',
      fontSize: 20,
      borderColor: 'cccccc'
    },
    list: {
      font: '宋体',
      fontSize: 24,
      indent: 480,
      spacing: 6
    }
  },
  excel: {
    fonts: {
      title: '黑体',
      heading: '黑体',
      body: '宋体',
      code: 'Consolas'
    },
    headerRow: {
      fill: 'f5f5f5',
      fontColor: '000000',
      font: '黑体',
      fontSize: 11,
      bold: true,
      alignment: 'center'
    },
    dataRow: {
      fill: 'ffffff',
      font: '宋体',
      fontSize: 10
    },
    alternateRow: { fill: 'fafafa' },
    border: { color: '000000', style: 'thin' },
    chartColors: ['000000', '333333', '666666', '999999', 'cccccc']
  }
}

// 默认使用商务风格
const defaultTheme = businessTheme

// 主题映射
const themes: Record<string, Theme> = {
  government: governmentTheme,
  academic: academicTheme,
  software: softwareDocTheme,
  business: businessTheme,
  alibaba: alibabaTheme,
  tencent: tencentTheme,
  bytedance: bytedanceTheme,
  minimal: minimalTheme,
  default: defaultTheme
}

/**
 * 获取主题
 */
export function getTheme(name: string): Theme {
  return themes[name] || defaultTheme
}

/**
 * 获取所有可用主题列表
 */
export function listThemes(): Array<{ name: string; displayName: string; description: string }> {
  return Object.values(themes)
    .filter(t => t.name !== 'default')
    .map(t => ({
      name: t.name,
      displayName: t.displayName,
      description: t.description
    }))
}

export {
  governmentTheme,
  academicTheme,
  softwareDocTheme,
  businessTheme,
  alibabaTheme,
  tencentTheme,
  bytedanceTheme,
  minimalTheme,
  defaultTheme
}
