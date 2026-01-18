/**
 * 测试脚本 - 生成调研报告
 * 使用学术论文主题，符合中国文档规范
 */

import { handleWordTool } from './tools/word/index.js'

async function generateReport() {
  console.log('🚀 开始生成调研报告...\n')

  // 1. 创建文档（使用学术论文主题）
  const createResult = await handleWordTool('word_create', {
    title: 'AI 办公文档生成技术调研报告',
    theme: 'government', // 党政公文风格：小标宋标题 + 仿宋正文
    author: '技术团队',
    outputPath: '/Users/yangjiahui/Desktop/调研报告.docx'
  })
  const docId = JSON.parse(createResult.content[0].text).docId
  console.log('✅ 文档创建成功:', docId)

  // 2. 添加标题
  await handleWordTool('word_add_heading', {
    docId,
    text: 'AI 办公文档生成技术调研报告',
    level: 1
  })

  // 3. 添加背景段落
  await handleWordTool('word_add_heading', { docId, text: '一、调研背景', level: 2 })
  await handleWordTool('word_add_paragraph', {
    docId,
    text: '随着 AI 技术的快速发展，传统的办公文档生成方式面临效率低、格式不统一等问题。本次调研旨在探索如何利用 MCP（Model Context Protocol）技术，让 AI 能够生成专业、美观的 Word 和 Excel 文档。'
  })

  // 4. 添加调研目标
  await handleWordTool('word_add_heading', { docId, text: '二、调研目标', level: 2 })
  await handleWordTool('word_add_list', {
    docId,
    items: [
      '调研主流的 Node.js 文档生成库',
      '设计支持多种企业风格的主题系统',
      '实现 MCP 工具接口，让 AI 可以调用',
      '验证生成文档的质量和可用性'
    ],
    ordered: true
  })

  // 5. 技术方案
  await handleWordTool('word_add_heading', { docId, text: '三、技术方案', level: 2 })
  await handleWordTool('word_add_paragraph', {
    docId,
    text: '经过对比分析，我们选择了以下技术栈：',
    bold: true
  })

  await handleWordTool('word_add_table', {
    docId,
    headers: ['技术', '用途', '优势'],
    rows: [
      ['docx', 'Word 文档生成', 'API 优雅，样式控制细腻'],
      ['exceljs', 'Excel 表格生成', '图表支持好，格式丰富'],
      ['MCP SDK', '协议实现', '官方支持，生态完善'],
      ['TypeScript', '开发语言', '类型安全，维护性强']
    ],
    style: 'striped'
  })

  // 6. 主题风格
  await handleWordTool('word_add_heading', { docId, text: '四、主题风格设计', level: 2 })
  await handleWordTool('word_add_paragraph', {
    docId,
    text: '为了满足不同企业的视觉需求，我们设计了四套主题风格：'
  })

  await handleWordTool('word_add_table', {
    docId,
    headers: ['主题', '主色调', '适用场景'],
    rows: [
      ['alibaba', '阿里橙 #FF6A00', '电商、零售行业'],
      ['tencent', '腾讯绿 #07C160', '社交、金融行业'],
      ['bytedance', '字节黑 #161823', '科技、媒体行业'],
      ['default', '专业蓝 #1890FF', '通用企业文档']
    ],
    style: 'professional'
  })

  // 7. 代码示例
  await handleWordTool('word_add_heading', { docId, text: '五、使用示例', level: 2 })
  await handleWordTool('word_add_paragraph', {
    docId,
    text: '以下是生成周报的代码示例：'
  })

  await handleWordTool('word_add_code', {
    docId,
    language: 'typescript',
    code: `// 创建周报文档
const result = await word_create({
  title: "2024年第15周工作周报",
  theme: "alibaba",
  outputPath: "./周报.docx"
});

// 添加内容
await word_add_heading({ docId, text: "本周完成", level: 2 });
await word_add_list({ docId, items: ["功能A", "功能B"], ordered: false });

// 保存
await word_save({ docId, addPageNumbers: true });`
  })

  // 8. 调研结论
  await handleWordTool('word_add_heading', { docId, text: '六、调研结论', level: 2 })
  await handleWordTool('word_add_paragraph', {
    docId,
    text: '本次调研验证了通过 MCP 服务让 AI 生成专业文档的可行性。主要结论如下：',
    style: 'note'
  })
  await handleWordTool('word_add_list', {
    docId,
    items: [
      'Node.js 生态有成熟的文档生成库，可以满足企业级需求',
      '主题系统可以有效解决文档风格统一的问题',
      'MCP 协议让 AI 调用工具变得标准化和简单',
      '后续可以扩展模板系统，支持更多文档类型'
    ],
    ordered: false
  })

  // 9. 引用风格段落
  await handleWordTool('word_add_paragraph', {
    docId,
    text: '"好的工具应该让复杂的事情变简单，而不是让简单的事情变复杂。" —— 某位智者',
    style: 'quote',
    italic: true
  })

  // 10. 保存文档
  const saveResult = await handleWordTool('word_save', {
    docId,
    addPageNumbers: true,
    addHeader: 'AI 办公文档生成技术调研',
    addFooter: '内部资料 请勿外传'
  })

  console.log('\n✅ 报告生成完成!')
  console.log(JSON.parse(saveResult.content[0].text))
}

generateReport().catch(console.error)
