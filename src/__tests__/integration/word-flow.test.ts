/**
 * Word 文档完整流程集成测试
 */

import { describe, it, expect, afterEach } from 'vitest'
import * as fs from 'fs'
import * as path from 'path'
import * as os from 'os'
import { handleWordTool } from '../../tools/word/index.js'

const tmpDir = path.join(os.tmpdir(), 'office-mcp-test-word')

function extractResult(response: { content: Array<{ text: string }> }) {
  return JSON.parse(response.content[0].text)
}

afterEach(() => {
  if (fs.existsSync(tmpDir)) {
    fs.rmSync(tmpDir, { recursive: true })
  }
})

describe('Word 完整流程', () => {
  it('创建 -> 添加内容 -> 保存', async () => {
    const outputPath = path.join(tmpDir, 'test-flow.docx')

    // 1. 创建文档
    const createRes = extractResult(await handleWordTool('word_create', {
      title: '测试文档',
      outputPath,
      theme: 'business'
    }))
    expect(createRes.success).toBe(true)
    const docId = createRes.docId
    expect(docId).toMatch(/^word_/)

    // 2. 添加标题
    const h1Res = extractResult(await handleWordTool('word_add_heading', {
      docId, text: '第一章 概述', level: 1
    }))
    expect(h1Res.success).toBe(true)

    // 3. 添加段落
    const paraRes = extractResult(await handleWordTool('word_add_paragraph', {
      docId, text: '这是一段测试正文内容。'
    }))
    expect(paraRes.success).toBe(true)

    // 4. 添加表格
    const tableRes = extractResult(await handleWordTool('word_add_table', {
      docId,
      headers: ['姓名', '年龄', '城市'],
      rows: [['张三', '28', '北京'], ['李四', '32', '上海']]
    }))
    expect(tableRes.success).toBe(true)

    // 5. 添加列表
    const listRes = extractResult(await handleWordTool('word_add_list', {
      docId, items: ['第一项', '第二项', '第三项'], ordered: true
    }))
    expect(listRes.success).toBe(true)

    // 6. 添加代码块
    const codeRes = extractResult(await handleWordTool('word_add_code', {
      docId, code: 'console.log("hello")', language: 'javascript'
    }))
    expect(codeRes.success).toBe(true)

    // 7. 保存
    const saveRes = extractResult(await handleWordTool('word_save', {
      docId, addPageNumbers: true, addHeader: '测试文档'
    }))
    expect(saveRes.success).toBe(true)
    expect(fs.existsSync(outputPath)).toBe(true)

    // 验证文件大小合理（至少 1KB）
    const stat = fs.statSync(outputPath)
    expect(stat.size).toBeGreaterThan(1024)
  })

  it('无效 docId 返回错误', async () => {
    const res = extractResult(await handleWordTool('word_add_heading', {
      docId: 'not_exist', text: '标题', level: 1
    }))
    expect(res.success).toBe(false)
    expect(res.message).toContain('不存在')
  })

  it('参数校验失败返回友好错误', async () => {
    const res = extractResult(await handleWordTool('word_create', {
      title: '测试'
      // 缺少 outputPath
    }))
    expect(res.success).toBe(false)
    expect(res.message).toContain('参数校验失败')
  })

  it('各主题均可正常创建和保存', async () => {
    const themes = ['government', 'academic', 'software', 'business', 'alibaba', 'tencent', 'bytedance', 'minimal']

    for (const theme of themes) {
      const outputPath = path.join(tmpDir, `theme-${theme}.docx`)
      const createRes = extractResult(await handleWordTool('word_create', {
        title: `${theme} 主题测试`, outputPath, theme
      }))
      expect(createRes.success).toBe(true)

      extractResult(await handleWordTool('word_add_heading', {
        docId: createRes.docId, text: '标题', level: 1
      }))

      const saveRes = extractResult(await handleWordTool('word_save', {
        docId: createRes.docId
      }))
      expect(saveRes.success).toBe(true)
      expect(fs.existsSync(outputPath)).toBe(true)
    }
  })

  it('quote/note/warning/tip 段落样式', async () => {
    const outputPath = path.join(tmpDir, 'styles.docx')
    const createRes = extractResult(await handleWordTool('word_create', {
      title: '样式测试', outputPath
    }))
    const docId = createRes.docId

    for (const style of ['quote', 'note', 'warning', 'tip']) {
      const res = extractResult(await handleWordTool('word_add_paragraph', {
        docId, text: `${style} 样式内容`, style
      }))
      expect(res.success).toBe(true)
    }

    const saveRes = extractResult(await handleWordTool('word_save', { docId }))
    expect(saveRes.success).toBe(true)
  })
})
