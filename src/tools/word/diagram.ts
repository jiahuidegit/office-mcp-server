/**
 * word_add_diagram - 添加 Mermaid 图表
 */

import * as https from 'https'
import { Paragraph, TextRun, ImageRun, AlignmentType } from 'docx'
import { wordAddDiagramSchema } from '../../schemas/index.js'
import { mermaidThemes } from '../../diagram/mermaid-themes.js'
import { documents } from './store.js'

/** 请求超时时间 */
const REQUEST_TIMEOUT_MS = 15_000
/** 最大重试次数 */
const MAX_RETRIES = 2

/**
 * 单次 HTTP 请求 kroki.io
 */
function krokiRequest(postData: string): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const options = {
      hostname: 'kroki.io',
      port: 443,
      path: '/mermaid/png',
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain',
        'Content-Length': Buffer.byteLength(postData, 'utf-8')
      },
      timeout: REQUEST_TIMEOUT_MS
    }

    const req = https.request(options, (res) => {
      if (res.statusCode !== 200) {
        let errorBody = ''
        res.on('data', chunk => errorBody += chunk)
        res.on('end', () => {
          reject(new Error(`Mermaid 渲染失败: HTTP ${res.statusCode}, ${errorBody.slice(0, 100)}`))
        })
        return
      }
      const chunks: Buffer[] = []
      res.on('data', (chunk: Buffer) => chunks.push(chunk))
      res.on('end', () => resolve(Buffer.concat(chunks)))
      res.on('error', reject)
    })

    req.on('timeout', () => {
      req.destroy()
      reject(new Error(`Mermaid 渲染超时（${REQUEST_TIMEOUT_MS / 1000}s）`))
    })
    req.on('error', reject)
    req.write(postData)
    req.end()
  })
}

/**
 * 将 Mermaid 代码转换为图片 Buffer
 * 使用 kroki.io 在线服务，支持超时和指数退避重试
 */
async function mermaidToImage(mermaidCode: string, theme: string = 'professional'): Promise<Buffer> {
  const themeConfig = mermaidThemes[theme] || mermaidThemes.professional
  const cleanCode = mermaidCode.trim()
  const fullCode = themeConfig ? `${themeConfig}\n${cleanCode}` : cleanCode

  let lastError: Error | undefined
  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    if (attempt > 0) {
      // 指数退避：1s, 2s
      const delay = 1000 * Math.pow(2, attempt - 1)
      await new Promise(r => setTimeout(r, delay))
    }
    try {
      return await krokiRequest(fullCode)
    } catch (err) {
      lastError = err instanceof Error ? err : new Error(String(err))
    }
  }
  throw lastError
}

/** 获取 PNG 图片尺寸 */
function getPngDimensions(buffer: Buffer): { width: number; height: number } {
  if (buffer.length < 24 || buffer[0] !== 0x89 || buffer[1] !== 0x50) {
    return { width: 800, height: 400 }
  }
  return {
    width: buffer.readUInt32BE(16),
    height: buffer.readUInt32BE(20)
  }
}

export async function wordAddDiagram(args: Record<string, unknown>) {
  const parsed = wordAddDiagramSchema.parse(args)
  const { docId, mermaid: mermaidCode, caption, width: targetWidthCm, theme } = parsed

  const doc = documents.get(docId)
  if (!doc) {
    return { success: false, message: `文档 ${docId} 不存在` }
  }

  try {
    const imageBuffer = await mermaidToImage(mermaidCode, theme)
    const { width: origWidth, height: origHeight } = getPngDimensions(imageBuffer)

    // 1cm ~ 37.8px (96 DPI)
    const targetWidthPx = Math.round(targetWidthCm * 37.8)
    const scale = targetWidthPx / origWidth
    const targetHeightPx = Math.round(origHeight * scale)

    const imageParagraph = new Paragraph({
      children: [
        new ImageRun({
          data: imageBuffer,
          transformation: { width: targetWidthPx, height: targetHeightPx }
        })
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 100 }
    })
    doc.children.push(imageParagraph)

    if (caption) {
      const captionParagraph = new Paragraph({
        children: [
          new TextRun({
            text: caption, size: 21, color: '666666',
            font: doc.theme.word.fonts.body
          })
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 }
      })
      doc.children.push(captionParagraph)
    }

    return { success: true, message: `已添加 Mermaid 图表${caption ? `: ${caption}` : ''}` }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error)
    return { success: false, message: `Mermaid 图表渲染失败: ${errMsg}` }
  }
}
