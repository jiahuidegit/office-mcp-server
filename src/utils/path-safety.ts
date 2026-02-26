/**
 * 输出路径安全校验
 */

import * as path from 'path'

const ALLOWED_EXTENSIONS: Record<string, string[]> = {
  word: ['.docx'],
  excel: ['.xlsx']
}

/**
 * 校验输出路径安全性
 * - 禁止路径穿越（..）
 * - 校验文件扩展名
 */
export function validateOutputPath(
  outputPath: string,
  type: 'word' | 'excel'
): { valid: boolean; message?: string } {
  // 禁止路径穿越
  const normalized = path.normalize(outputPath)
  if (normalized.includes('..')) {
    return { valid: false, message: '路径不允许包含 ".."' }
  }

  // 校验扩展名
  const ext = path.extname(normalized).toLowerCase()
  const allowed = ALLOWED_EXTENSIONS[type]
  if (!allowed.includes(ext)) {
    return {
      valid: false,
      message: `文件扩展名必须是 ${allowed.join(' 或 ')}，当前: ${ext || '无'}`
    }
  }

  return { valid: true }
}
