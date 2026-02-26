/**
 * Word 文档内存存储
 */

import { Paragraph, Table } from 'docx'
import { Theme } from '../../styles/themes/index.js'

export interface DocEntry {
  children: Array<Paragraph | Table>
  outputPath: string
  theme: Theme
  title: string
  author?: string
  createdAt: number
}

export const documents = new Map<string, DocEntry>()

/** TTL 清理：30 分钟过期 */
const DOC_TTL_MS = 30 * 60 * 1000

export function cleanExpiredDocs(): void {
  const now = Date.now()
  for (const [id, entry] of documents) {
    if (now - entry.createdAt > DOC_TTL_MS) {
      documents.delete(id)
    }
  }
}
