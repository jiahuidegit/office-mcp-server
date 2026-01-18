#!/usr/bin/env node

/**
 * @erliban/office-mcp-server
 * MCP 服务入口 - 生成专业的 Word 和 Excel 文档
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import {
  CallToolRequestSchema,
  ListToolsRequestSchema
} from '@modelcontextprotocol/sdk/types.js'

import { wordTools, handleWordTool } from './tools/word/index.js'
import { excelTools, handleExcelTool } from './tools/excel/index.js'

// 创建 MCP Server 实例
const server = new Server(
  {
    name: 'office-mcp-server',
    version: '0.1.0'
  },
  {
    capabilities: {
      tools: {}
    }
  }
)

// 注册工具列表
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [...wordTools, ...excelTools]
  }
})

// 处理工具调用
server.setRequestHandler(CallToolRequestSchema, async request => {
  const { name, arguments: args = {} } = request.params

  // Word 工具
  if (name.startsWith('word_')) {
    return handleWordTool(name, args as Record<string, unknown>)
  }

  // Excel 工具
  if (name.startsWith('excel_')) {
    return handleExcelTool(name, args as Record<string, unknown>)
  }

  throw new Error(`Unknown tool: ${name}`)
})

// 启动服务
async function main() {
  const transport = new StdioServerTransport()
  await server.connect(transport)
  console.error('Office MCP Server started')
}

main().catch(err => {
  console.error('Failed to start server:', err)
  process.exit(1)
})
