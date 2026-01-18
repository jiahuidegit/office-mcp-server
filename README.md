# @erliban/office-mcp-server

MCP 服务 - 生成专业的 Word 和 Excel 文档，支持多种企业级模板和主题风格。

## 特性

- 📝 **Word 文档生成** - 支持标题、段落、表格、代码块、列表等
- 📊 **Excel 表格生成** - 支持数据表、图表、公式、条件格式等
- 🎨 **多种主题风格** - 阿里、腾讯、字节、默认专业风格
- 📋 **预设模板** - 技术文档、周报、月报、PRD、会议纪要等

## 安装

```bash
npm install -g @erliban/office-mcp-server
```

或使用 npx：

```bash
npx @erliban/office-mcp-server
```

## 配置

在 Claude Desktop 配置文件中添加：

```json
{
  "mcpServers": {
    "office": {
      "command": "npx",
      "args": ["@erliban/office-mcp-server"]
    }
  }
}
```

### Docker 方式

```json
{
  "mcpServers": {
    "office": {
      "command": "docker",
      "args": ["run", "-i", "--rm", "ghcr.io/erliban/office-mcp-server:latest"]
    }
  }
}
```

## 可用工具

### Word 工具

| 工具名 | 描述 |
|--------|------|
| `word_create` | 创建 Word 文档 |
| `word_add_heading` | 添加标题 |
| `word_add_paragraph` | 添加段落 |
| `word_add_table` | 添加表格 |
| `word_add_list` | 添加列表 |
| `word_add_code` | 添加代码块 |
| `word_save` | 保存文档 |

### Excel 工具

| 工具名 | 描述 |
|--------|------|
| `excel_create` | 创建工作簿 |
| `excel_add_sheet` | 添加工作表 |
| `excel_write_data` | 写入数据 |
| `excel_add_chart` | 添加图表 |
| `excel_add_formula` | 添加公式 |
| `excel_save` | 保存工作簿 |

## 主题风格

- `alibaba` - 阿里风格（橙色系）
- `tencent` - 腾讯风格（绿色系）
- `bytedance` - 字节风格（黑白极简）
- `default` - 默认专业风格（蓝色系）

## 开发

```bash
# 安装依赖
npm install

# 开发模式
npm run dev

# 构建
npm run build

# 测试
npm test

# 代码检查
npm run lint
```

## License

MIT
