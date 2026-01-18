# ============================================
# Stage 1: Build
# ============================================
FROM node:20-alpine AS builder

WORKDIR /app

# 复制 package 文件
COPY package*.json ./

# 安装依赖（包括 devDependencies，用于构建）
RUN npm ci

# 复制源代码
COPY tsconfig.json ./
COPY src ./src

# 构建
RUN npm run build

# 移除 devDependencies
RUN npm prune --production

# ============================================
# Stage 2: Production
# ============================================
FROM node:20-alpine AS runner

WORKDIR /app

# 安全：使用非 root 用户
RUN addgroup --system --gid 1001 nodejs && \
    adduser --system --uid 1001 mcpuser

# 复制构建产物和生产依赖
COPY --from=builder --chown=mcpuser:nodejs /app/dist ./dist
COPY --from=builder --chown=mcpuser:nodejs /app/node_modules ./node_modules
COPY --from=builder --chown=mcpuser:nodejs /app/package.json ./

# 切换到非 root 用户
USER mcpuser

# 设置环境变量
ENV NODE_ENV=production

# MCP 服务通过 stdio 通信，不需要暴露端口
# 但如果需要 HTTP 模式，可以取消下面的注释
# EXPOSE 3000

# 启动命令
CMD ["node", "dist/index.js"]
