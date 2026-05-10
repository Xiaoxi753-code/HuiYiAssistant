# 单镜像一体化构建（API + 静态前端同容器）。双容器编排见仓库根目录 `docker-compose.yml`。
# 说明：仓库使用 Node 内置 `node:sqlite`，需 Node 22+。若文档曾写 Node 18，请以本镜像为准。
# 前端构建 + 后端运行 + Python（模板转 PDF 的 docxtpl 脚本；pdf-parse 为纯 Node 依赖）
FROM node:22-alpine AS frontend-builder
WORKDIR /build/frontend
COPY frontend/package*.json ./
# npm runs the frontend postinstall hook during npm ci; copy the hook script first.
COPY frontend/scripts ./scripts
RUN npm ci
COPY frontend/ ./
RUN npm run build

FROM node:22-alpine AS production
RUN apk add --no-cache python3 py3-pip \
  && ln -sf /usr/bin/python3 /usr/bin/python

WORKDIR /app

COPY backend/package*.json ./
RUN npm ci --omit=dev

COPY backend/ ./
RUN rm -rf ./public && mkdir -p ./public
COPY --from=frontend-builder /build/frontend/dist ./public

RUN pip3 install --no-cache-dir --break-system-packages -r ./py_service/requirements.txt

ENV NODE_ENV=production
ENV PORT=3000
# 默认数据根目录；生产编排请用环境变量覆盖并与卷挂载对齐（代码侧不写死宿主机路径）。
ENV OPENCLAW_DATA_DIR=/app/data
ENV PYTHON_BIN=python3

RUN mkdir -p /app/data/uploads /app/data/temp

EXPOSE 3000

CMD ["node", "server.js"]
