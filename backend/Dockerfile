# OpenClaw 后端（Node 22：内置 node:sqlite；勿降至 Node 18）
FROM node:22-alpine

RUN apk add --no-cache python3 py3-pip \
  && ln -sf /usr/bin/python3 /usr/bin/python

WORKDIR /app

COPY package.json package-lock.json* ./
RUN npm ci --omit=dev

COPY . .
RUN pip3 install --no-cache-dir --break-system-packages -r ./py_service/requirements.txt

ENV NODE_ENV=production
ENV PORT=3000
ENV PYTHON_BIN=python3
# 数据目录默认值；编排时可覆盖（须与卷挂载一致）。业务代码不得硬编码宿主机路径。
ENV OPENCLAW_DATA_DIR=/app/data

RUN mkdir -p /app/data/uploads /app/data/temp /app/logs

EXPOSE 3000

CMD ["node", "server.js"]
