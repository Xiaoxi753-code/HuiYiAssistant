# OpenClaw 前端（Nginx 托管静态资源 + `/api` 反代至后端容器）
FROM node:22-alpine AS builder
WORKDIR /app
COPY package.json package-lock.json* ./
# npm runs the project postinstall hook during npm ci; copy the hook script first.
COPY scripts ./scripts
RUN npm ci
COPY . .
RUN npm run build

FROM nginx:1.27-alpine
COPY nginx/default.conf /etc/nginx/conf.d/default.conf
COPY --from=builder /app/dist /usr/share/nginx/html
EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]
