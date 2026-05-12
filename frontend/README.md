# 会意助手 · 前端 (HuiYi Assistant Frontend)

基于 **Vue 3 + Vite** 的单页应用，与仓库根目录 [README](../README.md) 中的后端联调。

## 常用命令

| 命令 | 说明 |
|------|------|
| `npm install` | 安装依赖（有 `package-lock.json` 时 CI 可用 `npm ci`） |
| `npm run dev` | 本地开发（默认 `http://localhost:5173`，API 经 Vite 代理到后端） |
| `npm run build` | 生产构建（若启用 PWA，大体积 chunk 可能需在 `vite.config` 中调高 `maximumFileSizeToCacheInBytes`） |

## 目录提示

- `src/views/DashboardView.vue` — 工作台主体（对话、陪伴、API 配置等 Tab）
- `src/views/GrowthDashboardView.vue` — 情绪与成长看板
- `src/layouts/WorkspaceLayout.vue` — 侧栏 + 主内容区布局
- `src/components/TodoBell.vue` — 待办铃铛（Emoji）

更多说明见仓库根 **README.md**。
