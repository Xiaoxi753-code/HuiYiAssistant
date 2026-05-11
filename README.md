# 会意助手 (Knowing Assistant)

> **你的全能数字研友与智能办公引擎**  
> 将文献、模板、对话、技能广场、成长看板与待办调度统一在一套工作台中——从「整理」到「想清」再到「推到企微」，形成闭环。

---

## 目录

- [核心能力](#-核心能力-features)
- [技术栈](#-技术栈-tech-stack)
- [快速开始](#-快速开始-quick-start)
- [环境配置](#-环境配置)
- [手动启动（可选）](#-手动启动可选)
- [许可说明](#-许可说明)

---

## 核心能力 (Features)

| # | 能力 | 说明 |
|---|------|------|
| 1 | 🧠 **带「思考过程」的对话** | 对模型输出中的 **`<think>`** 做拦截与折叠展示，便于阅读正式回答（见前端对话渲染逻辑）。 |
| 2 | ⏰ **待办与推送调度** | **重复规则**（含自定义星期）、**稍后提醒**、修改提醒时间；**企业微信 / 飞书** Webhook 在应用内配置后，由后端轮询推送（见 `server.js` 与「外联推送配置」）。 |
| 3 | 📁 **模板与资料库** | **Word 模板**填充与预览、**PDF/Word** 等解析、个人资料库；上传侧对文件名做了规范化处理。 |
| 4 | 🎨 **工作台体验** | **可调整宽度的侧栏**、**对话区会话列表宽度与收起**、概览与图表（**ECharts**）等。 |
| 5 | 🏆 **技能与成长** | 独立路由：**技能广场**（`SkillLibraryView`）、**情绪与成长看板**（`GrowthDashboardView`），与主工作台协同。 |

> **说明**：侧栏中部分模块（例如历史上的「PPT 大纲」等）若已用 `v-if="false"` 等方式关闭入口，则默认用户界面不再展示；后端仍可能保留对应依赖或接口，以仓库内代码为准。

---

## 技术栈 (Tech Stack)

| 层级 | 技术 |
|------|------|
| 前端 | **Vue 3**、**Vite**、**Vue Router**、**Pinia**、**Element Plus**、**ECharts**、**PWA（vite-plugin-pwa）** |
| 后端 | **Node.js**、**Express**、**SQLite**（`node:sqlite` / `DatabaseSync`）、**dotenv** |
| 文档与模板 | **docx-preview**、**docxtemplater**、**pdf-parse**、**mammoth**；服务端另含 **pptxgenjs**（若未开放前端入口则仅作预留能力） |
| AI 与集成 | **OpenAI 兼容 API**（百炼 / DeepSeek 等）、**axios** |

---

## 快速开始 (Quick Start)

### Windows（推荐）

> **双击项目根目录下的 `start.bat`，即可一键拉起前后端。**

1. 安装 [Node.js](https://nodejs.org/)（建议 **≥ 22**，与 `backend/package.json` 中 `engines` 一致）。  
2. 双击 **`start.bat`**。  
3. 将打开 **两个控制台**（后端 + 前端）；脚本会 **轮询 Vite 默认端口 `5173`**，在端口可连接后（与终端中出现 `VITE … ready` / `Local:` 一致）再自动打开浏览器。

### macOS / Linux

```bash
chmod +x start.sh
./start.sh
```

- **macOS**：`osascript` 唤起 **Terminal**。  
- **Linux**：优先 **gnome-terminal**，其次 **x-terminal-emulator** / **xterm** / **konsole**。  
- 同样在 **检测到 Vite 端口就绪** 后再 `open` / `xdg-open`。  
- 自定义端口：`VITE_PORT=5174 ./start.sh`

---

## 环境配置

在 **`backend` 目录**下创建 **`.env`**（由 `server.js` 加载），常用项如下：

| 变量 | 说明 |
|------|------|
| `PORT` | HTTP 端口，默认 `3000`（与 `frontend/vite.config.js` 里 `/api` 代理目标一致）。 |
| `OPENCLAW_DB_PATH` | 可选，SQLite 数据库文件的**完整路径**；不设则使用数据目录下的默认库文件。 |
| `SQLITE_DB_FILE` | 可选，仅文件名（如 `database.sqlite`），拼在数据目录下。 |
| `BAILIAN_API_KEY` | 百炼 / DashScope 等 **OpenAI 兼容** Key。 |
| `DEEPSEEK_API_KEY` | 可选；与上百炼 Key 按代码逻辑选用。 |
| `BAILIAN_BASE_URL` | 可选，默认 `https://dashscope.aliyuncs.com/compatible-mode/v1`。 |
| `BAILIAN_MODEL` | 可选，默认 `qwen-plus`。 |
| `OPENCLAW_DATA_DIR` | 可选，数据根目录。 |
| `TODO_PUSH_POLL_MS` | 可选，待办推送轮询间隔（毫秒），默认 `10000`。 |
| `LOG_LEVEL` / `OPENCLAW_LOG_DIR` | 可选，见 `backend/src/utils/logger.js`。 |

### 企业微信 / 飞书 Webhook

> 在应用内 **「外联推送配置」**（或历史文案中的 Webhook 配置页）将 **机器人 Webhook URL** 写入系统配置（数据库 `system_configs`），与 `.env` 中的大模型 Key **分开管理**。  
> 配置生效后，待办到点可走 **企业微信 / 飞书** 推送（以当前 `server.js` 实现为准）。

---

## 手动启动（可选）

```bash
cd backend && npm install && npm run dev
```

```bash
cd frontend && (test -f package-lock.json && npm ci || npm install) && npm run dev
```

浏览器访问：**`http://localhost:5173`**（若修改了 Vite 端口则替换为对应地址）。

---

## 容器部署（可选）

仓库根目录提供 **`Dockerfile`**、**`docker-compose.yml`** 及前后端各自 Docker 配置；需要容器化部署时可从 compose 入手按需裁剪。

---

## 许可说明

当前仓库根目录 **未默认附带 `LICENSE` 文件**。使用、分发与二次开发条件 **以项目维护者另行提供的许可或书面约定为准**；在获得明确授权前，请勿假定本项目为公有领域或特定开源协议。

---

**会意助手 (Knowing Assistant)** — *Think clearly. Ship calmly.*
