import './crash-guard.js';
import logger from './src/utils/logger.js';
import axios from 'axios';
import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import path from 'path';
import fs from 'fs';
import crypto from 'crypto';
import { spawn } from 'child_process';
import { fileURLToPath } from 'url';
import { DatabaseSync } from 'node:sqlite';
import multer from 'multer';
import mammoth from 'mammoth';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import {
  sendChatStream,
  parseSsePayloads,
  pingChatModel,
  chatCompletionsUrl,
  buildHeadersFromConfig,
  normalizeModelId,
  summarizePaper,
  postChatCompletion,
  extractTextFromResponse,
  withThinkingParams,
  analyzeUserProfileFromTranscript,
  mergeIncrementalProfilePatch,
  appendEmotionHistoryPoint,
  sanitizeEmotionHistoryArray,
  generateWeeklyGrowthReview,
} from './src/services/openclawService.js';
import { extractPaperTextFromPath, parsePdfBuffer } from './src/paperExtract.js';
import { createToolExecutor, getOpenClawChatTools } from './tools.js';
import { runAgentChatToSse } from './src/openclawAgentLoop.js';
import {
  generateMeetingMinutesMarkdown,
  extractTodosJson,
  insertTodoRows,
  insertTodoRowsFromAgentTaskList,
  generatePptOutlineMarkdown,
} from './src/officeProductivity.js';
import { attachRealtimeAsrWss } from './src/realtimeAsrBrowserGateway.js';
import { transcribeUploadedAudioFile } from './src/asrService.js';
import { DEFAULT_REALTIME_WS, assertSafeDashscopeRealtimeUrl } from './src/asrRealtimeService.js';
import { markdownToPptxBuffer } from './src/pptExport.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
dotenv.config({ path: path.join(__dirname, '.env') });
const app = express();
const port = Number.parseInt(String(process.env.PORT || ''), 10) || 3000;
/**
 * 数据根目录：禁止在代码中写死盘符路径；默认使用后端包内相对目录 `data/`（便于 Docker 挂载卷或本地部署）。
 * 可通过 OPENCLAW_DATA_DIR 指定绝对路径（如容器内 /app/data）或相对当前工作目录的相对路径。
 */
function resolveDataRoot() {
  const raw = String(process.env.OPENCLAW_DATA_DIR || '').trim();
  if (!raw) {
    return path.resolve(__dirname, 'data');
  }
  return path.isAbsolute(raw) ? path.normalize(raw) : path.resolve(process.cwd(), raw);
}
const dataRoot = resolveDataRoot();

/**
 * 解析可选绝对/相对路径（用于 Docker 分别挂载 uploads 与 SQLite 文件）。
 * @param {string} envKey
 * @param {string} fallback
 */
function resolveOptionalPath(envKey, fallback) {
  const raw = String(process.env[envKey] || '').trim();
  if (!raw) {
    return fallback;
  }
  return path.isAbsolute(raw) ? path.normalize(raw) : path.resolve(process.cwd(), raw);
}

/** SQLite 文件名（当未设置 OPENCLAW_DB_PATH 时置于 dataRoot 下），如 database.sqlite、test.db */
function resolveSqliteFileName() {
  const raw = String(process.env.SQLITE_DB_FILE || 'database.sqlite').trim() || 'database.sqlite';
  const base = path.basename(raw).replace(/[^\w.\-]+/g, '_') || 'database.sqlite';
  if (!base.endsWith('.sqlite') && !base.endsWith('.db')) {
    return 'database.sqlite';
  }
  return base;
}

const uploadsDir = resolveOptionalPath('OPENCLAW_UPLOADS_DIR', path.join(dataRoot, 'uploads'));
const tempDir = resolveOptionalPath('OPENCLAW_TEMP_DIR', path.join(dataRoot, 'temp'));
const dbPath = resolveOptionalPath('OPENCLAW_DB_PATH', path.join(dataRoot, resolveSqliteFileName()));
const pythonScript = path.join(__dirname, 'py_service', 'convert.py');
const pythonBin = String(process.env.PYTHON_BIN || (process.platform === 'win32' ? 'python' : 'python3')).trim() || 'python3';
/** 待办外联推送轮询间隔（默认 10s；仍可由环境变量调节；Webhook 地址存 system_configs） */
const TODO_PUSH_POLL_MS = Math.min(300000, Math.max(10000, Number.parseInt(String(process.env.TODO_PUSH_POLL_MS || ''), 10) || 10000));

fs.mkdirSync(uploadsDir, { recursive: true });
fs.mkdirSync(tempDir, { recursive: true });
fs.mkdirSync(path.dirname(dbPath), { recursive: true });

app.use(cors());
app.use(express.json({ limit: '25mb' }));

const db = new DatabaseSync(dbPath);
try {
  db.exec('PRAGMA foreign_keys = ON;');
} catch {
  /* ignore */
}
db.exec(`
  CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT UNIQUE,
    password TEXT,
    name TEXT
  );

  CREATE TABLE IF NOT EXISTS templates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    filename TEXT NOT NULL,
    original_name TEXT NOT NULL,
    uploaded_at TEXT NOT NULL DEFAULT (datetime('now'))
  );
`);
db.exec('CREATE UNIQUE INDEX IF NOT EXISTS idx_templates_filename ON templates(filename);');

db.exec(`
  CREATE TABLE IF NOT EXISTS model_configs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    api_key TEXT NOT NULL DEFAULT '',
    base_url TEXT NOT NULL,
    model_name TEXT NOT NULL,
    priority INTEGER NOT NULL DEFAULT 0,
    is_enabled INTEGER NOT NULL DEFAULT 1,
    type TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    CHECK (type IN ('openai', 'aliyun', 'openclaw'))
  );
`);
db.exec('CREATE INDEX IF NOT EXISTS idx_model_configs_pick ON model_configs (is_enabled, priority, id);');

/** 无损升级：为 users 增加 JSON 画像列（旧库无此列时执行一次 ALTER，不删库、不改原有行数据） */
function ensureUsersProfileColumn() {
  try {
    const cols = db.prepare('PRAGMA table_info(users)').all();
    const has = cols.some((c) => c.name === 'user_profile');
    if (!has) {
      db.exec("ALTER TABLE users ADD COLUMN user_profile TEXT DEFAULT '{}'");
    }
  } catch (e) {
    logger.error('[db migrate] ensure user_profile column failed', e);
    throw e;
  }
}
ensureUsersProfileColumn();

/** 用户自定义 Skill（可复用 system 片段） */
function ensureSkillsTable() {
  try {
    db.exec(`
      CREATE TABLE IF NOT EXISTS skills (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        description TEXT NOT NULL DEFAULT '',
        system_prompt TEXT NOT NULL DEFAULT '',
        user_email TEXT,
        created_at TEXT NOT NULL DEFAULT (datetime('now'))
      );
    `);
    db.exec('CREATE INDEX IF NOT EXISTS idx_skills_user_email ON skills(user_email);');
  } catch (e) {
    logger.error('[db migrate] skills table failed', e);
    throw e;
  }
}
ensureSkillsTable();

/** Skill 表扩展：图标、唤醒词、用量统计 */
function ensureSkillsExtraColumns() {
  try {
    const cols = db.prepare('PRAGMA table_info(skills)').all();
    const names = new Set(cols.map((c) => c.name));
    if (!names.has('emoji')) {
      db.exec("ALTER TABLE skills ADD COLUMN emoji TEXT DEFAULT '✨'");
    }
    if (!names.has('wake_word')) {
      db.exec('ALTER TABLE skills ADD COLUMN wake_word TEXT DEFAULT ""');
    }
    if (!names.has('use_count')) {
      db.exec('ALTER TABLE skills ADD COLUMN use_count INTEGER DEFAULT 0');
    }
  } catch (e) {
    logger.error('[db migrate] skills extra columns', e);
    throw e;
  }
}
ensureSkillsExtraColumns();

/** 近期对话轮次缓冲（用于自动画像建模） */
function ensureProfileTurnBufferTable() {
  try {
    db.exec(`
      CREATE TABLE IF NOT EXISTS profile_turn_buffer (
        user_email TEXT PRIMARY KEY,
        turns_json TEXT NOT NULL DEFAULT '[]',
        updated_at TEXT NOT NULL DEFAULT (datetime('now'))
      );
    `);
  } catch (e) {
    logger.error('[db migrate] profile_turn_buffer', e);
    throw e;
  }
}
ensureProfileTurnBufferTable();

/** 用户级汇总统计（对话轮次、文献处理次数），用于工作台概览 */
function ensureUserStatsTable() {
  try {
    db.exec(`
      CREATE TABLE IF NOT EXISTS user_stats (
        user_email TEXT PRIMARY KEY,
        conversation_turns INTEGER NOT NULL DEFAULT 0,
        papers_processed INTEGER NOT NULL DEFAULT 0
      );
    `);
  } catch (e) {
    logger.error('[db migrate] user_stats', e);
    throw e;
  }
}
ensureUserStatsTable();

/** 每周成长与情绪回溯报告历史（每次生成一条） */
function ensureWeeklyReviewReportsTable() {
  try {
    db.exec(`
      CREATE TABLE IF NOT EXISTS weekly_review_reports (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_email TEXT NOT NULL,
        week_label TEXT NOT NULL DEFAULT '',
        report_body TEXT NOT NULL,
        created_at TEXT NOT NULL DEFAULT (datetime('now'))
      );
    `);
    db.exec(
      'CREATE INDEX IF NOT EXISTS idx_weekly_reports_user_time ON weekly_review_reports (user_email, created_at DESC);',
    );
  } catch (e) {
    logger.error('[db migrate] weekly_review_reports', e);
    throw e;
  }
}
ensureWeeklyReviewReportsTable();

/** Word 模板归属用户（NULL 表示旧版全局模板，fill-stream 在无 user_email 时不校验） */
function ensureTemplatesUserEmailColumn() {
  try {
    const cols = db.prepare('PRAGMA table_info(templates)').all();
    const names = new Set(cols.map((c) => c.name));
    if (!names.has('user_email')) {
      db.exec('ALTER TABLE templates ADD COLUMN user_email TEXT');
    }
  } catch (e) {
    logger.error('[db migrate] templates.user_email', e);
    throw e;
  }
}
ensureTemplatesUserEmailColumn();

/** 模板库组织字段：文件夹与标签，避免把路径塞进显示名称 */
function ensureTemplatesOrganizationColumns() {
  try {
    const cols = db.prepare('PRAGMA table_info(templates)').all();
    const names = new Set(cols.map((c) => c.name));
    if (!names.has('folder_path')) {
      db.exec('ALTER TABLE templates ADD COLUMN folder_path TEXT DEFAULT ""');
    }
    if (!names.has('tags_json')) {
      db.exec("ALTER TABLE templates ADD COLUMN tags_json TEXT DEFAULT '[]'");
    }
  } catch (e) {
    logger.error('[db migrate] templates organization columns', e);
    throw e;
  }
}
ensureTemplatesOrganizationColumns();


/** 个人资料库：集中归档用户上传的文档、会议纪要与可复用资料 */
function ensureKnowledgeItemsTable() {
  try {
    db.exec(`
      CREATE TABLE IF NOT EXISTS knowledge_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_email TEXT NOT NULL,
        title TEXT NOT NULL,
        item_type TEXT NOT NULL DEFAULT 'document',
        source_kind TEXT NOT NULL DEFAULT 'manual_upload',
        original_name TEXT NOT NULL DEFAULT '',
        stored_filename TEXT NOT NULL DEFAULT '',
        mime_type TEXT NOT NULL DEFAULT '',
        folder_path TEXT NOT NULL DEFAULT '',
        tags_json TEXT NOT NULL DEFAULT '[]',
        content_text TEXT NOT NULL DEFAULT '',
        summary_text TEXT NOT NULL DEFAULT '',
        metadata_json TEXT NOT NULL DEFAULT '{}',
        created_at TEXT NOT NULL DEFAULT (datetime('now')),
        updated_at TEXT NOT NULL DEFAULT (datetime('now'))
      );
    `);
    db.exec('CREATE INDEX IF NOT EXISTS idx_knowledge_items_user_time ON knowledge_items (user_email, created_at DESC);');
    db.exec('CREATE INDEX IF NOT EXISTS idx_knowledge_items_user_folder ON knowledge_items (user_email, folder_path, item_type);');
  } catch (e) {
    logger.error('[db migrate] knowledge_items', e);
    throw e;
  }
}
ensureKnowledgeItemsTable();

/** 待办与提醒 */
function ensureTodosTable() {
  try {
    db.exec(`
      CREATE TABLE IF NOT EXISTS todos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_email TEXT NOT NULL,
        title TEXT NOT NULL,
        remind_at TEXT NOT NULL,
        notes TEXT,
        source TEXT,
        done INTEGER NOT NULL DEFAULT 0,
        repeat_mode TEXT NOT NULL DEFAULT 'once',
        created_at TEXT NOT NULL DEFAULT (datetime('now'))
      );
    `);
    db.exec('CREATE INDEX IF NOT EXISTS idx_todos_user_remind ON todos (user_email, remind_at);');
    db.exec('CREATE INDEX IF NOT EXISTS idx_todos_user_done ON todos (user_email, done);');
  } catch (e) {
    logger.error('[db migrate] todos', e);
    throw e;
  }
}
ensureTodosTable();

/** 待办推送字段：避免重复发送微信提醒 */
function ensureTodosPushColumns() {
  try {
    const cols = db.prepare('PRAGMA table_info(todos)').all();
    const names = new Set(cols.map((c) => c.name));
    if (!names.has('pushed_wecom_at')) {
      db.exec('ALTER TABLE todos ADD COLUMN pushed_wecom_at TEXT');
    }
    if (!names.has('pushed_wecom_status')) {
      db.exec('ALTER TABLE todos ADD COLUMN pushed_wecom_status TEXT');
    }
  } catch (e) {
    logger.error('[db migrate] todos push columns', e);
    throw e;
  }
}
ensureTodosPushColumns();

/** 飞书待办推送去重字段 */
function ensureTodosFeishuPushColumns() {
  try {
    const cols = db.prepare('PRAGMA table_info(todos)').all();
    const names = new Set(cols.map((c) => c.name));
    if (!names.has('pushed_feishu_at')) {
      db.exec('ALTER TABLE todos ADD COLUMN pushed_feishu_at TEXT');
    }
    if (!names.has('pushed_feishu_status')) {
      db.exec('ALTER TABLE todos ADD COLUMN pushed_feishu_status TEXT');
    }
  } catch (e) {
    logger.error('[db migrate] todos feishu push columns', e);
    throw e;
  }
}
ensureTodosFeishuPushColumns();

/** 待办外联：推送状态（pending → notified，与 pushed_* 列共同防止重复发送） */
function ensureTodosNotifyStatusColumn() {
  try {
    const cols = db.prepare('PRAGMA table_info(todos)').all();
    const names = new Set(cols.map((c) => c.name));
    if (!names.has('status')) {
      db.exec(`ALTER TABLE todos ADD COLUMN status TEXT NOT NULL DEFAULT 'pending';`);
    }
    db.exec(`
      UPDATE todos SET status = 'notified'
      WHERE (TRIM(COALESCE(pushed_wecom_at, '')) != '' OR TRIM(COALESCE(pushed_feishu_at, '')) != '')
        AND (status IS NULL OR status = '' OR status = 'pending');
    `);
  } catch (e) {
    logger.error('[db migrate] todos notify status', e);
    throw e;
  }
}
ensureTodosNotifyStatusColumn();

/** 推送失败重试计数，超过阈值后不再进入调度，避免卡死队列 */
const TODO_PUSH_MAX_RETRY = 20;
const TODO_REMIND_CANDIDATE_SQL_LIMIT = 200;

function ensureTodosPushRetryCountColumn() {
  try {
    const cols = db.prepare('PRAGMA table_info(todos)').all();
    const names = new Set(cols.map((c) => c.name));
    if (!names.has('push_retry_count')) {
      db.exec('ALTER TABLE todos ADD COLUMN push_retry_count INTEGER NOT NULL DEFAULT 0');
    }
  } catch (e) {
    logger.error('[db migrate] todos push_retry_count', e);
    throw e;
  }
}
ensureTodosPushRetryCountColumn();

/** 循环闹钟：单次 / 每天 / 工作日（老库缺列时 ALTER，失败吞掉避免阻断启动） */
function ensureTodosRepeatModeColumn() {
  try {
    const cols = db.prepare('PRAGMA table_info(todos)').all();
    const names = new Set(cols.map((c) => c.name));
    if (!names.has('repeat_mode')) {
      try {
        db.exec(`ALTER TABLE todos ADD COLUMN repeat_mode TEXT DEFAULT 'once'`);
      } catch (inner) {
        logger.warn('[db migrate] todos repeat_mode ALTER skipped', inner?.message || inner);
      }
    }
  } catch (e) {
    logger.error('[db migrate] todos repeat_mode', e);
  }
}
ensureTodosRepeatModeColumn();

/** 自定义工作日重复：JSON 数组，1–7 表示周一至周日 */
function ensureTodosRepeatCustomDaysColumn() {
  try {
    const cols = db.prepare('PRAGMA table_info(todos)').all();
    const names = new Set(cols.map((c) => c.name));
    if (!names.has('repeat_custom_days')) {
      db.exec(`ALTER TABLE todos ADD COLUMN repeat_custom_days TEXT DEFAULT '[]'`);
    }
  } catch (e) {
    logger.error('[db migrate] todos repeat_custom_days', e);
  }
}
ensureTodosRepeatCustomDaysColumn();

/** 循环待办：用户当日在界面点击「完成」后记录日历日（Asia/Shanghai），用于铃铛统计「今日清爽」 */
function ensureTodosLastCompletedDateColumn() {
  try {
    const cols = db.prepare('PRAGMA table_info(todos)').all();
    const names = new Set(cols.map((c) => c.name));
    if (!names.has('last_completed_date')) {
      db.exec('ALTER TABLE todos ADD COLUMN last_completed_date TEXT');
    }
  } catch (e) {
    logger.error('[db migrate] todos last_completed_date', e);
  }
}
ensureTodosLastCompletedDateColumn();

/** AI 回复点赞 / 评分（用于工作台趋势图） */
function ensureAiRatingsTable() {
  try {
    db.exec(`
      CREATE TABLE IF NOT EXISTS ai_ratings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        score INTEGER NOT NULL,
        created_at TEXT NOT NULL DEFAULT (datetime('now'))
      );
    `);
    db.exec('CREATE INDEX IF NOT EXISTS idx_ai_ratings_created ON ai_ratings (created_at);');
  } catch (e) {
    logger.error('[db migrate] ai_ratings', e);
    throw e;
  }
}
ensureAiRatingsTable();

/** 强一致键名（INSERT OR REPLACE）；旧键仅用于读取兼容 */
const WEBHOOK_KEY_WECOM = 'wecom_webhook';
const WEBHOOK_KEY_FEISHU = 'feishu_webhook';
const WEBHOOK_KEY_WECOM_LEGACY = 'webhook_wecom_url';
const WEBHOOK_KEY_FEISHU_LEGACY = 'webhook_feishu_url';

function ensureSystemConfigsTable() {
  try {
    db.exec(`
      CREATE TABLE IF NOT EXISTS system_configs (
        config_key TEXT PRIMARY KEY,
        config_value TEXT
      );
    `);
  } catch (e) {
    logger.error('[db migrate] system_configs', e);
    throw e;
  }
}
ensureSystemConfigsTable();

function getSystemConfig(key) {
  const row = db.prepare('SELECT config_value FROM system_configs WHERE config_key = ?').get(String(key || ''));
  return row && row.config_value != null ? String(row.config_value).trim() : '';
}

function setSystemConfig(key, value) {
  db.prepare('INSERT OR REPLACE INTO system_configs (config_key, config_value) VALUES (?, ?)').run(
    String(key || ''),
    String(value ?? ''),
  );
}

function getWebhookWecomUrl() {
  const v = getSystemConfig(WEBHOOK_KEY_WECOM);
  if (v) {
    return v;
  }
  return getSystemConfig(WEBHOOK_KEY_WECOM_LEGACY);
}

function getWebhookFeishuUrl() {
  const v = getSystemConfig(WEBHOOK_KEY_FEISHU);
  if (v) {
    return v;
  }
  return getSystemConfig(WEBHOOK_KEY_FEISHU_LEGACY);
}

/** 对话会话与消息（多会话历史） */
function ensureChatSessionsTables() {
  try {
    db.exec(`
      CREATE TABLE IF NOT EXISTS chat_sessions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_email TEXT NOT NULL,
        title TEXT NOT NULL DEFAULT '新对话',
        created_at TEXT NOT NULL DEFAULT (datetime('now')),
        updated_at TEXT NOT NULL DEFAULT (datetime('now'))
      );
    `);
    db.exec('CREATE INDEX IF NOT EXISTS idx_chat_sessions_user_updated ON chat_sessions (user_email, updated_at DESC);');
    db.exec(`
      CREATE TABLE IF NOT EXISTS chat_messages (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id INTEGER NOT NULL,
        role TEXT NOT NULL,
        content TEXT NOT NULL DEFAULT '',
        extra_json TEXT,
        sort_order INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL DEFAULT (datetime('now')),
        FOREIGN KEY (session_id) REFERENCES chat_sessions(id) ON DELETE CASCADE
      );
    `);
    db.exec('CREATE INDEX IF NOT EXISTS idx_chat_messages_session_order ON chat_messages (session_id, sort_order);');
  } catch (e) {
    logger.error('[db migrate] chat_sessions', e);
    throw e;
  }
}
ensureChatSessionsTables();

/** 对话会话扩展：记录首轮使用的人格 / Skill 快照，便于历史复现 */
function ensureChatSessionsExtraColumns() {
  try {
    const cols = db.prepare('PRAGMA table_info(chat_sessions)').all();
    const names = new Set(cols.map((c) => c.name));
    if (!names.has('persona_id')) {
      db.exec('ALTER TABLE chat_sessions ADD COLUMN persona_id TEXT DEFAULT ""');
    }
    if (!names.has('persona_label')) {
      db.exec('ALTER TABLE chat_sessions ADD COLUMN persona_label TEXT DEFAULT ""');
    }
    if (!names.has('persona_emoji')) {
      db.exec('ALTER TABLE chat_sessions ADD COLUMN persona_emoji TEXT DEFAULT ""');
    }
    if (!names.has('skill_id')) {
      db.exec('ALTER TABLE chat_sessions ADD COLUMN skill_id TEXT DEFAULT ""');
    }
    if (!names.has('skill_name')) {
      db.exec('ALTER TABLE chat_sessions ADD COLUMN skill_name TEXT DEFAULT ""');
    }
  } catch (e) {
    logger.error('[db migrate] chat_sessions extra columns', e);
    throw e;
  }
}
ensureChatSessionsExtraColumns();

/** 模板填充历史：每次 AI 填充完成后保留变量、上下文与生成文件，支持回溯复用 */
function ensureTemplateFillRecordsTable() {
  try {
    db.exec(`
      CREATE TABLE IF NOT EXISTS template_fill_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        template_id INTEGER,
        user_email TEXT NOT NULL,
        template_name TEXT NOT NULL DEFAULT '',
        variables_json TEXT NOT NULL DEFAULT '{}',
        context_json TEXT NOT NULL DEFAULT '{}',
        filled_filename TEXT NOT NULL DEFAULT '',
        created_at TEXT NOT NULL DEFAULT (datetime('now')),
        FOREIGN KEY (template_id) REFERENCES templates(id) ON DELETE SET NULL
      );
    `);
    db.exec(
      'CREATE INDEX IF NOT EXISTS idx_template_fill_records_user_time ON template_fill_records (user_email, created_at DESC);',
    );
  } catch (e) {
    logger.error('[db migrate] template_fill_records', e);
    throw e;
  }
}
ensureTemplateFillRecordsTable();

function buildFallbackChatTitle(text) {
  const title = String(text || '')
    .replace(/[#*_`>\[\]{}()]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 48);
  return title || '新对话';
}

function sanitizeGeneratedChatTitle(text, fallback) {
  const raw = String(text || '')
    .replace(/```[\s\S]*?```/g, '')
    .replace(/[「」“”\"'`#*_>]/g, '')
    .replace(/^标题[:：]\s*/i, '')
    .replace(/\s+/g, ' ')
    .trim();
  const first = raw.split(/[\r\n。！？!?]/).find(Boolean) || raw;
  const title = first.slice(0, 28).trim();
  return title || fallback || '新对话';
}

async function generateChatTitleFromFirstTurn(userText, assistantText) {
  const fallback = buildFallbackChatTitle(userText);
  const modelConfig = rowToModelConfig(getBestEnabledModelConfig());
  if (!modelConfig) {
    return fallback;
  }
  try {
    const data = await postChatCompletion(modelConfig, {
      model: normalizeModelId(modelConfig.model_name),
      messages: [
        {
          role: 'system',
          content: '你是对话标题生成器。请用中文给这轮对话生成一个简洁标题，8到18个汉字，不要引号、不要句号、不要解释。',
        },
        {
          role: 'user',
          content: `用户第一问：${String(userText || '').slice(0, 1200)}\n\n助手首答摘要：${String(assistantText || '').slice(0, 1200)}`,
        },
      ],
      temperature: 0.1,
    });
    return sanitizeGeneratedChatTitle(extractTextFromResponse(data), fallback);
  } catch (error) {
    logger.warn('[chat/title] fallback', error?.message || error);
    return fallback;
  }
}

function sanitizeTemplateTags(raw) {
  const input = Array.isArray(raw)
    ? raw
    : typeof raw === 'string'
      ? (() => {
          try {
            const parsed = JSON.parse(raw);
            return Array.isArray(parsed) ? parsed : raw.split(/[，,;；\s]+/);
          } catch {
            return raw.split(/[，,;；\s]+/);
          }
        })()
      : [];
  return Array.from(
    new Set(
      input
        .map((x) => String(x || '').trim().slice(0, 40))
        .filter(Boolean)
        .slice(0, 12),
    ),
  );
}

function parseTemplateTagsJson(raw) {
  try {
    return sanitizeTemplateTags(JSON.parse(String(raw || '[]')));
  } catch {
    return [];
  }
}


function sanitizeLibraryFolderPath(raw) {
  return String(raw || '')
    .replace(/\\/g, '/')
    .split('/')
    .map((part) => part.trim().replace(/[<>:"|?*]+/g, '').slice(0, 80))
    .filter(Boolean)
    .join('/')
    .slice(0, 500);
}

function parseJsonObjectSafe(raw, fallback = {}) {
  try {
    const v = JSON.parse(String(raw || ''));
    return v && typeof v === 'object' && !Array.isArray(v) ? v : fallback;
  } catch {
    return fallback;
  }
}

function parseLibraryRow(row) {
  return {
    id: row.id,
    title: row.title,
    itemType: row.item_type,
    sourceKind: row.source_kind,
    originalName: row.original_name,
    storedFilename: row.stored_filename,
    mimeType: row.mime_type,
    folderPath: row.folder_path,
    tags: parseTemplateTagsJson(row.tags_json),
    contentText: row.content_text,
    summaryText: row.summary_text,
    metadata: parseJsonObjectSafe(row.metadata_json, {}),
    createdAt: row.created_at,
    updatedAt: row.updated_at,
    canDownload: !!row.stored_filename,
  };
}

function inferKnowledgeFolder(item) {
  const text = `${item.title || ''} ${item.originalName || ''} ${item.summaryText || ''} ${item.contentText || ''}`.toLowerCase();
  if (item.itemType === 'meeting' || /会议|纪要|meeting|transcript/.test(text)) return '会议纪要';
  if (/论文|文献|paper|article|研究|摘要|pdf/.test(text)) return '文献资料';
  if (/合同|协议|法务|采购/.test(text)) return '合同法务';
  if (/周报|日报|月报|汇报|复盘/.test(text)) return '工作汇报';
  if (/需求|产品|prd|原型/.test(text)) return '产品资料';
  return '待整理';
}

function insertKnowledgeItem(input) {
  const userEmail = String(input.userEmail || '').trim();
  if (!userEmail) return null;
  const title = String(input.title || input.originalName || '未命名资料').trim().slice(0, 240) || '未命名资料';
  const info = db
    .prepare(
      `INSERT INTO knowledge_items
       (user_email, title, item_type, source_kind, original_name, stored_filename, mime_type, folder_path, tags_json, content_text, summary_text, metadata_json)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    )
    .run(
      userEmail,
      title,
      String(input.itemType || 'document').slice(0, 40),
      String(input.sourceKind || 'manual_upload').slice(0, 60),
      String(input.originalName || '').slice(0, 260),
      String(input.storedFilename || '').replace(/\\/g, '/').slice(0, 600),
      String(input.mimeType || '').slice(0, 120),
      sanitizeLibraryFolderPath(input.folderPath || ''),
      JSON.stringify(sanitizeTemplateTags(input.tags || [])),
      String(input.contentText || '').slice(0, 120000),
      String(input.summaryText || '').slice(0, 120000),
      JSON.stringify(input.metadata && typeof input.metadata === 'object' ? input.metadata : {}),
    );
  return Number(info.lastInsertRowid);
}

function resolveKnowledgeStoredFile(storedFilename) {
  const rel = String(storedFilename || '').replace(/\\/g, '/').replace(/^\/+/, '');
  if (!rel) return '';
  const abs = path.resolve(uploadsDir, rel);
  const root = path.resolve(uploadsDir);
  if (abs !== root && !abs.startsWith(root + path.sep)) {
    return '';
  }
  return abs;
}

function scheduleChatTodoExtraction({ userEmail, userMessage, assistantMessage, sessionId }) {
  const ue = String(userEmail || '').trim();
  const u = String(userMessage || '').trim().slice(0, 12000);
  const a = String(assistantMessage || '').trim().slice(0, 16000);
  if (!ue || !u || !a) {
    return;
  }
  setImmediate(() => {
    void (async () => {
      try {
        let modelConfig = rowToModelConfig(getBestEnabledModelConfig());
        if (!modelConfig) {
          modelConfig = fallbackEnvModelConfig();
        }
        if (!modelConfig) {
          logger.warn('[chat/todos-auto] skipped: no model config');
          return;
        }
        const combined = `【用户】\n${u}\n\n【助手】\n${a}`;
        const items = await extractTodosJson(modelConfig, combined, new Date().toISOString());
        const inserted = insertTodoRows(db, ue, items, 'chat');
        logger.info('[chat/todos-auto]', { userEmail: ue, sessionId: sessionId || null, inserted });
      } catch (error) {
        logger.error('[chat/todos-auto]', error);
      }
    })();
  });
}

function incrementUserConversationTurns(userEmail) {
  const ue = String(userEmail || '').trim();
  if (!ue) {
    return;
  }
  db.prepare(
    `
    INSERT INTO user_stats (user_email, conversation_turns, papers_processed)
    VALUES (?, 1, 0)
    ON CONFLICT(user_email) DO UPDATE SET
      conversation_turns = user_stats.conversation_turns + 1
  `,
  ).run(ue);
}

function incrementUserPapersProcessed(userEmail) {
  const ue = String(userEmail || '').trim();
  if (!ue) {
    return;
  }
  db.prepare(
    `
    INSERT INTO user_stats (user_email, conversation_turns, papers_processed)
    VALUES (?, 0, 1)
    ON CONFLICT(user_email) DO UPDATE SET
      papers_processed = user_stats.papers_processed + 1
  `,
  ).run(ue);
}

/** 自动画像更新频率控制 */
function ensureUsersProfileAutoColumns() {
  try {
    const cols = db.prepare('PRAGMA table_info(users)').all();
    const names = new Set(cols.map((c) => c.name));
    if (!names.has('last_profile_auto_at')) {
      db.exec('ALTER TABLE users ADD COLUMN last_profile_auto_at TEXT');
    }
    if (!names.has('profile_auto_day')) {
      db.exec('ALTER TABLE users ADD COLUMN profile_auto_day TEXT');
    }
    if (!names.has('profile_auto_runs')) {
      db.exec('ALTER TABLE users ADD COLUMN profile_auto_runs INTEGER DEFAULT 0');
    }
  } catch (e) {
    logger.error('[db migrate] users profile_auto columns', e);
    throw e;
  }
}
ensureUsersProfileAutoColumns();

const PROFILE_AUTO_MIN_INTERVAL_MS = 5 * 60 * 60 * 1000;
const PROFILE_AUTO_MAX_RUNS_PER_DAY = 8;
const PROFILE_TURNS_MIN = 3;
const PROFILE_TURNS_MAX_BUFFER = 5;

const profileAnalysisInFlight = new Set();

function parseTurnsJson(raw) {
  try {
    const a = JSON.parse(String(raw || '[]'));
    return Array.isArray(a) ? a : [];
  } catch {
    return [];
  }
}

function appendProfileTurn(userEmail, userMsg, assistantMsg, chatMode) {
  const ue = String(userEmail || '').trim();
  if (!ue || String(chatMode || '').toLowerCase() === 'companion') {
    return;
  }
  const u = String(userMsg || '').trim().slice(0, 12000);
  const a = String(assistantMsg || '').trim().slice(0, 16000);
  if (!u || !a) {
    return;
  }

  const row = db.prepare('SELECT turns_json FROM profile_turn_buffer WHERE user_email = ?').get(ue);
  let turns = row ? parseTurnsJson(row.turns_json) : [];
  turns.push({ u, a });
  if (turns.length > PROFILE_TURNS_MAX_BUFFER) {
    turns = turns.slice(-PROFILE_TURNS_MAX_BUFFER);
  }

  db.prepare(`
    INSERT INTO profile_turn_buffer (user_email, turns_json, updated_at)
    VALUES (?, ?, datetime('now'))
    ON CONFLICT(user_email) DO UPDATE SET turns_json = excluded.turns_json, updated_at = excluded.updated_at
  `).run(ue, JSON.stringify(turns));

  tryScheduleProfileAnalysis(ue);
}

function tryScheduleProfileAnalysis(userEmail) {
  const ue = String(userEmail || '').trim();
  if (!ue) {
    return;
  }

  const rowBuf = db.prepare('SELECT turns_json FROM profile_turn_buffer WHERE user_email = ?').get(ue);
  const turns = rowBuf ? parseTurnsJson(rowBuf.turns_json) : [];
  if (turns.length < PROFILE_TURNS_MIN) {
    return;
  }

  const userRow = db
    .prepare(
      'SELECT email, last_profile_auto_at, profile_auto_day, profile_auto_runs FROM users WHERE email = ?',
    )
    .get(ue);
  if (!userRow) {
    return;
  }

  const now = Date.now();
  const lastAt = userRow.last_profile_auto_at ? Date.parse(userRow.last_profile_auto_at) : 0;
  if (lastAt && now - lastAt < PROFILE_AUTO_MIN_INTERVAL_MS) {
    return;
  }

  const dayKey = new Date().toISOString().slice(0, 10);
  let runsToday = Number(userRow.profile_auto_runs || 0);
  if (userRow.profile_auto_day !== dayKey) {
    runsToday = 0;
  }
  if (runsToday >= PROFILE_AUTO_MAX_RUNS_PER_DAY) {
    return;
  }

  if (profileAnalysisInFlight.has(ue)) {
    return;
  }
  profileAnalysisInFlight.add(ue);

  setImmediate(() => {
    void (async () => {
      try {
        const modelRow = getBestEnabledModelConfig();
        const modelConfig = rowToModelConfig(modelRow);
        if (!modelConfig) {
          return;
        }

        const bufRow = db.prepare('SELECT turns_json FROM profile_turn_buffer WHERE user_email = ?').get(ue);
        const bufTurns = bufRow ? parseTurnsJson(bufRow.turns_json) : [];
        if (bufTurns.length < PROFILE_TURNS_MIN) {
          return;
        }

        const transcript = bufTurns
          .map((t) => `用户：${String(t.u || '').trim()}\n助手：${String(t.a || '').trim()}`)
          .join('\n---\n')
          .slice(0, 16000);

        const urow = db.prepare('SELECT user_profile FROM users WHERE email = ?').get(ue);
        const existing = parseStoredUserProfile(urow?.user_profile);
        const patch = await analyzeUserProfileFromTranscript(modelConfig, transcript, existing);
        const merged = mergeIncrementalProfilePatch(existing, patch);
        merged.emotion_history = appendEmotionHistoryPoint(merged.emotion_history, patch.stress, patch.mood);
        const json = JSON.stringify(merged);

        const nextRuns = runsToday + 1;
        db.prepare(
          `UPDATE users SET user_profile = ?, last_profile_auto_at = datetime('now'), profile_auto_day = ?, profile_auto_runs = ? WHERE email = ?`,
        ).run(json, dayKey, nextRuns, ue);
        db.prepare('DELETE FROM profile_turn_buffer WHERE user_email = ?').run(ue);
      } catch (e) {
        console.warn('[profile-auto]', ue, e?.message || e);
      } finally {
        profileAnalysisInFlight.delete(ue);
      }
    })();
  });
}

/** 合并/校验画像字段时用（空串基准，便于 PATCH） */
function defaultUserProfile() {
  return {
    role: '',
    personality: '',
    preferences: '',
    current_status: '',
    emotion_history: [],
    companion_persona: 'default',
    companion_address: '',
    companion_custom_traits: '',
    chat_active_persona_id: 'workplace_helper',
    chat_custom_personas: [],
  };
}

/** @param {unknown} raw */
function sanitizeChatCustomPersonas(raw) {
  if (!Array.isArray(raw)) {
    return [];
  }
  const out = [];
  for (const item of raw.slice(0, 40)) {
    if (!item || typeof item !== 'object') {
      continue;
    }
    const id = String(/** @type {{ id?: unknown }} */ (item).id || '')
      .trim()
      .slice(0, 80);
    if (!id.startsWith('custom_')) {
      continue;
    }
    const name = String(/** @type {{ name?: unknown }} */ (item).name || '自定义伙伴')
      .trim()
      .slice(0, 80);
    const emoji = String(/** @type {{ emoji?: unknown }} */ (item).emoji || '')
      .trim()
      .slice(0, 8);
    const systemPrompt = String(/** @type {{ systemPrompt?: unknown }} */ (item).systemPrompt || '')
      .trim()
      .slice(0, 2500);
    if (!systemPrompt) {
      continue;
    }
    out.push({ id, name, emoji, systemPrompt });
  }
  return out.slice(0, 24);
}

/** 尚无用户记录或接口兜底展示用（避免前端误以为路由 404） */
function placeholderPublicProfile() {
  return {
    role: '普通用户',
    personality: '待观察',
    preferences: '',
    current_status: '',
    emotion_history: [],
  };
}

function parseStoredUserProfile(raw) {
  const d = defaultUserProfile();
  if (raw == null || raw === '') {
    return { ...d };
  }
  if (typeof raw !== 'string') {
    return { ...d };
  }
  try {
    const o = JSON.parse(raw);
    if (!o || typeof o !== 'object') {
      return { ...d };
    }
    return {
      role: String(o.role ?? d.role).slice(0, 500),
      personality: String(o.personality ?? d.personality).slice(0, 500),
      preferences: String(o.preferences ?? d.preferences).slice(0, 1000),
      current_status: String(o.current_status ?? d.current_status).slice(0, 1000),
      emotion_history: sanitizeEmotionHistoryArray(o.emotion_history ?? d.emotion_history),
      companion_persona: String(o.companion_persona ?? d.companion_persona ?? 'default').slice(0, 64),
      companion_address: String(o.companion_address ?? d.companion_address).slice(0, 64),
      companion_custom_traits: String(o.companion_custom_traits ?? d.companion_custom_traits).slice(0, 1200),
      chat_active_persona_id: String(o.chat_active_persona_id ?? d.chat_active_persona_id ?? 'workplace_helper').slice(
        0,
        64,
      ),
      chat_custom_personas:
        o.chat_custom_personas !== undefined
          ? sanitizeChatCustomPersonas(o.chat_custom_personas)
          : Array.isArray(d.chat_custom_personas)
            ? sanitizeChatCustomPersonas(d.chat_custom_personas)
            : [],
    };
  } catch {
    return { ...d };
  }
}

function runDbTask(fn) {
  return new Promise((resolve, reject) => {
    setImmediate(() => {
      try {
        resolve(fn());
      } catch (error) {
        reject(error);
      }
    });
  });
}

function maskApiKeySecret(key) {
  const s = String(key || '');
  if (!s) {
    return '';
  }
  if (s.length <= 6) {
    return '******';
  }
  return `${s.slice(0, 4)}…${s.slice(-3)}`;
}

function rowToPublicModelConfig(row) {
  if (!row) {
    return null;
  }
  return {
    id: row.id,
    name: row.name,
    base_url: row.base_url,
    model_name: normalizeModelId(row.model_name),
    priority: row.priority,
    is_enabled: Boolean(row.is_enabled),
    type: row.type,
    created_at: row.created_at,
    api_key_masked: maskApiKeySecret(row.api_key),
    has_api_key: Boolean(String(row.api_key || '').trim()),
  };
}

function rowToModelConfig(row) {
  if (!row || !String(row.api_key || '').trim()) {
    return null;
  }
  return {
    type: row.type,
    base_url: row.base_url,
    model_name: normalizeModelId(row.model_name),
    api_key: String(row.api_key).trim(),
  };
}

function getBestEnabledModelConfig() {
  return db
    .prepare(
      `SELECT * FROM model_configs
       WHERE is_enabled = 1 AND length(trim(api_key)) > 0
       ORDER BY priority DESC, id ASC
       LIMIT 1`,
    )
    .get();
}

function fallbackEnvModelConfig() {
  const apiKey = (process.env.BAILIAN_API_KEY || process.env.DEEPSEEK_API_KEY || '').trim();
  if (!apiKey) {
    return null;
  }
  return {
    type: 'aliyun',
    base_url: process.env.BAILIAN_BASE_URL || 'https://dashscope.aliyuncs.com/compatible-mode/v1',
    model_name: normalizeModelId(process.env.BAILIAN_MODEL || 'qwen-plus') || 'qwen-plus',
    api_key: apiKey,
  };
}

function resolveTemplateModelConfig() {
  const row = getBestEnabledModelConfig();
  const fromDb = rowToModelConfig(row);
  if (fromDb && (fromDb.type === 'aliyun' || fromDb.type === 'openai')) {
    return fromDb;
  }
  return fallbackEnvModelConfig();
}

/** 文献摘要与模板填充共用同一套模型配置 */
function resolvePaperModelConfig() {
  return resolveTemplateModelConfig();
}

const PAPER_PARSE_RETURN_MAX_CHARS = 220000;

function isSafePaperFetchUrl(urlStr) {
  try {
    const u = new URL(String(urlStr).trim());
    if (!['http:', 'https:'].includes(u.protocol)) {
      return false;
    }
    const host = u.hostname.toLowerCase();
    if (['localhost', '127.0.0.1', '0.0.0.0', '[::1]', '::1'].includes(host)) {
      return false;
    }
    return true;
  } catch {
    return false;
  }
}

/** 文献网页拉取：常见桌面 Chrome UA，降低源站与 Jina Reader 网关拦截概率 */
const PAPER_FETCH_BROWSER_UA =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36';

const PAPER_FETCH_JINA_TIMEOUT_MS = 10000;
const PAPER_FETCH_JINA_RETRIES = 3;

function delayPaperFetch(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/** Jina Reader 请求头：桌面 UA + 语言，降低网关与源站拦截概率 */
function buildJinaReaderFetchHeaders() {
  return {
    Accept: 'text/plain,text/markdown;q=0.9,*/*;q=0.8',
    'Accept-Language': 'zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7',
    'User-Agent': PAPER_FETCH_BROWSER_UA,
  };
}

function htmlToPlainTextForPaper(html) {
  return String(html)
    .replace(/<script\b[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style\b[\s\S]*?<\/style>/gi, ' ')
    .replace(/<\/(p|div|br|h[1-6]|li|tr)>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

/** multer diskStorage：统一解析为绝对路径，兼容 Windows 反斜杠与仅有 destination+filename 的情况 */
function getMulterAbsPath(file) {
  if (!file) {
    return '';
  }
  const raw =
    typeof file.path === 'string' && file.path.trim().length > 0
      ? file.path.trim()
      : file.destination && file.filename
        ? path.join(String(file.destination), String(file.filename))
        : '';
  if (!raw) {
    return '';
  }
  return path.resolve(path.normalize(raw));
}

/**
 * pdf-parse 包根 index.js 在 ESM 动态 import 下 `module.parent` 为 undefined，
 * 会误判为“调试模式”并同步读取 cwd 下 ./test/data/05-versions-space.pdf（幽灵路径）。
 * 必须只加载 lib 入口，且仅传入 Buffer。
 */
function seedModelConfigsIfEmpty() {
  const { c } = db.prepare('SELECT COUNT(*) AS c FROM model_configs').get();
  if (c > 0) {
    return;
  }
  const defaultBase = 'https://dashscope.aliyuncs.com/compatible-mode/v1';
  /** 答辩展示：默认模型名显式为豆包系列（实际调用仍走 OpenAI 兼容协议，需在控制台配置对应 Base URL 与密钥） */
  const defaultModel = 'Doubao-2.0-Pro';
  const envKey = (process.env.DEEPSEEK_API_KEY || process.env.BAILIAN_API_KEY || '').trim();
  if (envKey) {
    db.prepare(
      `INSERT INTO model_configs (name, api_key, base_url, model_name, priority, is_enabled, type)
       VALUES (?, ?, ?, ?, ?, ?, ?)`,
    ).run('豆包 · 从环境变量导入', envKey, defaultBase, defaultModel, 100, 1, 'aliyun');
    return;
  }
  db.prepare(
    `INSERT INTO model_configs (name, api_key, base_url, model_name, priority, is_enabled, type)
     VALUES (?, ?, ?, ?, ?, ?, ?)`,
  ).run('豆包 · 默认接入（请填写 API Key）', '', defaultBase, defaultModel, 0, 0, 'aliyun');
}

seedModelConfigsIfEmpty();

/**
 * Multer 在部分环境下把 UTF-8 文件名误作 latin1，导致中文乱码。
 * 仅在「latin1→utf8 后出现汉字且原文件名不含中日韩等字符」时采用修复结果，避免破坏已为 UTF-8 的名称。
 */
function normalizeMulterOriginalName(name) {
  if (name == null || name === '') {
    return name;
  }
  const s = String(name);
  if (/[\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af\u3000-\u303f\uff00-\uffef]/.test(s)) {
    return s;
  }
  try {
    const fixed = Buffer.from(s, 'latin1').toString('utf8');
    if (fixed && fixed !== s && /[\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]/.test(fixed)) {
      return fixed;
    }
  } catch {
    /* ignore */
  }
  return s;
}

const storage = multer.diskStorage({
  destination: (_req, _file, cb) => {
    cb(null, uploadsDir);
  },
  filename: (_req, file, cb) => {
    file.originalname = normalizeMulterOriginalName(file.originalname);
    const originalName = path.basename(file.originalname || 'template.docx');
    const safeName = `${Date.now()}-${crypto.randomUUID()}-${originalName}`;
    cb(null, safeName);
  },
});

const upload = multer({
  storage,
  fileFilter: (_req, file, cb) => {
    file.originalname = normalizeMulterOriginalName(file.originalname);
    const ext = path.extname(file.originalname || '').toLowerCase();
    if (ext === '.docx') {
      cb(null, true);
      return;
    }
    cb(new Error('仅支持 .docx 文件'));
  },
  limits: { fileSize: 20 * 1024 * 1024 },
});

const paperUpload = multer({
  storage,
  fileFilter: (_req, file, cb) => {
    file.originalname = normalizeMulterOriginalName(file.originalname);
    const ext = path.extname(file.originalname || '').toLowerCase();
    if (ext === '.pdf' || ext === '.docx' || ext === '.doc') {
      cb(null, true);
      return;
    }
    cb(new Error('仅支持 .pdf、.doc、.docx'));
  },
  limits: { fileSize: 25 * 1024 * 1024 },
});


const libraryUpload = multer({
  storage,
  fileFilter: (_req, file, cb) => {
    file.originalname = normalizeMulterOriginalName(file.originalname);
    const ext = path.extname(file.originalname || '').toLowerCase();
    if (['.pdf', '.doc', '.docx', '.txt', '.md'].includes(ext)) {
      cb(null, true);
      return;
    }
    cb(new Error('资料库仅支持 .pdf、.doc、.docx、.txt、.md'));
  },
  limits: { fileSize: 30 * 1024 * 1024 },
});

const meetingAudioUploadDir = path.join(uploadsDir, 'meeting-audio');
fs.mkdirSync(meetingAudioUploadDir, { recursive: true });
const meetingFileStorage = multer.diskStorage({
  destination: (_req, _file, cb) => {
    cb(null, meetingAudioUploadDir);
  },
  filename: (_req, file, cb) => {
    file.originalname = normalizeMulterOriginalName(file.originalname);
    const originalName = path.basename(file.originalname || 'recording.mp3');
    cb(null, `${Date.now()}-${crypto.randomUUID()}-${originalName}`);
  },
});
const meetingFileUpload = multer({
  storage: meetingFileStorage,
  fileFilter: (_req, file, cb) => {
    file.originalname = normalizeMulterOriginalName(file.originalname);
    const ext = path.extname(file.originalname || '').toLowerCase();
    if (['.mp3', '.wav', '.m4a'].includes(ext)) {
      cb(null, true);
      return;
    }
    cb(new Error('仅支持 .mp3、.wav、.m4a'));
  },
  limits: { fileSize: 52 * 1024 * 1024 },
});

const meetingFileMulterGate = (req, res, next) => {
  meetingFileUpload.single('file')(req, res, (err) => {
    if (err) {
      return res.status(400).json({ success: false, message: String(err.message || err || '上传失败') });
    }
    next();
  });
};

app.post('/api/auth/register', (req, res) => {
  const account = String(req.body.account || req.body.email || '').trim();
  const password = String(req.body.password || '');
  const name = String(req.body.name || '').trim();
  if (!account || !password || !name) {
    return res.status(400).json({ success: false, message: '请填写账号、密码与用户名' });
  }
  try {
    /** users.email 列作为通用「账号」存储，兼容历史邮箱数据 */
    const stmt = db.prepare('INSERT INTO users (email, password, name) VALUES (?, ?, ?)');
    const info = stmt.run(account, password, name);
    res.json({ success: true, message: '注册成功', userId: info.lastInsertRowid });
  } catch (error) {
    const msg = String(error?.message || '');
    if (msg.includes('UNIQUE constraint')) {
      res.status(400).json({ success: false, message: '该账号已被注册' });
    } else {
      res.status(500).json({ success: false, message: '注册失败，请稍后重试' });
    }
  }
});

app.post('/api/auth/login', (req, res) => {
  const account = String(req.body.account || req.body.email || '').trim();
  const password = String(req.body.password || '');
  if (!account || !password) {
    return res.status(400).json({ success: false, message: '请填写账号与密码' });
  }
  const row = db.prepare('SELECT * FROM users WHERE email = ?').get(account);
  if (!row) {
    return res.status(404).json({ success: false, error: 'ACCOUNT_NOT_FOUND', message: '该账号未注册' });
  }
  if (String(row.password || '') !== password) {
    return res.status(401).json({ success: false, message: '账号或密码错误' });
  }
  res.json({
    success: true,
    message: '登录成功',
    user: { id: row.id, name: row.name, email: row.email },
  });
});

app.get('/api/user/profile', (req, res) => {
  try {
    const email = String(req.query.email || req.query.userEmail || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '缺少查询参数 email' });
    }
    const user = db.prepare('SELECT id, email, name, user_profile FROM users WHERE email = ?').get(email);
    if (!user) {
      const profile = placeholderPublicProfile();
      return res.json({
        success: true,
        profile,
        email,
        name: '',
        persisted: false,
        placeholder: true,
      });
    }
    const profile = parseStoredUserProfile(user.user_profile);
    res.json({
      success: true,
      profile,
      email: user.email,
      name: user.name,
      persisted: true,
      placeholder: false,
    });
  } catch (error) {
    logger.error('[user/profile GET]', error);
    res.status(500).json({ success: false, message: String(error?.message || error || '读取失败') });
  }
});

/** 工作台概览：真实统计 */
app.get('/api/stats', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const st = db.prepare('SELECT conversation_turns, papers_processed FROM user_stats WHERE user_email = ?').get(email);
    const sc = db.prepare('SELECT COUNT(*) AS c FROM skills WHERE user_email = ?').get(email);
    res.json({
      success: true,
      conversation_turns: st ? Number(st.conversation_turns || 0) : 0,
      papers_processed: st ? Number(st.papers_processed || 0) : 0,
      skills_count: sc ? Number(sc.c || 0) : 0,
    });
  } catch (error) {
    logger.error('[stats GET]', error);
    res.status(500).json({ success: false, message: '读取统计失败' });
  }
});

/** 记录一次评分（如点赞 score=1） */
app.post('/api/ratings', (req, res) => {
  try {
    const score = Number(req.body?.score);
    if (!Number.isFinite(score) || Math.abs(Math.trunc(score)) > 10) {
      return res.status(400).json({ success: false, message: '需要有效 score（整数，如 1 表示点赞）' });
    }
    const s = Math.trunc(score);
    const info = db.prepare('INSERT INTO ai_ratings (score) VALUES (?)').run(s);
    res.json({ success: true, id: Number(info.lastInsertRowid) });
  } catch (error) {
    logger.error('[ratings POST]', error);
    res.status(500).json({ success: false, message: '保存评分失败' });
  }
});

function formatRatingTrendDateKey(d) {
  const parts = String(d || '').split('-');
  if (parts.length >= 3) {
    return `${parts[1]}-${parts[2]}`;
  }
  return String(d || '');
}

/** 按日汇总评分条数，供折线图 */
app.get('/api/ratings/trend', (req, res) => {
  try {
    const days = Math.min(365, Math.max(7, Number.parseInt(String(req.query.days || '90'), 10) || 90));
    const rows = db
      .prepare(
        `SELECT date(created_at) AS d, COUNT(*) AS c
         FROM ai_ratings
         WHERE datetime(created_at) >= datetime('now', ?)
         GROUP BY date(created_at)
         ORDER BY d ASC`,
      )
      .all(`-${days} days`);
    const items = (rows || []).map((r) => ({
      date: formatRatingTrendDateKey(r.d),
      count: Number(r.c || 0),
    }));
    res.json({ success: true, items, days });
  } catch (error) {
    logger.error('[ratings/trend GET]', error);
    res.status(500).json({ success: false, message: '读取趋势失败' });
  }
});

app.put('/api/user/profile', (req, res) => {
  try {
    const { email, profile } = req.body || {};
    const em = String(email || '').trim();
    if (!em) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const p = profile && typeof profile === 'object' && !Array.isArray(profile) ? profile : {};
    const row = db.prepare('SELECT id, user_profile FROM users WHERE email = ?').get(em);
    if (!row) {
      const prior = defaultUserProfile();
      const merged = {
        role: String(p.role ?? prior.role).slice(0, 500),
        personality: String(p.personality ?? prior.personality).slice(0, 500),
        preferences: String(p.preferences ?? prior.preferences).slice(0, 1000),
        current_status: String(p.current_status ?? prior.current_status).slice(0, 1000),
        emotion_history: Object.prototype.hasOwnProperty.call(p, 'emotion_history') && Array.isArray(p.emotion_history)
          ? sanitizeEmotionHistoryArray(p.emotion_history)
          : prior.emotion_history,
        companion_persona: String(
          p.companion_persona != null ? p.companion_persona : prior.companion_persona || 'default',
        ).slice(0, 64),
        companion_address: String(p.companion_address != null ? p.companion_address : prior.companion_address).slice(0, 64),
        companion_custom_traits: String(
          p.companion_custom_traits != null ? p.companion_custom_traits : prior.companion_custom_traits,
        ).slice(0, 1200),
        chat_active_persona_id: String(
          p.chat_active_persona_id != null ? p.chat_active_persona_id : prior.chat_active_persona_id || 'workplace_helper',
        ).slice(0, 64),
        chat_custom_personas:
          p.chat_custom_personas !== undefined
            ? sanitizeChatCustomPersonas(p.chat_custom_personas)
            : prior.chat_custom_personas || [],
      };
      return res.status(200).json({
        success: false,
        code: 'USER_NOT_FOUND',
        message: '数据库中尚无该邮箱账号，请先注册登录后再保存画像',
        profile: merged,
      });
    }
    const prior = parseStoredUserProfile(row.user_profile);
    const merged = {
      role: String(p.role ?? prior.role).slice(0, 500),
      personality: String(p.personality ?? prior.personality).slice(0, 500),
      preferences: String(p.preferences ?? prior.preferences).slice(0, 1000),
      current_status: String(p.current_status ?? prior.current_status).slice(0, 1000),
      emotion_history:
        Object.prototype.hasOwnProperty.call(p, 'emotion_history') && Array.isArray(p.emotion_history)
          ? sanitizeEmotionHistoryArray(p.emotion_history)
          : prior.emotion_history,
      companion_persona: String(
        p.companion_persona != null ? p.companion_persona : prior.companion_persona || 'default',
      ).slice(0, 64),
      companion_address: String(p.companion_address != null ? p.companion_address : prior.companion_address).slice(0, 64),
      companion_custom_traits: String(
        p.companion_custom_traits != null ? p.companion_custom_traits : prior.companion_custom_traits,
      ).slice(0, 1200),
      chat_active_persona_id: String(
        p.chat_active_persona_id != null ? p.chat_active_persona_id : prior.chat_active_persona_id || 'workplace_helper',
      ).slice(0, 64),
      chat_custom_personas:
        p.chat_custom_personas !== undefined
          ? sanitizeChatCustomPersonas(p.chat_custom_personas)
          : prior.chat_custom_personas || [],
    };
    const json = JSON.stringify(merged);
    db.prepare('UPDATE users SET user_profile = ? WHERE id = ?').run(json, row.id);
    res.json({ success: true, profile: merged });
  } catch (error) {
    logger.error('[user/profile PUT]', error);
    res.status(500).json({ success: false, message: String(error?.message || error || '保存失败') });
  }
});

/** 按需：结合画像与情绪历史生成「本周成长与情绪回溯」短文，并写入历史表 */
app.post('/api/user/weekly-review', async (req, res) => {
  try {
    const email = String(req.body?.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const row = db.prepare('SELECT user_profile FROM users WHERE email = ?').get(email);
    if (!row) {
      return res.status(404).json({ success: false, message: '用户不存在' });
    }
    const profile = parseStoredUserProfile(row.user_profile);
    const contextBlock = JSON.stringify(
      {
        role: profile.role,
        personality: profile.personality,
        preferences: profile.preferences,
        current_status: profile.current_status,
        emotion_history: profile.emotion_history || [],
      },
      null,
      2,
    );

    const modelRow = getBestEnabledModelConfig();
    const modelConfig = rowToModelConfig(modelRow);
    if (!modelConfig) {
      return res.status(503).json({
        success: false,
        message: '没有可用的已启用模型配置，无法生成周回顾',
      });
    }
    const report = await generateWeeklyGrowthReview(modelConfig, contextBlock);
    const now = new Date();
    const weekLabel = `${now.getFullYear()}-W${String(getIsoWeekNumber(now)).padStart(2, '0')}`;
    const ins = db
      .prepare(
        'INSERT INTO weekly_review_reports (user_email, week_label, report_body, created_at) VALUES (?, ?, ?, datetime(\'now\'))',
      )
      .run(email, weekLabel, String(report || '').slice(0, 120000));
    res.json({
      success: true,
      report,
      id: Number(ins.lastInsertRowid),
      week_label: weekLabel,
      created_at: new Date().toISOString(),
    });
  } catch (error) {
    logger.error('[weekly-review]', error);
    res.status(500).json({ success: false, message: String(error?.message || error || '生成失败') });
  }
});

/** 列出当前用户的周回顾历史（新在前） */
app.get('/api/user/weekly-reports', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const limit = Math.min(100, Math.max(1, Number.parseInt(String(req.query.limit || '50'), 10) || 50));
    const rows = db
      .prepare(
        'SELECT id, week_label, report_body, created_at FROM weekly_review_reports WHERE user_email = ? ORDER BY datetime(created_at) DESC LIMIT ?',
      )
      .all(email, limit);
    res.json({ success: true, items: rows });
  } catch (error) {
    logger.error('[weekly-reports GET]', error);
    res.status(500).json({ success: false, message: '读取历史失败' });
  }
});

function getIsoWeekNumber(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  return Math.ceil(((date - yearStart) / 86400000 + 1) / 7);
}

/** 演示专用：重置 user_profile 与画像缓冲，不删除 Skill 库 */
app.post('/api/user/demo-reset', (req, res) => {
  try {
    const email = String(req.body?.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const row = db.prepare('SELECT id FROM users WHERE email = ?').get(email);
    if (!row) {
      return res.status(404).json({ success: false, message: '用户不存在' });
    }
    const resetJson = JSON.stringify(defaultUserProfile());
    db.prepare('UPDATE users SET user_profile = ? WHERE id = ?').run(resetJson, row.id);
    db.prepare('DELETE FROM profile_turn_buffer WHERE user_email = ?').run(email);
    res.json({ success: true, profile: parseStoredUserProfile(resetJson) });
  } catch (error) {
    logger.error('[demo-reset]', error);
    res.status(500).json({ success: false, message: String(error?.message || error || '重置失败') });
  }
});

/** 本轮对话结束后由前端上报，用于缓冲 3～5 轮后异步更新画像（限频在 tryScheduleProfileAnalysis） */
app.post('/api/user/profile/record-turn', (req, res) => {
  try {
    const { userEmail, userMessage, assistantMessage, chatMode } = req.body || {};
    const ue = String(userEmail || '').trim();
    if (ue) {
      incrementUserConversationTurns(ue);
    }
    appendProfileTurn(userEmail, userMessage, assistantMessage, chatMode);
    res.json({ success: true, hintPulse: true });
  } catch (error) {
    logger.error('[profile record-turn]', error);
    res.status(500).json({ success: false, message: String(error?.message || error || '记录失败') });
  }
});

const MAX_SKILL_TRANSCRIPT_CHARS = 28000;
const MAX_SKILL_PROMPT_STORE = 16000;

function parseSkillJsonFromModelText(text) {
  const raw = String(text || '').trim();
  const fence = raw.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const inner = fence ? fence[1].trim() : raw;
  try {
    const o = JSON.parse(inner);
    if (o && typeof o === 'object' && !Array.isArray(o)) {
      return o;
    }
  } catch {
    /* try brace */
  }
  const brace = raw.match(/\{[\s\S]*\}/);
  if (brace) {
    try {
      const o = JSON.parse(brace[0]);
      if (o && typeof o === 'object' && !Array.isArray(o)) {
        return o;
      }
    } catch {
      /* ignore */
    }
  }
  throw new Error('JSON 解析失败');
}

app.get('/api/skills', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    const detail = String(req.query.detail || '') === '1';
    if (!email) {
      return res.status(400).json({ success: false, message: '缺少查询参数 email' });
    }
    const sql = detail
      ? 'SELECT id, name, description, emoji, wake_word, use_count, system_prompt, created_at FROM skills WHERE user_email = ? ORDER BY id DESC LIMIT 200'
      : 'SELECT id, name, description, emoji, wake_word, use_count, created_at FROM skills WHERE user_email = ? ORDER BY id DESC LIMIT 200';
    const rows = db.prepare(sql).all(email);
    res.json({ success: true, items: rows });
  } catch (error) {
    logger.error('[skills GET]', error);
    res.status(500).json({ success: false, message: '读取 Skill 列表失败' });
  }
});

/** 手动新增 Skill */
app.post('/api/skills', (req, res) => {
  try {
    const { email, name, description, system_prompt, emoji, wake_word } = req.body || {};
    const em = String(email || '').trim();
    const n = String(name || '').trim().slice(0, 120);
    if (!em || !n) {
      return res.status(400).json({ success: false, message: '需要 email 与名称' });
    }
    const desc = String(description ?? '').trim().slice(0, 500);
    const sp = String(system_prompt ?? '').trim().slice(0, MAX_SKILL_PROMPT_STORE);
    if (!sp) {
      return res.status(400).json({ success: false, message: '请填写 System Prompt' });
    }
    const info = db
      .prepare(
        'INSERT INTO skills (name, description, system_prompt, user_email, emoji, wake_word) VALUES (?, ?, ?, ?, ?, ?)',
      )
      .run(
        n,
        desc,
        sp,
        em,
        String(emoji || '✨')
          .trim()
          .slice(0, 8) || '✨',
        String(wake_word ?? '')
          .trim()
          .slice(0, 64),
      );
    res.json({ success: true, id: Number(info.lastInsertRowid) });
  } catch (error) {
    logger.error('[skills POST]', error);
    res.status(500).json({ success: false, message: String(error?.message || error || '创建失败') });
  }
});

app.put('/api/skills/:id', (req, res) => {
  try {
    const id = Number.parseInt(String(req.params.id), 10);
    const { email, name, description, system_prompt, emoji, wake_word } = req.body || {};
    const em = String(email || '').trim();
    if (!em || !Number.isFinite(id) || id < 1) {
      return res.status(400).json({ success: false, message: '参数无效' });
    }
    const existing = db.prepare('SELECT * FROM skills WHERE id = ? AND user_email = ?').get(id, em);
    if (!existing) {
      return res.status(404).json({ success: false, message: 'Skill 不存在' });
    }
    const nextName = typeof name === 'string' ? String(name).trim().slice(0, 120) : existing.name;
    const nextDesc =
      typeof description === 'string' ? String(description).trim().slice(0, 500) : existing.description;
    const nextPrompt =
      typeof system_prompt === 'string'
        ? String(system_prompt).trim().slice(0, MAX_SKILL_PROMPT_STORE)
        : existing.system_prompt;
    const nextEmoji = typeof emoji === 'string' ? String(emoji).trim().slice(0, 8) : existing.emoji || '✨';
    const nextWake =
      typeof wake_word === 'string' ? String(wake_word).trim().slice(0, 64) : existing.wake_word || '';

    db.prepare(
      'UPDATE skills SET name = ?, description = ?, system_prompt = ?, emoji = ?, wake_word = ? WHERE id = ? AND user_email = ?',
    ).run(nextName, nextDesc, nextPrompt, nextEmoji || '✨', nextWake, id, em);

    res.json({
      success: true,
      skill: {
        id,
        name: nextName,
        description: nextDesc,
        emoji: nextEmoji,
        wake_word: nextWake,
      },
    });
  } catch (error) {
    logger.error('[skills PUT]', error);
    res.status(500).json({ success: false, message: String(error?.message || error || '更新失败') });
  }
});

app.delete('/api/skills/:id', (req, res) => {
  try {
    const id = Number.parseInt(String(req.params.id), 10);
    const email = String(req.query.email || req.body?.email || '').trim();
    if (!email || !Number.isFinite(id)) {
      return res.status(400).json({ success: false, message: '参数无效' });
    }
    const info = db.prepare('DELETE FROM skills WHERE id = ? AND user_email = ?').run(id, email);
    if (info.changes === 0) {
      return res.status(404).json({ success: false, message: '未找到 Skill' });
    }
    res.json({ success: true });
  } catch (error) {
    logger.error('[skills DELETE]', error);
    res.status(500).json({ success: false, message: '删除失败' });
  }
});

app.post('/api/skills/extract', async (req, res) => {
  try {
    const { messages, userEmail } = req.body || {};
    const ue = String(userEmail || '').trim();
    if (!ue) {
      return res.status(400).json({ success: false, message: '请先登录后再提炼 Skill' });
    }
    if (!Array.isArray(messages) || messages.length < 2) {
      return res.status(400).json({ success: false, message: '请提供至少两条对话记录（建议完成一轮以上对话）' });
    }
    const lines = [];
    for (const m of messages.slice(-80)) {
      if (!m || typeof m !== 'object') {
        continue;
      }
      const role = String(m.role || '').toLowerCase();
      const content = String(m.content ?? '').trim();
      if (!content || (role !== 'user' && role !== 'assistant')) {
        continue;
      }
      lines.push(`${role === 'user' ? '用户' : '助手'}：${content}`);
    }
    const transcript = lines.join('\n\n').slice(0, MAX_SKILL_TRANSCRIPT_CHARS);
    if (transcript.length < 24) {
      return res.status(400).json({ success: false, message: '对话内容过短，无法提炼有效 Skill' });
    }

    const modelRow = getBestEnabledModelConfig();
    const modelConfig = rowToModelConfig(modelRow);
    if (!modelConfig) {
      return res.status(503).json({
        success: false,
        message: '没有可用的已启用模型配置，无法提炼 Skill',
      });
    }

    const extractSystem = `你是资深 Prompt 与办公自动化顾问。请阅读对话记录，逆向提炼一套可复用于「同类任务」的系统指令（system_prompt），并给出名称、描述、展示用 Emoji 与唤醒词。

硬性规则（必须遵守）：
1) system_prompt 必须是泛化的流程与风格约束：步骤、角色、输出结构、检查清单均可；禁止写入对话中出现的具体论文题名、作者、精确统计值、文件名、内部项目代号等——一律抽象为占位符，如「[受众]」「[核心指标]」「[参考文献列表]」。
2) name：不超过 24 字，概括任务类型。
3) description：一句话适用场景，不出现私密细节。
4) emoji：单个 Unicode 表情符号，贴合技能气质（如 📊 📝 🧭）。
5) wake_word：用户可在输入框使用的唤醒短语，建议以 / 开头、简短无空格，例如 /周报助手。
6) 仅输出一个 JSON 对象，键严格为 name、description、system_prompt、emoji、wake_word；不要用 markdown 围栏包裹以外的废话。`;

    const userMsg = `请分析以下对话记录，提取标准化系统指令并起名。\n返回 JSON：{ "name", "description", "system_prompt", "emoji", "wake_word" }\n\n对话记录：\n${transcript}`;

    const model = normalizeModelId(modelConfig.model_name) || 'deepseek-v3';
    const body = withThinkingParams(
      {
        model,
        messages: [
          { role: 'system', content: extractSystem },
          { role: 'user', content: userMsg.slice(0, 32000) },
        ],
        temperature: 0.25,
        stream: false,
      },
      modelConfig.model_name,
    );

    let data;
    try {
      data = await postChatCompletion(modelConfig, body);
    } catch (err) {
      logger.error('[skills/extract] upstream', err);
      return res.status(502).json({
        success: false,
        message: `模型调用失败：${String(err?.message || err)}`,
      });
    }

    let replyText = '';
    try {
      replyText = extractTextFromResponse(data);
    } catch {
      replyText = '';
    }
    if (!replyText || !String(replyText).trim()) {
      return res.status(422).json({ success: false, message: '模型未返回有效文本，请稍后重试' });
    }

    let parsed;
    try {
      parsed = parseSkillJsonFromModelText(replyText);
    } catch (e) {
      console.warn('[skills/extract] parse failed snippet:', String(replyText).slice(0, 400));
      return res.status(422).json({
        success: false,
        message: '模型返回格式无法解析为 JSON，请缩短对话记录或稍后重试',
      });
    }

    const name = String(parsed.name || '').trim().slice(0, 120);
    const description = String(parsed.description || '').trim().slice(0, 500);
    let system_prompt = String(parsed.system_prompt || '').trim();
    const emojiRaw = String(parsed.emoji || '✨').trim().slice(0, 8) || '✨';
    let wake_word = String(parsed.wake_word || '').trim().slice(0, 64);
    if (!wake_word) {
      const slug = name.replace(/\s+/g, '').replace(/[^\w\u4e00-\u9fff]/g, '').slice(0, 14);
      wake_word = `/${slug || 'skill'}`;
    }
    if (!name || !system_prompt) {
      return res.status(422).json({ success: false, message: '提炼结果缺少 name 或 system_prompt' });
    }
    system_prompt = system_prompt.slice(0, MAX_SKILL_PROMPT_STORE);

    let info;
    try {
      info = db
        .prepare(
          'INSERT INTO skills (name, description, system_prompt, user_email, emoji, wake_word) VALUES (?, ?, ?, ?, ?, ?)',
        )
        .run(name, description || '', system_prompt, ue, emojiRaw, wake_word);
    } catch (insErr) {
      logger.error('[skills/extract] insert', insErr);
      return res.status(500).json({ success: false, message: '保存 Skill 失败' });
    }

    res.json({
      success: true,
      message: 'Skill 已保存',
      skill: {
        id: Number(info.lastInsertRowid),
        name,
        description,
      },
    });
  } catch (error) {
    logger.error('[skills/extract]', error);
    res.status(500).json({ success: false, message: String(error?.message || error || '服务器异常') });
  }
});

app.get('/api/model-configs', async (_req, res) => {
  try {
    const rows = await runDbTask(() => db.prepare('SELECT * FROM model_configs ORDER BY priority DESC, id ASC').all());
    res.json({ success: true, items: rows.map((r) => rowToPublicModelConfig(r)) });
  } catch (error) {
    logger.error('[model-configs] list failed', error);
    res.status(500).json({ success: false, message: '读取模型配置失败' });
  }
});

app.post('/api/model-configs', async (req, res) => {
  const { name, api_key, base_url, model_name, priority, is_enabled, type } = req.body || {};
  if (!name || !base_url || !model_name || !type) {
    return res.status(400).json({ success: false, message: '缺少必填字段：name / base_url / model_name / type' });
  }
  if (!['openai', 'aliyun', 'openclaw'].includes(type)) {
    return res.status(400).json({ success: false, message: 'type 必须为 openai | aliyun | openclaw' });
  }
  const key = String(api_key || '').trim();
  if (!key) {
    return res.status(400).json({ success: false, message: '新增配置必须提供 API Key' });
  }
  const pri = Number.parseInt(String(priority ?? '0'), 10) || 0;
  const enabled = is_enabled === false || is_enabled === 0 ? 0 : 1;
  try {
    const insertId = await runDbTask(() => {
      const r = db
        .prepare(
          `INSERT INTO model_configs (name, api_key, base_url, model_name, priority, is_enabled, type)
           VALUES (?, ?, ?, ?, ?, ?, ?)`,
        )
        .run(String(name).trim(), key, String(base_url).trim(), normalizeModelId(String(model_name).trim()), pri, enabled, type);
      return Number(r.lastInsertRowid);
    });
    const row = db.prepare('SELECT * FROM model_configs WHERE id = ?').get(insertId);
    res.status(201).json({ success: true, item: rowToPublicModelConfig(row) });
  } catch (error) {
    logger.error('[model-configs] create failed', error);
    res.status(500).json({ success: false, message: '创建失败' });
  }
});

app.put('/api/model-configs/:id', async (req, res) => {
  const id = Number.parseInt(req.params.id, 10);
  if (!Number.isFinite(id)) {
    return res.status(400).json({ success: false, message: '无效的 id' });
  }
  const existing = db.prepare('SELECT * FROM model_configs WHERE id = ?').get(id);
  if (!existing) {
    return res.status(404).json({ success: false, message: '配置不存在' });
  }
  const { name, api_key, base_url, model_name, priority, is_enabled, type } = req.body || {};
  if (!name || !base_url || !model_name || !type) {
    return res.status(400).json({ success: false, message: '缺少必填字段' });
  }
  if (!['openai', 'aliyun', 'openclaw'].includes(type)) {
    return res.status(400).json({ success: false, message: 'type 无效' });
  }
  const keyIn = api_key === undefined || api_key === null ? '' : String(api_key).trim();
  const keepKey = keyIn === '' || keyIn === '__UNCHANGED__';
  const newKey = keepKey ? existing.api_key : keyIn;
  const pri = Number.parseInt(String(priority ?? existing.priority), 10) || 0;
  const enabled = is_enabled === false || is_enabled === 0 ? 0 : 1;
  try {
    await runDbTask(() =>
      db
        .prepare(
          `UPDATE model_configs SET
            name = ?, api_key = ?, base_url = ?, model_name = ?, priority = ?, is_enabled = ?, type = ?
           WHERE id = ?`,
        )
        .run(String(name).trim(), newKey, String(base_url).trim(), normalizeModelId(String(model_name).trim()), pri, enabled, type, id),
    );
    const row = db.prepare('SELECT * FROM model_configs WHERE id = ?').get(id);
    res.json({ success: true, item: rowToPublicModelConfig(row) });
  } catch (error) {
    logger.error('[model-configs] update failed', error);
    res.status(500).json({ success: false, message: '更新失败' });
  }
});

app.delete('/api/model-configs/:id', async (req, res) => {
  const id = Number.parseInt(req.params.id, 10);
  if (!Number.isFinite(id)) {
    return res.status(400).json({ success: false, message: '无效的 id' });
  }
  try {
    const info = await runDbTask(() => db.prepare('DELETE FROM model_configs WHERE id = ?').run(id));
    if (info.changes === 0) {
      return res.status(404).json({ success: false, message: '配置不存在' });
    }
    res.json({ success: true });
  } catch (error) {
    logger.error('[model-configs] delete failed', error);
    res.status(500).json({ success: false, message: '删除失败' });
  }
});

app.post('/api/models/test', async (req, res) => {
  try {
    const body = req.body || {};
    const config = {
      type: body.type,
      base_url: body.base_url,
      model_name: body.model_name,
      api_key: body.api_key,
    };
    if (!['openai', 'aliyun', 'openclaw'].includes(config.type)) {
      return res.status(200).json({ success: false, message: 'type 必须为 openai | aliyun | openclaw' });
    }
    if (!String(config.base_url || '').trim()) {
      return res.status(200).json({ success: false, message: '请填写 Base URL' });
    }
    if (!String(config.api_key || '').trim()) {
      return res.status(200).json({ success: false, message: '请填写 API Key' });
    }
    if (!String(config.model_name || '').trim()) {
      return res.status(200).json({ success: false, message: '请填写模型 ID' });
    }
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 20000);
    try {
      await pingChatModel(config, controller.signal);
      clearTimeout(timer);
      return res.status(200).json({ success: true });
    } catch (error) {
      clearTimeout(timer);
      const msg = String(error?.message || error || '测试失败');
      return res.status(200).json({ success: false, message: msg });
    }
  } catch (error) {
    logger.error('[models/test] fatal', error);
    return res.status(200).json({
      success: false,
      message: String(error?.message || error || '服务异常，请稍后重试'),
    });
  }
});

// —— 文献论文：解析上传 / URL 拉取 / 结构化摘要 ——
app.post(
  '/api/papers/parse-file',
  (req, res, next) => {
    paperUpload.single('file')(req, res, (err) => {
      if (err) {
        return res.status(400).json({ success: false, message: String(err.message || err || '上传被拒绝') });
      }
      next();
    });
  },
  async (req, res) => {
    let diskPath = '';
    try {
      diskPath = getMulterAbsPath(req.file);
      if (!req.file || !diskPath) {
        return res.status(400).json({ success: false, message: '未收到文件' });
      }
      if (!fs.existsSync(diskPath)) {
        return res.status(400).json({ success: false, message: '上传文件未成功保存到服务器' });
      }
      const ext = path.extname(req.file.originalname || '').toLowerCase();
      let text = await extractPaperTextFromPath(diskPath, ext);
      const fullLen = text.length;
      const truncated = text.length > PAPER_PARSE_RETURN_MAX_CHARS;
      if (truncated) {
        text = text.slice(0, PAPER_PARSE_RETURN_MAX_CHARS);
      }
      const ue = String(req.body?.userEmail || '').trim();
      let libraryItemId = null;
      if (ue) {
        incrementUserPapersProcessed(ue);
        libraryItemId = insertKnowledgeItem({
          userEmail: ue,
          title: path.basename(req.file.originalname || '文献资料', ext),
          itemType: 'document',
          sourceKind: 'paper_upload',
          originalName: req.file.originalname || '',
          storedFilename: req.file.filename || '',
          mimeType: req.file.mimetype || '',
          folderPath: '文献资料',
          contentText: text,
          metadata: {
            charCount: fullLen,
            truncated,
            mimeKind: ext === '.pdf' ? 'pdf' : ext === '.doc' ? 'doc' : 'docx',
          },
        });
      }
      return res.json({
        success: true,
        text,
        charCount: fullLen,
        truncated,
        filename: req.file.originalname || '',
        mimeKind: ext === '.pdf' ? 'pdf' : ext === '.doc' ? 'doc' : 'docx',
        libraryItemId,
      });
    } catch (err) {
      logger.error('[papers/parse-file]', err);
      const msg = String(err?.message || err || '解析失败');
      if (!res.headersSent) {
        res.status(500).json({ success: false, message: msg, error: msg });
      }
    } finally {
      try {
        const ue = String(req.body?.userEmail || '').trim();
        if (!ue && diskPath && fs.existsSync(diskPath)) {
          await fs.promises.unlink(diskPath);
        }
      } catch {
        /* ignore */
      }
    }
  },
);

app.post('/api/papers/fetch-url', async (req, res) => {
  const { url } = req.body || {};
  const target = String(url || '').trim();
  if (!target || !isSafePaperFetchUrl(target)) {
    return res.status(400).json({ success: false, message: '请提供有效的 http(s) 公开链接' });
  }
  try {
    /** Jina Reader 约定：完整目标 URL 接在 https://r.jina.ai/ 之后（勿对整个 URL 再做 encodeURIComponent，否则部分站点无法命中） */
    const jinaReader = `https://r.jina.ai/${target}`;
    let text = '';
    for (let attempt = 1; attempt <= PAPER_FETCH_JINA_RETRIES; attempt++) {
      if (attempt > 1) {
        await delayPaperFetch(350 * (attempt - 1));
      }
      try {
        const jr = await fetch(jinaReader, {
          headers: buildJinaReaderFetchHeaders(),
          redirect: 'follow',
          signal: AbortSignal.timeout(PAPER_FETCH_JINA_TIMEOUT_MS),
        });
        if (jr.ok) {
          const body = await jr.text();
          if (body && String(body).trim().length >= 120) {
            text = body;
            break;
          }
          console.warn(
            `[papers/fetch-url] jina reader short body attempt=${attempt} status=${jr.status} len=${String(body || '').length} preview=${String(body || '').slice(0, 240).replace(/\s+/g, ' ')}`,
          );
        } else {
          const errBody = await jr.text().catch(() => '');
          console.warn(
            `[papers/fetch-url] jina reader non-OK attempt=${attempt} status=${jr.status} statusText=${jr.statusText || ''} body=${String(errBody).slice(0, 800)}`,
          );
        }
      } catch (e) {
        const name = e?.name || '';
        const msg = e?.message || String(e);
        const isTimeout = name === 'AbortError' || /aborted|timeout/i.test(msg);
        console.warn(
          `[papers/fetch-url] jina reader failed attempt=${attempt} errName=${name} timeout=${isTimeout} message=${msg}`,
        );
      }
    }

    if (!text || text.length < 120) {
      const r = await fetch(target, {
        headers: {
          'User-Agent': PAPER_FETCH_BROWSER_UA,
          Accept: 'text/html,application/pdf;q=0.9,*/*;q=0.8',
          'Accept-Language': 'zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7',
        },
        redirect: 'follow',
        signal: AbortSignal.timeout(22000),
      });
      if (!r.ok) {
        return res.status(502).json({ success: false, message: `拉取失败 HTTP ${r.status}` });
      }
      const ct = (r.headers.get('content-type') || '').toLowerCase();
      const buf = Buffer.from(await r.arrayBuffer());
      if (ct.includes('application/pdf') || target.toLowerCase().endsWith('.pdf')) {
        const data = await parsePdfBuffer(buf);
        text = String(data.text || '').replace(/\r\n/g, '\n');
      } else {
        text = htmlToPlainTextForPaper(buf.toString('utf8'));
      }
    }

    const cap = 400000;
    const fullLen = text.length;
    const truncated = text.length > cap;
    if (truncated) {
      text = text.slice(0, cap);
    }
    const uePaper = String(req.body?.userEmail || '').trim();
    if (uePaper) {
      incrementUserPapersProcessed(uePaper);
    }
    res.json({ success: true, text, charCount: fullLen, truncated });
  } catch (error) {
    logger.error('[papers/fetch-url]', error);
    res.status(500).json({ success: false, message: String(error?.message || error || '拉取失败') });
  }
});

app.post('/api/papers/summarize', async (req, res) => {
  try {
    const { text, title } = req.body || {};
    const bodyText = String(text || '').trim();
    if (!bodyText) {
      return res.status(400).json({ success: false, message: '请提供论文文本 text' });
    }
    const modelCfg = resolvePaperModelConfig();
    if (!modelCfg) {
      return res.status(503).json({
        success: false,
        message: '没有可用的模型配置，请在「API & 模型配置」中启用含密钥的 aliyun/openai 项。',
      });
    }
    const summary = await summarizePaper(modelCfg, { text: bodyText, title: String(title || '') });
    const userEmail = String(req.body?.userEmail || '').trim();
    let libraryItemId = null;
    if (userEmail) {
      libraryItemId = insertKnowledgeItem({
        userEmail,
        title: `${String(title || '').trim() || '文献摘要'} · 摘要`,
        itemType: 'summary',
        sourceKind: 'paper_summary',
        originalName: String(title || '').trim(),
        folderPath: '文献资料/摘要',
        contentText: bodyText.slice(0, 20000),
        summaryText: JSON.stringify(summary, null, 2),
        metadata: { title: String(title || '').trim() },
      });
    }
    return res.json({ success: true, summary, libraryItemId });
  } catch (err) {
    logger.error('[papers/summarize]', err);
    const msg = String(err?.message || err || '摘要生成失败');
    if (!res.headersSent) {
      res.status(500).json({ success: false, message: msg, error: msg });
    }
  }
});


// —— 个人资料库：模板复用 + 文件资料归档 ——
app.get('/api/library/items', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const type = String(req.query.type || '').trim();
    const q = String(req.query.q || '').trim().toLowerCase();
    const args = [email];
    let where = 'WHERE user_email = ?';
    if (type) {
      where += ' AND item_type = ?';
      args.push(type);
    }
    const rows = db
      .prepare(
        `SELECT * FROM knowledge_items ${where}
         ORDER BY folder_path ASC, datetime(updated_at) DESC, id DESC LIMIT 500`,
      )
      .all(...args)
      .map(parseLibraryRow)
      .filter((row) => !q || `${row.title} ${row.originalName} ${row.folderPath} ${row.tags.join(' ')}`.toLowerCase().includes(q));
    res.json({ success: true, items: rows });
  } catch (error) {
    logger.error('[library/items GET]', error);
    res.status(500).json({ success: false, message: '读取资料库失败' });
  }
});

app.post('/api/library/upload', libraryUpload.single('file'), async (req, res) => {
  try {
    const userEmail = String(req.body?.userEmail || '').trim();
    if (!userEmail) {
      return res.status(400).json({ success: false, message: '需要登录后上传资料' });
    }
    if (!req.file) {
      return res.status(400).json({ success: false, message: '未接收到文件' });
    }
    const ext = path.extname(req.file.originalname || '').toLowerCase();
    let contentText = '';
    if (ext === '.pdf' || ext === '.docx' || ext === '.doc') {
      try {
        contentText = (await extractPaperTextFromPath(getMulterAbsPath(req.file), ext)).slice(0, 120000);
      } catch (e) {
        logger.warn('[library/upload] extract failed', e?.message || e);
      }
    } else if (ext === '.txt' || ext === '.md') {
      try {
        contentText = (await fs.promises.readFile(getMulterAbsPath(req.file), 'utf8')).slice(0, 120000);
      } catch (e) {
        logger.warn('[library/upload] text read failed', e?.message || e);
      }
    }
    const id = insertKnowledgeItem({
      userEmail,
      title: path.basename(req.file.originalname || '资料', ext),
      itemType: 'document',
      sourceKind: 'manual_upload',
      originalName: req.file.originalname || '',
      storedFilename: req.file.filename || '',
      mimeType: req.file.mimetype || '',
      folderPath: sanitizeLibraryFolderPath(req.body?.folderPath || '手动上传'),
      tags: req.body?.tags ? String(req.body.tags).split(/[，,;；\s]+/) : [],
      contentText,
      metadata: { size: req.file.size || 0, ext, docLegacy: ext === '.doc' && !contentText },
    });
    const row = db.prepare('SELECT * FROM knowledge_items WHERE id = ?').get(id);
    res.json({ success: true, item: parseLibraryRow(row) });
  } catch (error) {
    logger.error('[library/upload POST]', error);
    res.status(500).json({ success: false, message: '资料上传失败' });
  }
});

function sanitizeKnowledgeItemTitleInput(name) {
  return String(name || '')
    .trim()
    .replace(/[/\\<>:*?"|]/g, '')
    .replace(/\u0000/g, '')
    .slice(0, 240);
}

/** 资料库文件重命名（仅改展示标题 title） */
app.put('/api/files/:id/rename', (req, res) => {
  try {
    const email = String(req.body?.email || req.query.email || '').trim();
    const id = Number.parseInt(String(req.params.id || ''), 10);
    const newTitle = sanitizeKnowledgeItemTitleInput(req.body?.newName ?? req.body?.title ?? '');
    if (!email || !Number.isFinite(id)) {
      return res.status(400).json({ success: false, message: '参数无效' });
    }
    if (!newTitle) {
      return res.status(400).json({ success: false, message: '名称不能为空或含有非法字符' });
    }
    const row = db.prepare('SELECT * FROM knowledge_items WHERE id = ? AND user_email = ?').get(id, email);
    if (!row) {
      return res.status(404).json({ success: false, message: '文件不存在或无权操作' });
    }
    db.prepare(`UPDATE knowledge_items SET title = ?, updated_at = datetime('now') WHERE id = ? AND user_email = ?`).run(newTitle, id, email);
    const next = db.prepare('SELECT * FROM knowledge_items WHERE id = ?').get(id);
    res.json({ success: true, item: parseLibraryRow(next) });
  } catch (error) {
    logger.error('[files/rename PUT]', error);
    res.status(500).json({ success: false, message: '重命名失败' });
  }
});

app.put('/api/library/items/:id', (req, res) => {
  try {
    const email = String(req.body?.email || '').trim();
    const id = Number.parseInt(String(req.params.id || ''), 10);
    if (!email || !Number.isFinite(id)) {
      return res.status(400).json({ success: false, message: '参数无效' });
    }
    const row = db.prepare('SELECT * FROM knowledge_items WHERE id = ? AND user_email = ?').get(id, email);
    if (!row) {
      return res.status(404).json({ success: false, message: '资料不存在或无权访问' });
    }
    const title = String(req.body?.title ?? row.title).trim().slice(0, 240) || row.title;
    const folderPath = sanitizeLibraryFolderPath(req.body?.folderPath ?? req.body?.folder_path ?? row.folder_path);
    const tags = sanitizeTemplateTags(req.body?.tags ?? row.tags_json);
    db.prepare('UPDATE knowledge_items SET title = ?, folder_path = ?, tags_json = ?, updated_at = datetime(\'now\') WHERE id = ?').run(
      title,
      folderPath,
      JSON.stringify(tags),
      id,
    );
    const next = db.prepare('SELECT * FROM knowledge_items WHERE id = ?').get(id);
    res.json({ success: true, item: parseLibraryRow(next) });
  } catch (error) {
    logger.error('[library/items PUT]', error);
    res.status(500).json({ success: false, message: '资料整理失败' });
  }
});

app.delete('/api/library/items/:id', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    const id = Number.parseInt(String(req.params.id || ''), 10);
    if (!email || !Number.isFinite(id)) {
      return res.status(400).json({ success: false, message: '参数无效' });
    }
    const row = db.prepare('SELECT * FROM knowledge_items WHERE id = ? AND user_email = ?').get(id, email);
    if (!row) {
      return res.status(404).json({ success: false, message: '资料不存在或无权访问' });
    }
    const filePath = resolveKnowledgeStoredFile(row.stored_filename);
    if (filePath && fs.existsSync(filePath)) {
      try { fs.unlinkSync(filePath); } catch { /* ignore */ }
    }
    db.prepare('DELETE FROM knowledge_items WHERE id = ?').run(id);
    res.json({ success: true });
  } catch (error) {
    logger.error('[library/items DELETE]', error);
    res.status(500).json({ success: false, message: '删除资料失败' });
  }
});

app.get('/api/library/items/:id/file', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    const id = Number.parseInt(String(req.params.id || ''), 10);
    if (!email || !Number.isFinite(id)) {
      return res.status(400).json({ success: false, message: '参数无效' });
    }
    const row = db.prepare('SELECT * FROM knowledge_items WHERE id = ? AND user_email = ?').get(id, email);
    if (!row || !row.stored_filename) {
      return res.status(404).json({ success: false, message: '文件不存在或无权访问' });
    }
    const filePath = resolveKnowledgeStoredFile(row.stored_filename);
    if (!filePath || !fs.existsSync(filePath)) {
      return res.status(404).json({ success: false, message: '资料文件缺失' });
    }
    const dlName = String(row.original_name || row.title || 'openclaw-file').replace(/[^\w.\-\u4e00-\u9fff]+/g, '_');
    res.download(filePath, dlName || 'openclaw-file');
  } catch (error) {
    logger.error('[library/items file]', error);
    res.status(500).json({ success: false, message: '下载资料失败' });
  }
});

app.post('/api/library/agent-organize', async (req, res) => {
  try {
    const email = String(req.body?.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const rows = db.prepare('SELECT * FROM knowledge_items WHERE user_email = ? ORDER BY datetime(updated_at) DESC, id DESC LIMIT 200').all(email);
    if (rows.length === 0) {
      return res.json({ success: true, updated: 0, items: [] });
    }
    const updates = [];
    const modelConfig = rowToModelConfig(getBestEnabledModelConfig()) || fallbackEnvModelConfig();
    if (modelConfig) {
      try {
        const compact = rows.map((r) => ({ id: r.id, title: r.title, type: r.item_type, source: r.source_kind, folder: r.folder_path, tags: parseTemplateTagsJson(r.tags_json), excerpt: String(r.summary_text || r.content_text || '').slice(0, 500) }));
        const data = await postChatCompletion(modelConfig, {
          model: normalizeModelId(modelConfig.model_name),
          messages: [
            { role: 'system', content: '你是资料库整理助手。只输出 JSON 数组，每项包含 id、folderPath、tags。folderPath 用中文短路径，可用 / 分层；tags 为最多5个中文标签。不要解释。' },
            { role: 'user', content: JSON.stringify(compact).slice(0, 30000) },
          ],
          temperature: 0.1,
        });
        const text = extractTextFromResponse(data).replace(/```json|```/g, '').trim();
        const parsed = JSON.parse(text);
        if (Array.isArray(parsed)) {
          for (const item of parsed) {
            const id = Number(item?.id);
            if (Number.isFinite(id)) {
              updates.push({ id, folderPath: sanitizeLibraryFolderPath(item.folderPath || item.folder_path || ''), tags: sanitizeTemplateTags(item.tags || []) });
            }
          }
        }
      } catch (e) {
        logger.warn('[library/agent-organize] model fallback', e?.message || e);
      }
    }
    if (updates.length === 0) {
      for (const r of rows) {
        updates.push({ id: r.id, folderPath: inferKnowledgeFolder(parseLibraryRow(r)), tags: parseTemplateTagsJson(r.tags_json) });
      }
    }
    const stmt = db.prepare('UPDATE knowledge_items SET folder_path = ?, tags_json = ?, updated_at = datetime(\'now\') WHERE id = ? AND user_email = ?');
    let updated = 0;
    for (const u of updates) {
      const folderPath = sanitizeLibraryFolderPath(u.folderPath) || '待整理';
      const tags = sanitizeTemplateTags(u.tags || []);
      const info = stmt.run(folderPath, JSON.stringify(tags), u.id, email);
      updated += Number(info.changes || 0);
    }
    const items = db.prepare('SELECT * FROM knowledge_items WHERE user_email = ? ORDER BY folder_path ASC, datetime(updated_at) DESC, id DESC LIMIT 500').all(email).map(parseLibraryRow);
    res.json({ success: true, updated, items });
  } catch (error) {
    logger.error('[library/agent-organize POST]', error);
    res.status(500).json({ success: false, message: 'Agent 整理失败' });
  }
});

// API：AI 对话接口 (对接 DeepSeek)
function flushSseResponse(res) {
  if (typeof res.flush === 'function') {
    try {
      res.flush();
    } catch {
      /* compression 未启用时无 flush */
    }
  }
}

/** 与 openclawService 中 system 附加块总预算协调：文献正文 + 结构化摘要合并后截断 */
const MAX_CHAT_SYSTEM_CONTEXT = 16000;
const MAX_PAPER_CONTEXT_BODY_CHARS = 5000;

/**
 * 会议纪要：接收已完成的实时转写文本 → Agent 结构化纪要 → 可选写入待办（不再经 HTTP 上传语音文件）
 */
function createMeetingProcessHandler() {
  return async (req, res) => {
    try {
      req.socket?.setTimeout?.(900000);
      const userEmail = String(req.body?.userEmail || '').trim();
      const transcript = String(req.body?.transcript || '').trim();
      const modelRow = getBestEnabledModelConfig();
      let modelConfig = rowToModelConfig(modelRow);
      if (!modelConfig) {
        modelConfig = fallbackEnvModelConfig();
      }
      if (!modelConfig) {
        return res.status(503).json({
          success: false,
          message: '没有可用的已启用模型配置，无法生成纪要。',
        });
      }

      if (!transcript) {
        return res.status(400).json({
          success: false,
          message: '缺少会议转写文本：请先完成实时录音转写，或粘贴逐字稿后再生成纪要。',
        });
      }

      const { markdown: minutesMarkdown, tasks } = await generateMeetingMinutesMarkdown(modelConfig, transcript);

      let todosInserted = 0;
      if (userEmail) {
        todosInserted = insertTodoRowsFromAgentTaskList(db, userEmail, tasks, 'meeting', new Date());
        if (todosInserted === 0) {
          const items = await extractTodosJson(
            modelConfig,
            `【会议纪要】\n${minutesMarkdown}\n\n【逐字稿摘录】\n${transcript.slice(0, 6000)}`,
            new Date().toISOString(),
          );
          todosInserted = insertTodoRows(db, userEmail, items, 'meeting');
        }
      }

      let libraryItemId = null;
      if (userEmail) {
        libraryItemId = insertKnowledgeItem({
          userEmail,
          title: `会议纪要 ${new Date().toLocaleString('zh-CN', { hour12: false })}`,
          itemType: 'meeting',
          sourceKind: 'meeting_minutes',
          originalName: '会议纪要',
          storedFilename: '',
          mimeType: 'text/plain',
          folderPath: '会议纪要',
          contentText: transcript,
          summaryText: minutesMarkdown,
          metadata: { todosInserted, hasAudio: false },
        });
      }

      res.json({
        success: true,
        transcript,
        minutesMarkdown,
        savedAudioPath: null,
        todosInserted,
        libraryItemId,
      });
    } catch (error) {
      logger.error('[meeting]', error);
      const msg = String(error?.message || error || '会议纪要处理失败');
      res.status(500).json({
        success: false,
        message: msg,
      });
    }
  };
}

app.post('/api/meeting/process', createMeetingProcessHandler());

/**
 * 离线音频上传：磁盘流式接收 → DashScope 实时 ASR 文件链路 → 纪要 + 结构化待办入库（不落盘长期保留音频）
 */
app.post('/api/meeting/file-transcribe', meetingFileMulterGate, async (req, res) => {
  const absPath = getMulterAbsPath(req.file);
  const unlinkSafe = () => {
    try {
      if (absPath && fs.existsSync(absPath)) {
        fs.unlinkSync(absPath);
      }
    } catch {
      /* ignore */
    }
  };
  const ac = new AbortController();
  const onClose = () => ac.abort();
  req.on('close', onClose);
  try {
    req.socket?.setTimeout?.(900000);
    if (!req.file || !absPath) {
      logger.error('[meeting/file-transcribe] 未收到上传文件', {
        hasFile: Boolean(req.file),
        contentType: req.headers['content-type'],
        fieldNames: req.body && typeof req.body === 'object' ? Object.keys(req.body) : [],
      });
      unlinkSafe();
      return res.status(400).json({ success: false, message: '请上传 .mp3 / .wav / .m4a 文件（字段名须为 file）' });
    }
    const userEmail = String(req.body?.userEmail || '').trim();
    const asrWsUrl = String(req.body?.asrWsUrl || '').trim();
    const asrApiKey = String(req.body?.asrApiKey || '').trim();
    if (!asrApiKey) {
      unlinkSafe();
      return res.status(400).json({ success: false, message: '缺少实时 ASR API Key（百炼）' });
    }
    try {
      assertSafeDashscopeRealtimeUrl(asrWsUrl || DEFAULT_REALTIME_WS);
    } catch (e) {
      unlinkSafe();
      return res.status(400).json({ success: false, message: String(e?.message || e || 'URL 无效') });
    }
    const modelRow = getBestEnabledModelConfig();
    let modelConfig = rowToModelConfig(modelRow);
    if (!modelConfig) {
      modelConfig = fallbackEnvModelConfig();
    }
    if (!modelConfig) {
      unlinkSafe();
      return res.status(503).json({ success: false, message: '没有可用的已启用模型配置，无法生成纪要。' });
    }

    const transcript = await transcribeUploadedAudioFile(absPath, {
      apiKey: asrApiKey,
      wsUrl: asrWsUrl || DEFAULT_REALTIME_WS,
      signal: ac.signal,
    });
    const text = String(transcript || '').trim();
    if (!text) {
      unlinkSafe();
      return res.status(400).json({ success: false, message: '转写结果为空，请检查音频与 ASR 配置' });
    }

    const { markdown: minutesMarkdown, tasks } = await generateMeetingMinutesMarkdown(modelConfig, text);

    let todosInserted = 0;
    if (userEmail) {
      todosInserted = insertTodoRowsFromAgentTaskList(db, userEmail, tasks, 'meeting', new Date());
      if (todosInserted === 0) {
        const items = await extractTodosJson(
          modelConfig,
          `【会议纪要】\n${minutesMarkdown}\n\n【逐字稿摘录】\n${text.slice(0, 6000)}`,
          new Date().toISOString(),
        );
        todosInserted = insertTodoRows(db, userEmail, items, 'meeting');
      }
    }

    let libraryItemId = null;
    if (userEmail) {
      libraryItemId = insertKnowledgeItem({
        userEmail,
        title: `会议纪要 ${new Date().toLocaleString('zh-CN', { hour12: false })}`,
        itemType: 'meeting',
        sourceKind: 'meeting_minutes',
        originalName: req.file?.originalname || '上传录音',
        storedFilename: '',
        mimeType: String(req.file?.mimetype || 'audio/mpeg'),
        folderPath: '会议纪要',
        contentText: text,
        summaryText: minutesMarkdown,
        metadata: { todosInserted, hasAudio: true, source: 'file-transcribe' },
      });
    }

    res.json({
      success: true,
      transcript: text,
      minutesMarkdown,
      todosInserted,
      libraryItemId,
    });
  } catch (error) {
    logger.error('[meeting/file-transcribe]', error);
    res.status(500).json({ success: false, message: String(error?.message || error || '处理失败') });
  } finally {
    req.off('close', onClose);
    unlinkSafe();
  }
});

/** PPT：由模型生成 Markdown 提纲（非流式，避免阻塞过久可单独重试） */
app.post('/api/ppt/generate-outline', async (req, res) => {
  try {
    const { topic, pages, style } = req.body || {};
    const t = String(topic || '').trim();
    if (!t) {
      return res.status(400).json({ success: false, message: '需要 topic' });
    }
    const modelRow = getBestEnabledModelConfig();
    let modelConfig = rowToModelConfig(modelRow);
    if (!modelConfig) {
      modelConfig = fallbackEnvModelConfig();
    }
    if (!modelConfig) {
      return res.status(503).json({ success: false, message: '没有可用的模型配置' });
    }
    const n = Math.min(50, Math.max(3, Number.parseInt(String(pages || '10'), 10) || 10));
    const st = String(style || '商务专业').trim().slice(0, 80);
    const markdown = await generatePptOutlineMarkdown(modelConfig, t, n, st);
    res.json({ success: true, markdown });
  } catch (error) {
    logger.error('[ppt/generate-outline]', error);
    res.status(500).json({ success: false, message: String(error?.message || error || '生成失败') });
  }
});

/** PPT：Markdown → .pptx 下载（纯 CPU，通常较快） */
app.post('/api/ppt/export', async (req, res) => {
  try {
    const { markdown, title } = req.body || {};
    const buf = await markdownToPptxBuffer(String(markdown || ''), String(title || '演示文稿').trim() || '演示文稿');
    const fname = `${String(title || 'outline').replace(/[^\w\u4e00-\u9fff\-]+/g, '_').slice(0, 60) || 'outline'}.pptx`;
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    );
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(fname)}`);
    res.send(buf);
  } catch (error) {
    logger.error('[ppt/export]', error);
    res.status(500).json({ success: false, message: String(error?.message || error || '导出失败') });
  }
});

app.get('/api/todos', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const rows = db
      .prepare(
        `SELECT id, title, remind_at AS remindAt, notes, source, done, created_at AS createdAt,
                COALESCE(status, 'pending') AS status,
                COALESCE(repeat_mode, 'once') AS repeatMode,
                COALESCE(repeat_custom_days, '[]') AS repeatCustomDays,
                COALESCE(TRIM(pushed_wecom_at), '') AS pushedWecomAt,
                COALESCE(TRIM(pushed_feishu_at), '') AS pushedFeishuAt,
                TRIM(COALESCE(last_completed_date, '')) AS lastCompletedDate
         FROM todos WHERE user_email = ? ORDER BY remind_at ASC LIMIT 500`,
      )
      .all(email);
    res.json({ success: true, items: rows });
  } catch (error) {
    logger.error('[todos GET]', error);
    res.status(500).json({ success: false, message: '读取待办失败' });
  }
});

app.get('/api/todos/attention-count', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const today = getTodayYmdShanghai();
    if (!today) {
      return res.json({ success: true, count: 0, today: '' });
    }
    const row = db
      .prepare(
        `SELECT COUNT(1) AS c FROM todos
         WHERE user_email = ?
           AND done = 0
           AND (
                COALESCE(repeat_mode, 'once') = 'once'
             OR (
                  COALESCE(repeat_mode, 'once') != 'once'
                  AND (last_completed_date IS NULL OR trim(last_completed_date) = '' OR last_completed_date != ?)
                )
           )`,
      )
      .get(email, today);
    res.json({ success: true, count: Number(row?.c || 0), today });
  } catch (error) {
    logger.error('[todos/attention-count]', error);
    res.status(500).json({ success: false, message: '读取失败' });
  }
});

app.post('/api/todos', (req, res) => {
  try {
    const {
      userEmail,
      title,
      remindAt,
      notes,
      source,
      repeatMode: repeatModeRaw,
      repeatCustomDays: repeatCustomDaysRaw,
    } = req.body || {};
    const email = String(userEmail || '').trim();
    const t = String(title || '').trim().slice(0, 500);
    const remind = String(remindAt || '').trim().slice(0, 40);
    const memo = String(notes || '').trim().slice(0, 2000);
    const src = String(source || 'manual').trim().slice(0, 40) || 'manual';
    const rm = String(repeatModeRaw || 'once').toLowerCase().trim();
    const repeatMode = ['once', 'daily', 'workday', 'custom_days'].includes(rm) ? rm : 'once';
    let repeatCustomDaysJson = '[]';
    if (repeatMode === 'custom_days') {
      const arr = Array.isArray(repeatCustomDaysRaw) ? repeatCustomDaysRaw : [];
      const norm = [];
      const seen = new Set();
      for (const x of arr) {
        const n = Number.parseInt(String(x), 10);
        if (Number.isFinite(n) && n >= 1 && n <= 7 && !seen.has(n)) {
          seen.add(n);
          norm.push(n);
        }
      }
      norm.sort((a, b) => a - b);
      repeatCustomDaysJson = JSON.stringify(norm);
      if (norm.length === 0) {
        return res.status(400).json({ success: false, message: '自定义重复请至少选择一个星期' });
      }
    }
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 userEmail' });
    }
    if (!t) {
      return res.status(400).json({ success: false, message: '需要 title' });
    }
    if (!remind || Number.isNaN(Date.parse(remind))) {
      return res.status(400).json({ success: false, message: 'remindAt 无效' });
    }
    const info = db
      .prepare(
        `INSERT INTO todos (user_email, title, remind_at, notes, source, done, repeat_mode, repeat_custom_days, status)
         VALUES (?, ?, ?, ?, ?, 0, ?, ?, 'pending')`,
      )
      .run(email, t, remind, memo, src, repeatMode, repeatCustomDaysJson);
    res.json({
      success: true,
      item: {
        id: Number(info.lastInsertRowid),
        title: t,
        remindAt: remind,
        notes: memo,
        source: src,
        done: 0,
        repeatMode,
        repeatCustomDays: repeatCustomDaysJson,
        status: 'pending',
      },
    });
  } catch (error) {
    logger.error('[todos POST]', error);
    res.status(500).json({ success: false, message: '新增待办失败' });
  }
});

app.get('/api/todos/upcoming', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const hours = Math.min(168, Math.max(1, Number.parseInt(String(req.query.hours || '72'), 10) || 72));
    const nowIso = new Date().toISOString();
    const untilIso = new Date(Date.now() + hours * 3600000).toISOString();
    const soon = db
      .prepare(
        `SELECT id, title, remind_at AS remindAt, notes, source,
                COALESCE(repeat_mode, 'once') AS repeatMode,
                COALESCE(repeat_custom_days, '[]') AS repeatCustomDays,
                COALESCE(status, 'pending') AS status
         FROM todos
         WHERE user_email = ? AND done = 0
           AND (status IS NULL OR status = '' OR status = 'pending')
           AND remind_at >= ? AND remind_at <= ?
         ORDER BY remind_at ASC LIMIT 30`,
      )
      .all(email, nowIso, untilIso);
    const overdue = db
      .prepare(
        `SELECT id, title, remind_at AS remindAt, notes, source,
                COALESCE(repeat_mode, 'once') AS repeatMode,
                COALESCE(repeat_custom_days, '[]') AS repeatCustomDays,
                COALESCE(status, 'pending') AS status
         FROM todos
         WHERE user_email = ? AND done = 0
           AND (status IS NULL OR status = '' OR status = 'pending')
           AND remind_at < ?
         ORDER BY remind_at ASC LIMIT 30`,
      )
      .all(email, nowIso);
    res.json({ success: true, items: soon, overdue, windowHours: hours });
  } catch (error) {
    logger.error('[todos/upcoming]', error);
    res.status(500).json({ success: false, message: '读取即将到来的提醒失败' });
  }
});

app.post('/api/todos/extract-from-chat', async (req, res) => {
  try {
    const { userEmail, userMessage, assistantMessage } = req.body || {};
    const ue = String(userEmail || '').trim();
    if (!ue) {
      return res.status(400).json({ success: false, message: '需要 userEmail' });
    }
    const u = String(userMessage || '').trim().slice(0, 12000);
    const a = String(assistantMessage || '').trim().slice(0, 16000);
    if (!u || !a) {
      return res.status(400).json({ success: false, message: '需要 userMessage 与 assistantMessage' });
    }
    const modelRow = getBestEnabledModelConfig();
    let modelConfig = rowToModelConfig(modelRow);
    if (!modelConfig) {
      modelConfig = fallbackEnvModelConfig();
    }
    if (!modelConfig) {
      return res.json({ success: true, inserted: 0, skipped: true });
    }
    const combined = `【用户】\n${u}\n\n【助手】\n${a}`;
    const items = await extractTodosJson(modelConfig, combined, new Date().toISOString());
    const inserted = insertTodoRows(db, ue, items, 'chat');
    res.json({ success: true, inserted });
  } catch (error) {
    logger.error('[todos/extract-from-chat]', error);
    res.status(500).json({ success: false, message: String(error?.message || error || '提取失败') });
  }
});

app.patch('/api/todos/:id/done', (req, res) => {
  try {
    const email = String(req.body?.email || req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const id = Number.parseInt(String(req.params.id || ''), 10);
    if (!Number.isFinite(id)) {
      return res.status(400).json({ success: false, message: '无效 ID' });
    }
    const row = db.prepare('SELECT repeat_mode FROM todos WHERE id = ? AND user_email = ?').get(id, email);
    if (!row) {
      return res.status(404).json({ success: false, message: '未找到待办或无权操作' });
    }
    const rm = String(row.repeat_mode || 'once').toLowerCase().trim();
    if (rm === 'once') {
      const info = db.prepare('UPDATE todos SET done = 1 WHERE id = ? AND user_email = ?').run(id, email);
      if (info.changes === 0) {
        return res.status(404).json({ success: false, message: '未找到待办或无权操作' });
      }
      return res.json({ success: true, mode: 'once' });
    }
    const today = getTodayYmdShanghai();
    if (!today) {
      return res.status(500).json({ success: false, message: '日期计算失败' });
    }
    const info = db
      .prepare('UPDATE todos SET last_completed_date = ? WHERE id = ? AND user_email = ? AND done = 0')
      .run(today, id, email);
    if (info.changes === 0) {
      return res.status(404).json({ success: false, message: '未找到待办或无权操作' });
    }
    return res.json({ success: true, mode: 'repeat', lastCompletedDate: today });
  } catch (error) {
    logger.error('[todos/done]', error);
    res.status(500).json({ success: false, message: '更新失败' });
  }
});

/** 修改提醒时间（未到点或可编辑的 pending 待办） */
app.patch('/api/todos/:id/remind-at', (req, res) => {
  try {
    const email = String(req.body?.email || req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const id = Number.parseInt(String(req.params.id || ''), 10);
    if (!Number.isFinite(id)) {
      return res.status(400).json({ success: false, message: '无效 ID' });
    }
    const remindAt = String(req.body?.remindAt || '').trim();
    if (!remindAt || Number.isNaN(Date.parse(remindAt))) {
      return res.status(400).json({ success: false, message: 'remindAt 无效' });
    }
    const info = db
      .prepare(
        `UPDATE todos SET remind_at = ?, status = 'pending',
              pushed_wecom_at = NULL, pushed_feishu_at = NULL,
              pushed_wecom_status = NULL, pushed_feishu_status = NULL,
              push_retry_count = 0
         WHERE id = ? AND user_email = ? AND done = 0`,
      )
      .run(remindAt, id, email);
    if (info.changes === 0) {
      return res.status(404).json({ success: false, message: '未找到待办或无权操作' });
    }
    res.json({ success: true, remindAt });
  } catch (error) {
    logger.error('[todos/remind-at]', error);
    res.status(500).json({ success: false, message: '更新时间失败' });
  }
});

/** 稍后提醒：基于当前时间推迟 remind_at，恢复为 pending */
app.post('/api/todos/:id/snooze', (req, res) => {
  try {
    const email = String(req.body?.email || req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const id = Number.parseInt(String(req.params.id || ''), 10);
    if (!Number.isFinite(id)) {
      return res.status(400).json({ success: false, message: '无效 ID' });
    }
    const minutes = Math.min(10080, Math.max(1, Number.parseInt(String(req.body?.minutes ?? 10), 10) || 10));
    const newRemindIso = new Date(Date.now() + minutes * 60000).toISOString();
    const info = db
      .prepare(
        `UPDATE todos SET status = 'pending', remind_at = ?,
              pushed_wecom_at = NULL, pushed_feishu_at = NULL,
              pushed_wecom_status = NULL, pushed_feishu_status = NULL,
              push_retry_count = 0
         WHERE id = ? AND user_email = ? AND done = 0`,
      )
      .run(newRemindIso, id, email);
    if (info.changes === 0) {
      return res.status(404).json({ success: false, message: '未找到待办或无权操作' });
    }
    res.json({ success: true, remindAt: newRemindIso, minutes });
  } catch (error) {
    logger.error('[todos/snooze]', error);
    res.status(500).json({ success: false, message: '稍后提醒失败' });
  }
});

app.delete('/api/todos/:id', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const id = Number.parseInt(String(req.params.id || ''), 10);
    if (!Number.isFinite(id)) {
      return res.status(400).json({ success: false, message: '无效 ID' });
    }
    const info = db.prepare('DELETE FROM todos WHERE id = ? AND user_email = ?').run(id, email);
    if (info.changes === 0) {
      return res.status(404).json({ success: false, message: '未找到待办或无权操作' });
    }
    res.json({ success: true });
  } catch (error) {
    logger.error('[todos DELETE]', error);
    res.status(500).json({ success: false, message: '删除失败' });
  }
});

function buildWecomTodoText(row) {
  const title = String(row.title || '').trim() || '待办提醒';
  const time = formatRemindAtZhBeijingYMDHM(row.remindAt || '');
  const notes = String(row.notes || '').trim();
  const who = String(row.userEmail || '').trim();
  const parts = [`⏰ 待办提醒`, `事项：${title}`];
  if (time && time !== '—') {
    parts.push(`时间：${time}`);
  }
  if (notes) {
    parts.push(`备注：${notes}`);
  }
  if (who) {
    parts.push(`用户：${who}`);
  }
  return parts.join('\n');
}

/**
 * 企微推送：将 remind_at 规范为北京时间（Asia/Shanghai）「YYYY年MM月DD日 HH:mm」，禁止裸 ISO UTC 串。
 * @param {string|undefined|null} remindAtStr
 * @returns {string}
 */
function formatRemindAtZhBeijingYMDHM(remindAtStr) {
  const ms = parseRemindAtToUtcMs(remindAtStr);
  if (!Number.isFinite(ms)) {
    return '—';
  }
  const parts = new Intl.DateTimeFormat('zh-CN', {
    timeZone: 'Asia/Shanghai',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  }).formatToParts(new Date(ms));
  const pick = (t) => parts.find((p) => p.type === t)?.value || '';
  const y = pick('year');
  const mo = pick('month');
  const da = pick('day');
  const h = pick('hour');
  const mi = pick('minute');
  if (!y || !mo || !da) {
    return '—';
  }
  return `${y}年${mo}月${da}日 ${h}:${mi}`;
}

/** 企微 Markdown 字段：单行展示，避免换行破坏机器人解析 */
function sanitizeWecomMarkdownField(s, maxLen = 800) {
  return String(s ?? '')
    .replace(/\r\n/g, '\n')
    .replace(/\n+/g, ' ')
    .trim()
    .slice(0, maxLen);
}

/**
 * 严格遵守企微群机器人 markdown 消息体（msgtype + markdown.content）
 * @param {{ title?: string, remindAt?: string }} row
 */
function buildWecomTodoMarkdownPayload(row) {
  const task = sanitizeWecomMarkdownField(row.title || '') || '待办提醒';
  const remindTime = sanitizeWecomMarkdownField(formatRemindAtZhBeijingYMDHM(row.remindAt || '')) || '—';
  const content = `**<font color="warning">待办提醒触发</font>**\n> **任务**：${task}\n> **预定时间**：${remindTime}\n\n*来自会意助手*`;
  return {
    msgtype: 'markdown',
    markdown: { content },
  };
}

const WEBHOOK_HTTP_TIMEOUT_MS = 5000;

async function httpPostJsonWithTimeout(url, body, timeoutMs = WEBHOOK_HTTP_TIMEOUT_MS) {
  const controller = new AbortController();
  const tid = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
      signal: controller.signal,
    });
  } finally {
    clearTimeout(tid);
  }
}

async function sendWecomText(webhookUrl, content) {
  try {
    const res = await httpPostJsonWithTimeout(webhookUrl, { msgtype: 'text', text: { content } });
    const data = await res.json().catch(() => ({}));
    if (!res.ok || Number(data.errcode || 0) !== 0) {
      throw new Error(data.errmsg || `HTTP ${res.status}`);
    }
    return true;
  } catch (e) {
    if (e?.name === 'AbortError') {
      throw new Error('企业微信 Webhook 请求超时');
    }
    throw e;
  }
}

/**
 * 定时推送主路径：企微 markdown（axios + 超时），失败时抛出并在 catch 中记录完整响应体
 * @param {string} wecomUrl
 * @param {{ title?: string, remindAt?: string }} row
 */
async function sendWecomTodoMarkdownWebhook(wecomUrl, row) {
  const wecomPayload = buildWecomTodoMarkdownPayload(row);
  try {
    const res = await axios.post(wecomUrl, wecomPayload, {
      headers: { 'Content-Type': 'application/json' },
      timeout: WEBHOOK_HTTP_TIMEOUT_MS,
      validateStatus: () => true,
    });
    const data = res.data;
    const errcode = data != null && typeof data === 'object' ? Number(data.errcode ?? 0) : 0;
    if (res.status < 200 || res.status >= 300) {
      logger.error('[todos/wecom] HTTP 非 2xx', {
        status: res.status,
        responseData: data,
      });
      throw new Error(`HTTP ${res.status}: ${typeof data === 'string' ? data : JSON.stringify(data)}`);
    }
    if (errcode !== 0) {
      logger.error('[todos/wecom] 企微业务错误', { errcode, errmsg: data?.errmsg, responseData: data });
      throw new Error(String(data?.errmsg || `errcode=${errcode}`));
    }
    return true;
  } catch (err) {
    const responseData = err.response?.data;
    logger.error('[todos/wecom] webhook axios 请求失败', {
      message: err.message,
      code: err.code,
      httpStatus: err.response?.status,
      responseData: responseData !== undefined ? responseData : null,
      isTimeout: err.code === 'ECONNABORTED',
    });
    if (err.code === 'ECONNABORTED') {
      throw new Error('企业微信 Webhook 请求超时');
    }
    if (responseData !== undefined) {
      const serialized =
        typeof responseData === 'string' ? responseData : JSON.stringify(responseData);
      throw new Error(serialized.slice(0, 800) || err.message);
    }
    throw err;
  }
}

/** 飞书群机器人 Webhook（文本消息） */
async function sendFeishuText(webhookUrl, content) {
  try {
    const res = await httpPostJsonWithTimeout(webhookUrl, {
      msg_type: 'text',
      content: { text: content },
    });
    let data = {};
    try {
      data = await res.json();
    } catch {
      data = {};
    }
    if (!res.ok) {
      throw new Error(data.msg || data.StatusMessage || `HTTP ${res.status}`);
    }
    const code = data.code ?? data.StatusCode;
    if (code !== undefined && Number(code) !== 0) {
      throw new Error(data.msg || data.StatusMessage || `飞书返回 code=${code}`);
    }
    return true;
  } catch (e) {
    if (e?.name === 'AbortError') {
      throw new Error('飞书 Webhook 请求超时');
    }
    throw e;
  }
}

app.get('/api/configs/webhooks', (_req, res) => {
  try {
    res.json({
      success: true,
      wecom_url: getWebhookWecomUrl(),
      feishu_url: getWebhookFeishuUrl(),
    });
  } catch (error) {
    logger.error('[configs/webhooks GET]', error);
    res.status(500).json({ success: false, message: '读取配置失败' });
  }
});

app.post('/api/configs/webhooks', (req, res) => {
  try {
    const wecom = String(req.body?.wecom_url ?? '').trim().slice(0, 4000);
    const feishu = String(req.body?.feishu_url ?? '').trim().slice(0, 4000);
    db.prepare('INSERT OR REPLACE INTO system_configs (config_key, config_value) VALUES (?, ?)').run(WEBHOOK_KEY_WECOM, wecom);
    db.prepare('INSERT OR REPLACE INTO system_configs (config_key, config_value) VALUES (?, ?)').run(WEBHOOK_KEY_FEISHU, feishu);
    res.json({ success: true });
  } catch (error) {
    logger.error('[configs/webhooks POST]', error);
    res.status(500).json({ success: false, message: '保存失败' });
  }
});

/** 将 outbound Webhook 错误整理为前端可读文案 */
function formatOutboundWebhookError(err) {
  if (err == null) {
    return '未知错误';
  }
  const code = err.code;
  if (code === 'ENOTFOUND') {
    return '域名解析失败（ENOTFOUND）：无法解析 Webhook 主机名，请检查 URL 或 DNS。';
  }
  if (code === 'ECONNREFUSED') {
    return '连接被拒绝（ECONNREFUSED）：目标端口未开放或服务不可达。';
  }
  if (code === 'ETIMEDOUT' || code === 'ECONNABORTED') {
    return '请求超时或网络中断：请检查网络或防火墙。';
  }
  if (code === 'CERT_HAS_EXPIRED' || code === 'UNABLE_TO_VERIFY_LEAF_SIGNATURE') {
    return `TLS 证书错误（${code}）：${err.message || ''}`;
  }
  const st = err.response?.status;
  const data = err.response?.data;
  const dataStr =
    data === undefined ? '' : typeof data === 'string' ? data : JSON.stringify(data);
  if (st === 404) {
    return `HTTP 404：Webhook 地址不存在或 key 无效。服务端返回：${dataStr || err.message || ''}`;
  }
  if (st === 400) {
    return `HTTP 400：请求格式错误或参数无效。服务端返回：${dataStr || err.message || ''}`;
  }
  if (st != null) {
    return `HTTP ${st}：${dataStr || err.message || '无响应体'}`;
  }
  return String(err.message || err);
}

/**
 * 即时链路测试：使用请求体中的 url（企微 text / 飞书 text）
 * POST body: { url: string, channel?: 'wecom'|'feishu' }
 */
app.post('/api/configs/test-push', async (req, res) => {
  const url = String(req.body?.url ?? '').trim().slice(0, 4000);
  const channel = String(req.body?.channel ?? 'wecom').toLowerCase().trim();
  if (!url) {
    return res.status(400).json({ success: false, message: '缺少 url：请先填写 Webhook 地址' });
  }
    const text = '会意助手链路测试成功！';
  try {
    if (channel === 'feishu') {
      const r = await axios.post(
        url,
        { msg_type: 'text', content: { text } },
        {
          headers: { 'Content-Type': 'application/json' },
          timeout: WEBHOOK_HTTP_TIMEOUT_MS,
          validateStatus: () => true,
        },
      );
      const data = r.data;
      const code = data != null && typeof data === 'object' ? Number(data.code ?? data.StatusCode ?? 0) : 0;
      if (r.status === 404) {
        return res.status(400).json({
          success: false,
          message: formatOutboundWebhookError({ response: { status: 404, data }, message: 'HTTP 404' }),
        });
      }
      if (r.status < 200 || r.status >= 300) {
        return res.status(400).json({
          success: false,
          message: `HTTP ${r.status}：${typeof data === 'string' ? data : JSON.stringify(data)}`,
        });
      }
      if (code !== 0) {
        logger.error('[configs/test-push] 飞书业务错误', { responseData: data });
        return res.status(400).json({
          success: false,
          message: `飞书接口返回错误 code=${code}：${data?.msg || data?.StatusMessage || JSON.stringify(data)}`,
        });
      }
      return res.json({ success: true });
    }
    const wecomPayload = { msgtype: 'text', text: { content: text } };
    const r = await axios.post(url, wecomPayload, {
      headers: { 'Content-Type': 'application/json' },
      timeout: WEBHOOK_HTTP_TIMEOUT_MS,
      validateStatus: () => true,
    });
    const data = r.data;
    const errcode = data != null && typeof data === 'object' ? Number(data.errcode ?? 0) : 0;
    if (r.status === 404) {
      return res.status(400).json({
        success: false,
        message: formatOutboundWebhookError({ response: { status: 404, data }, message: 'HTTP 404' }),
      });
    }
    if (r.status < 200 || r.status >= 300) {
      logger.error('[configs/test-push] 企微 HTTP 非 2xx', { status: r.status, responseData: data });
      return res.status(400).json({
        success: false,
        message: `HTTP ${r.status}：${typeof data === 'string' ? data : JSON.stringify(data)}`,
      });
    }
    if (errcode !== 0) {
      logger.error('[configs/test-push] 企微业务错误', { errcode, errmsg: data?.errmsg, responseData: data });
      return res.status(400).json({
        success: false,
        message: `企业微信返回 errcode=${errcode}：${data?.errmsg || JSON.stringify(data)}`,
      });
    }
    return res.json({ success: true });
  } catch (err) {
    const responseData = err.response?.data;
    logger.error('[configs/test-push] 请求异常', {
      message: err.message,
      code: err.code,
      httpStatus: err.response?.status,
      responseData: responseData !== undefined ? responseData : null,
    });
    return res.status(400).json({ success: false, message: formatOutboundWebhookError(err) });
  }
});

app.post('/api/configs/webhooks/test', async (req, res) => {
  try {
    const channel = String(req.body?.channel || '').toLowerCase().trim();
    let url = String(req.body?.url || '').trim().slice(0, 4000);
    if (!url) {
      if (channel === 'wecom') {
        url = getWebhookWecomUrl();
      } else if (channel === 'feishu') {
        url = getWebhookFeishuUrl();
      }
    }
    if (!url) {
      return res.status(400).json({ success: false, message: '请先填写 Webhook URL 或保存配置' });
    }
    const text = `会意助手外联推送测试 · ${new Date().toLocaleString('zh-CN', { hour12: false })}`;
    if (channel === 'wecom') {
      await sendWecomText(url, text);
    } else if (channel === 'feishu') {
      await sendFeishuText(url, text);
    } else {
      return res.status(400).json({ success: false, message: 'channel 须为 wecom 或 feishu' });
    }
    res.json({ success: true });
  } catch (error) {
    logger.error('[configs/webhooks/test] 推送失败', {
      message: error?.message,
      code: error?.code,
      httpStatus: error?.response?.status,
      responseData: error?.response?.data,
      stack: error?.stack,
    });
    res.status(400).json({ success: false, message: String(error?.message || error || '推送失败') });
  }
});

/** 将待办 remind_at 解析为 UTC 毫秒，与 Date.now() 同源比较，避免字符串与 ISO 混排导致误触发/漏触发 */
/** 东八区日历日 YYYY-MM-DD（与企微时间口径一致，用于循环待办「今日清爽」） */
function getTodayYmdShanghai() {
  try {
    const parts = new Intl.DateTimeFormat('en-CA', {
      timeZone: 'Asia/Shanghai',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
    }).formatToParts(new Date());
    const y = parts.find((p) => p.type === 'year')?.value;
    const m = parts.find((p) => p.type === 'month')?.value;
    const d = parts.find((p) => p.type === 'day')?.value;
    if (!y || !m || !d) {
      return '';
    }
    return `${y}-${m}-${d}`;
  } catch {
    return '';
  }
}

function parseRemindAtToUtcMs(remindAtStr) {
  const raw = String(remindAtStr ?? '').trim();
  if (!raw) {
    return NaN;
  }
  let ms = Date.parse(raw);
  if (Number.isFinite(ms)) {
    return ms;
  }
  const m = raw.match(/^(\d{4}-\d{2}-\d{2})[T ](\d{2}:\d{2})(?::(\d{2}))?/);
  if (m) {
    const secTail = m[3] != null ? `:${m[3]}` : ':00';
    ms = Date.parse(`${m[1]}T${m[2]}${secTail}.000Z`);
    if (Number.isFinite(ms)) {
      return ms;
    }
    ms = Date.parse(`${m[1]}T${m[2]}${secTail}`);
    if (Number.isFinite(ms)) {
      return ms;
    }
  }
  return NaN;
}

function isRemindDueNow(remindAtStr, nowMs = Date.now()) {
  const t = parseRemindAtToUtcMs(remindAtStr);
  if (!Number.isFinite(t)) {
    return false;
  }
  return t <= nowMs;
}

/** JS Date.getDay(): 0=周日 … 6=周六 → 转为 1=周一 … 7=周日 */
function weekdayMonSunFromDate(d) {
  const js = d.getDay();
  return js === 0 ? 7 : js;
}

function parseRepeatCustomDaysJson(raw) {
  try {
    const arr = JSON.parse(String(raw || '[]'));
    if (!Array.isArray(arr)) {
      return [];
    }
    const out = new Set();
    for (const x of arr) {
      const n = Number.parseInt(String(x), 10);
      if (Number.isFinite(n) && n >= 1 && n <= 7) {
        out.add(n);
      }
    }
    return [...out].sort((a, b) => a - b);
  } catch {
    return [];
  }
}

/** 循环提醒：在推送成功后计算下一次 remind_at（daily / workday / custom_days） */
function computeNextRemindIso(remindAtStr, mode, repeatCustomDaysRaw) {
  const modeNorm = String(mode || 'once').toLowerCase();
  const baseMs = parseRemindAtToUtcMs(remindAtStr);
  if (!Number.isFinite(baseMs)) {
    return new Date(Date.now() + 60000).toISOString();
  }
  if (modeNorm === 'daily') {
    return new Date(baseMs + 86400000).toISOString();
  }
  if (modeNorm === 'workday') {
    const d = new Date(baseMs);
    d.setDate(d.getDate() + 1);
    while (d.getDay() === 0 || d.getDay() === 6) {
      d.setDate(d.getDate() + 1);
    }
    return d.toISOString();
  }
  if (modeNorm === 'custom_days') {
    const days = parseRepeatCustomDaysJson(repeatCustomDaysRaw);
    if (days.length === 0) {
      return new Date(baseMs + 86400000).toISOString();
    }
    const baseDate = new Date(baseMs);
    const hh = baseDate.getHours();
    const mi = baseDate.getMinutes();
    const ss = baseDate.getSeconds();
    const ms = baseDate.getMilliseconds();
    for (let add = 1; add <= 366; add += 1) {
      const cand = new Date(baseMs + add * 86400000);
      cand.setHours(hh, mi, ss, ms);
      if (days.includes(weekdayMonSunFromDate(cand))) {
        return cand.toISOString();
      }
    }
    return new Date(baseMs + 86400000).toISOString();
  }
  return new Date(baseMs + 86400000).toISOString();
}

let todoPushInFlight = false;

async function pushSingleTodoWebhooks(row, wecomUrl, feishuUrl) {
  const textBody = buildWecomTodoText(row);
  const needWe = Boolean(wecomUrl);
  const needFe = Boolean(feishuUrl);
  let wecomOk = !needWe || Boolean(String(row.pushedWecomAt || '').trim());
  let feishuOk = !needFe || Boolean(String(row.pushedFeishuAt || '').trim());
  if (needWe && !wecomOk) {
    try {
      await sendWecomTodoMarkdownWebhook(wecomUrl, row);
      db.prepare(`UPDATE todos SET pushed_wecom_at = datetime('now'), pushed_wecom_status = ? WHERE id = ?`).run('ok', row.id);
      wecomOk = true;
    } catch (error) {
      const msg = String(error?.message || error || 'push failed').slice(0, 200);
      db.prepare('UPDATE todos SET pushed_wecom_status = ? WHERE id = ?').run(msg, row.id);
      logger.error('[todos/wecom] 单条待办推送失败', {
        todoId: row.id,
        errorMessage: error?.message,
        stack: error?.stack,
        responseData: error?.response?.data,
      });
    }
  }
  if (needFe && !feishuOk) {
    try {
      await sendFeishuText(feishuUrl, textBody);
      db.prepare(`UPDATE todos SET pushed_feishu_at = datetime('now'), pushed_feishu_status = ? WHERE id = ?`).run('ok', row.id);
      feishuOk = true;
    } catch (error) {
      const msg = String(error?.message || error || 'push failed').slice(0, 200);
      db.prepare('UPDATE todos SET pushed_feishu_status = ? WHERE id = ?').run(msg, row.id);
      logger.error('[todos/feishu] 单条待办推送失败', {
        todoId: row.id,
        errorMessage: error?.message,
        stack: error?.stack,
        responseData: error?.response?.data,
      });
    }
  }
  const allOk = (!needWe || wecomOk) && (!needFe || feishuOk);
  if (allOk) {
    const repeatMode = String(row.repeatMode || row.repeat_mode || 'once').toLowerCase();
    if (repeatMode === 'daily' || repeatMode === 'workday' || repeatMode === 'custom_days') {
      const nextAt = computeNextRemindIso(row.remindAt, repeatMode, row.repeatCustomDays || row.repeat_custom_days);
      db.prepare(
        `UPDATE todos SET remind_at = ?, status = 'pending', push_retry_count = 0,
              pushed_wecom_at = NULL, pushed_feishu_at = NULL,
              pushed_wecom_status = NULL, pushed_feishu_status = NULL
         WHERE id = ? AND done = 0`,
      ).run(nextAt, row.id);
      logger.info('[调度器] 循环待办已顺延下一次提醒', { todoId: row.id, nextAt, repeatMode });
    } else {
      db.prepare(`UPDATE todos SET status = 'notified', push_retry_count = 0 WHERE id = ? AND done = 0`).run(row.id);
      logger.info('[调度器] 待办已标记为已通知', { todoId: row.id });
    }
  } else {
    db.prepare(`UPDATE todos SET push_retry_count = COALESCE(push_retry_count, 0) + 1 WHERE id = ?`).run(row.id);
    const cnt = db.prepare('SELECT push_retry_count AS c FROM todos WHERE id = ?').get(row.id);
    logger.info('[调度器] 待办推送未完全成功，已增加 push_retry_count', {
      todoId: row.id,
      pushRetryCount: cnt?.c,
    });
  }
}

async function runTodoWebhookPush() {
  logger.info(`[调度器] 正在扫描待办... 当前系统时间：${new Date().toLocaleString('zh-CN', { hour12: false })}`);
  if (todoPushInFlight) {
    logger.info('[调度器] 上一轮推送尚未结束，跳过本轮');
    return;
  }
  let wecomUrl = '';
  let feishuUrl = '';
  try {
    wecomUrl = getWebhookWecomUrl();
    feishuUrl = getWebhookFeishuUrl();
  } catch (readErr) {
    logger.warn('[todos/webhooks] read system_configs failed', readErr?.message || readErr);
    return;
  }
  if (!wecomUrl && !feishuUrl) {
    logger.info('[调度器] Webhook URL：企微与飞书均为空，跳过推送');
    return;
  }
  logger.info(
    `[调度器] 正在获取 Webhook URL... 企业微信：${wecomUrl ? '已获取' : '为空'}；飞书：${feishuUrl ? '已获取' : '为空'}`,
  );
  todoPushInFlight = true;
  const nowMs = Date.now();
  try {
    const candidates = db
      .prepare(
        `SELECT id, user_email AS userEmail, title, remind_at AS remindAt, notes,
                pushed_wecom_at AS pushedWecomAt, pushed_feishu_at AS pushedFeishuAt,
                COALESCE(push_retry_count, 0) AS pushRetryCount,
                COALESCE(repeat_mode, 'once') AS repeatMode,
                COALESCE(repeat_custom_days, '[]') AS repeatCustomDays
         FROM todos
         WHERE done = 0
           AND (status IS NULL OR status = '' OR status = 'pending')
           AND COALESCE(push_retry_count, 0) < ?
           AND (
             (? != '' AND (pushed_wecom_at IS NULL OR TRIM(pushed_wecom_at) = ''))
             OR (? != '' AND (pushed_feishu_at IS NULL OR TRIM(pushed_feishu_at) = ''))
           )
         ORDER BY remind_at ASC LIMIT ?`,
      )
      .all(TODO_PUSH_MAX_RETRY, wecomUrl, feishuUrl, TODO_REMIND_CANDIDATE_SQL_LIMIT);
    logger.info(`[调度器] 从数据库找到 ${candidates.length} 条候选待办`);
    const due = candidates.filter((row) => isRemindDueNow(row.remindAt, nowMs));
    if (due.length > 0) {
      logger.info('[调度器] 时间条件命中，待推送条数', {
        count: due.length,
        wecomUrlState: wecomUrl ? '已获取' : '为空',
        feishuUrlState: feishuUrl ? '已获取' : '为空',
      });
    }
    await Promise.allSettled(due.map((row) => pushSingleTodoWebhooks(row, wecomUrl, feishuUrl)));
  } catch (error) {
    logger.error('[todos/webhooks] scheduler failed', {
      message: error?.message,
      stack: error?.stack,
      responseData: error?.response?.data,
    });
  } finally {
    todoPushInFlight = false;
  }
}

/** —— 对话会话历史 —— */
app.get('/api/chat/sessions', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const rows = db
      .prepare(
        `SELECT id, title, persona_id AS personaId, persona_label AS personaLabel, persona_emoji AS personaEmoji,
                skill_id AS skillId, skill_name AS skillName, created_at AS createdAt, updated_at AS updatedAt
         FROM chat_sessions WHERE user_email = ? ORDER BY datetime(updated_at) DESC LIMIT 200`,
      )
      .all(email);
    res.json({ success: true, items: rows });
  } catch (error) {
    logger.error('[chat/sessions GET]', error);
    res.status(500).json({ success: false, message: '读取会话列表失败' });
  }
});

app.post('/api/chat/sessions', (req, res) => {
  try {
    const email = String(req.body?.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const title = String(req.body?.title || '新对话').trim().slice(0, 200) || '新对话';
    const personaId = String(req.body?.personaId || '').trim().slice(0, 64);
    const personaLabel = String(req.body?.personaLabel || '').trim().slice(0, 80);
    const personaEmoji = String(req.body?.personaEmoji || '').trim().slice(0, 8);
    const skillId = req.body?.skillId != null ? String(req.body.skillId).trim().slice(0, 64) : '';
    const skillName = String(req.body?.skillName || '').trim().slice(0, 120);
    const info = db
      .prepare(
        `INSERT INTO chat_sessions (user_email, title, persona_id, persona_label, persona_emoji, skill_id, skill_name, created_at, updated_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'), datetime('now'))`,
      )
      .run(email, title, personaId, personaLabel, personaEmoji, skillId, skillName);
    const row = db
      .prepare(
        `SELECT id, title, persona_id AS personaId, persona_label AS personaLabel, persona_emoji AS personaEmoji,
                skill_id AS skillId, skill_name AS skillName, created_at AS createdAt, updated_at AS updatedAt
         FROM chat_sessions WHERE id = ?`,
      )
      .get(info.lastInsertRowid);
    res.json({ success: true, session: row });
  } catch (error) {
    logger.error('[chat/sessions POST]', error);
    res.status(500).json({ success: false, message: '创建会话失败' });
  }
});

app.delete('/api/chat/sessions/all', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const sess = db.prepare('SELECT id FROM chat_sessions WHERE user_email = ?').all(email);
    for (const s of sess) {
      db.prepare('DELETE FROM chat_messages WHERE session_id = ?').run(s.id);
    }
    db.prepare('DELETE FROM chat_sessions WHERE user_email = ?').run(email);
    res.json({ success: true, deletedSessions: sess.length });
  } catch (error) {
    logger.error('[chat/sessions DELETE all]', error);
    res.status(500).json({ success: false, message: '清空会话失败' });
  }
});

app.delete('/api/chat/sessions/:id', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const id = Number.parseInt(String(req.params.id || ''), 10);
    if (!Number.isFinite(id)) {
      return res.status(400).json({ success: false, message: '无效会话 ID' });
    }
    const row = db.prepare('SELECT id FROM chat_sessions WHERE id = ? AND user_email = ?').get(id, email);
    if (!row) {
      return res.status(404).json({ success: false, message: '会话不存在或无权删除' });
    }
    db.prepare('DELETE FROM chat_messages WHERE session_id = ?').run(id);
    db.prepare('DELETE FROM chat_sessions WHERE id = ?').run(id);
    res.json({ success: true });
  } catch (error) {
    logger.error('[chat/sessions DELETE]', error);
    res.status(500).json({ success: false, message: '删除会话失败' });
  }
});

app.get('/api/chat/sessions/:id/messages', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const id = Number.parseInt(String(req.params.id || ''), 10);
    if (!Number.isFinite(id)) {
      return res.status(400).json({ success: false, message: '无效会话 ID' });
    }
    const sess = db.prepare('SELECT id FROM chat_sessions WHERE id = ? AND user_email = ?').get(id, email);
    if (!sess) {
      return res.status(404).json({ success: false, message: '会话不存在或无权访问' });
    }
    const rows = db
      .prepare(
        `SELECT id, role, content, extra_json AS extraJson, sort_order AS sortOrder
         FROM chat_messages WHERE session_id = ? ORDER BY sort_order ASC, id ASC`,
      )
      .all(id);
    res.json({ success: true, items: rows });
  } catch (error) {
    logger.error('[chat/messages GET]', error);
    res.status(500).json({ success: false, message: '读取消息失败' });
  }
});

app.post('/api/chat/sessions/:id/append-turn', async (req, res) => {
  try {
    const email = String(req.body?.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const id = Number.parseInt(String(req.params.id || ''), 10);
    if (!Number.isFinite(id)) {
      return res.status(400).json({ success: false, message: '无效会话 ID' });
    }
    const sess = db.prepare('SELECT id, title, persona_id, skill_id FROM chat_sessions WHERE id = ? AND user_email = ?').get(id, email);
    if (!sess) {
      return res.status(404).json({ success: false, message: '会话不存在或无权写入' });
    }
    const userText = String(req.body?.userText || '').trim().slice(0, 24000);
    const assistant = req.body?.assistant && typeof req.body.assistant === 'object' ? req.body.assistant : {};
    const sessionMeta = req.body?.sessionMeta && typeof req.body.sessionMeta === 'object' ? req.body.sessionMeta : {};
    const chatMode = String(req.body?.chatMode || sessionMeta.chatMode || '').trim().toLowerCase();
    const aContentRaw = String(assistant.content ?? '').trim().slice(0, 64000);
    const aContent = aContentRaw || '（本轮暂无可见正文）';
    if (!userText) {
      return res.status(400).json({ success: false, message: '需要 userText' });
    }
    const extra = {
      thoughtChain: String(assistant.thoughtChain || '').slice(0, 120000),
      clarificationCard: assistant.clarificationCard || null,
      isHtml: !!assistant.isHtml,
      streamFailed: !!assistant.streamFailed,
    };
    const extraStr = JSON.stringify(extra);
    const maxRow = db.prepare('SELECT COALESCE(MAX(sort_order), -1) AS m FROM chat_messages WHERE session_id = ?').get(id);
    const base = Number(maxRow?.m ?? -1) + 1;
    db.prepare(
      `INSERT INTO chat_messages (session_id, role, content, extra_json, sort_order)
       VALUES (?, 'user', ?, NULL, ?)`,
    ).run(id, userText, base);
    db.prepare(
      `INSERT INTO chat_messages (session_id, role, content, extra_json, sort_order)
       VALUES (?, 'assistant', ?, ?, ?)`,
    ).run(id, aContent, extraStr, base + 1);
    let newTitle = String(sess.title || '').trim();
    if (!newTitle || newTitle === '新对话') {
      newTitle = await generateChatTitleFromFirstTurn(userText, aContent);
    }
    const personaId = String(sessionMeta.personaId || '').trim().slice(0, 64);
    const personaLabel = String(sessionMeta.personaLabel || '').trim().slice(0, 80);
    const personaEmoji = String(sessionMeta.personaEmoji || '').trim().slice(0, 8);
    const skillId = sessionMeta.skillId != null ? String(sessionMeta.skillId).trim().slice(0, 64) : '';
    const skillName = String(sessionMeta.skillName || '').trim().slice(0, 120);
    db.prepare(
      `UPDATE chat_sessions
       SET title = ?,
           persona_id = CASE WHEN COALESCE(persona_id, '') = '' THEN ? ELSE persona_id END,
           persona_label = CASE WHEN COALESCE(persona_label, '') = '' THEN ? ELSE persona_label END,
           persona_emoji = CASE WHEN COALESCE(persona_emoji, '') = '' THEN ? ELSE persona_emoji END,
           skill_id = CASE WHEN COALESCE(skill_id, '') = '' THEN ? ELSE skill_id END,
           skill_name = CASE WHEN COALESCE(skill_name, '') = '' THEN ? ELSE skill_name END,
           updated_at = datetime('now')
       WHERE id = ?`,
    ).run(newTitle, personaId, personaLabel, personaEmoji, skillId, skillName, id);
    if (chatMode !== 'companion') {
      scheduleChatTodoExtraction({ userEmail: email, userMessage: userText, assistantMessage: aContent, sessionId: id });
    }
    res.json({ success: true, title: newTitle, todosExtractionQueued: chatMode !== 'companion' });
  } catch (error) {
    logger.error('[chat/append-turn]', error);
    res.status(500).json({ success: false, message: String(error?.message || error || '保存失败') });
  }
});

app.post('/api/chat', async (req, res) => {
  const {
    message,
    chatMode,
    systemContext,
    temperatureProfile,
    agentContext,
    enableTools,
    paperContext,
    userEmail,
    skillId,
    clientLocalTimeLabel,
    companionPersona,
    companionAddress,
    companionCustomTraits,
    personaId,
    personaPrompt,
    personaLabel,
    personaEmoji,
  } = req.body || {};
  const paperSlice = String(paperContext || '')
    .trim()
    .slice(0, MAX_PAPER_CONTEXT_BODY_CHARS);
  const summaryPart = String(systemContext || '').trim();
  const mergedParts = [];
  if (paperSlice) {
    mergedParts.push(
      `【文献正文摘录（本轮请求已截断至前 ${MAX_PAPER_CONTEXT_BODY_CHARS} 字符，与前端工作台解析结果一致）】\n${paperSlice}`,
    );
  }
  if (summaryPart) {
    mergedParts.push(summaryPart);
  }
  const addon = mergedParts.join('\n\n').slice(0, MAX_CHAT_SYSTEM_CONTEXT);
  const litProfile = String(temperatureProfile || '').toLowerCase() === 'literature' ? 'literature' : '';

  let userProfileForChat = null;
  const ue = String(userEmail || '').trim();
  if (ue) {
    try {
      const urow = db.prepare('SELECT user_profile FROM users WHERE email = ?').get(ue);
      if (urow) {
        userProfileForChat = parseStoredUserProfile(urow.user_profile);
      }
    } catch (e) {
      console.warn('[chat] load user_profile skipped', e?.message || e);
    }
  }

  let skillSystemPromptForChat = '';
  const sidRaw = skillId != null && skillId !== '' ? Number.parseInt(String(skillId), 10) : NaN;
  if (Number.isFinite(sidRaw) && ue) {
    try {
      const srow = db.prepare('SELECT system_prompt, user_email FROM skills WHERE id = ?').get(sidRaw);
      if (srow && String(srow.user_email || '') === ue) {
        skillSystemPromptForChat = String(srow.system_prompt || '').trim();
        try {
          db.prepare('UPDATE skills SET use_count = COALESCE(use_count, 0) + 1 WHERE id = ? AND user_email = ?').run(
            sidRaw,
            ue,
          );
        } catch (ucErr) {
          console.warn('[chat] skill use_count', ucErr?.message || ucErr);
        }
      }
    } catch (e) {
      console.warn('[chat] skill load skipped', e?.message || e);
    }
  }

  const msgTrimmed = String(message || '').trim();
  /** 澄清卡片「确认并继续」发来的正文：启用后续交付约束，并在 Agent 内关闭工具调用，避免第二轮追问未走卡片 */
  const clarificationFollowUp =
    String(chatMode || '').toLowerCase() !== 'companion' && /关于你的疑问[，,]?\s*我的补充是/.test(msgTrimmed);

  const isCompanionMode = String(chatMode || '').toLowerCase() === 'companion';
  let companionOpt = {};
  if (isCompanionMode) {
    const prof = userProfileForChat || defaultUserProfile();
    const pk = String(
      companionPersona != null && companionPersona !== ''
        ? companionPersona
        : prof.companion_persona || 'default',
    ).trim() || 'default';
    const addr = String(companionAddress != null ? companionAddress : prof.companion_address || '')
      .trim()
      .slice(0, 64);
    const traits = String(
      companionCustomTraits != null ? companionCustomTraits : prof.companion_custom_traits || '',
    )
      .trim()
      .slice(0, 1200);
    companionOpt = { companionPersona: pk, companionAddress: addr, companionCustomTraits: traits };
  }

  const streamChatOptions = {
    ...(isCompanionMode ? { chatMode: 'companion', ...companionOpt } : {}),
    ...(addon ? { systemAddon: addon } : {}),
    ...(litProfile ? { temperatureProfile: litProfile } : {}),
    ...(userProfileForChat ? { userProfile: userProfileForChat } : {}),
    ...(clarificationFollowUp ? { clarificationFollowUp: true } : {}),
    ...(skillSystemPromptForChat ? { skillSystemPrompt: skillSystemPromptForChat } : {}),
    ...(typeof clientLocalTimeLabel === 'string' && clientLocalTimeLabel.trim()
      ? { clientLocalTimeLabel: clientLocalTimeLabel.trim().slice(0, 200) }
      : {}),
  };
  if (!isCompanionMode) {
    const pid = String(personaId || '').trim().slice(0, 64);
    if (pid) {
      streamChatOptions.chatPersonaId = pid;
      if (pid.startsWith('custom_')) {
        streamChatOptions.chatPersonaPrompt = String(personaPrompt || '').trim().slice(0, 2500);
        streamChatOptions.chatPersonaLabel = String(personaLabel || '').trim().slice(0, 80);
        streamChatOptions.chatPersonaEmoji = String(personaEmoji || '').trim().slice(0, 8);
      }
    }
  }

  if (!message || !msgTrimmed) {
    return res.status(400).json({ reply: '请输入内容' });
  }

  const modelRow = getBestEnabledModelConfig();
  const modelConfig = rowToModelConfig(modelRow);
  if (!modelConfig) {
    return res.status(503).json({
      success: false,
      message: '没有可用的已启用模型配置（需填写 API Key）。请到「API & 模型配置」中添加并启用。',
    });
  }

  res.writeHead(200, {
    'Content-Type': 'text/event-stream; charset=utf-8',
    'Cache-Control': 'no-cache, no-transform',
    Pragma: 'no-cache',
    Connection: 'keep-alive',
    'X-Accel-Buffering': 'no',
    'X-Accel-Filtering': 'no',
    'Content-Encoding': 'none',
  });
  res.write(':ok\n\n');
  flushSseResponse(res);
  res.socket?.setNoDelay?.(true);

  try {
    const isCompanion = String(chatMode || '').toLowerCase() === 'companion';
    const useAgentTools = !isCompanion && enableTools !== false;

    if (useAgentTools) {
      const ctx = agentContext && typeof agentContext === 'object' && !Array.isArray(agentContext) ? agentContext : {};
      const executeTool = createToolExecutor({
        db,
        uploadsDir,
        tempDir,
        agentContext: ctx,
      });
      await runAgentChatToSse(res, {
        userMessage: msgTrimmed,
        modelConfig,
        streamOptions: streamChatOptions,
        executeTool,
        flushSse: () => flushSseResponse(res),
      });
      return res.end();
    }

    const { reader, response } = await sendChatStream(message, modelConfig, streamChatOptions);
    if (!response.ok || !reader) {
      let errorBody = '';
      try {
        const resClone = response.clone();
        errorBody = await resClone.text();
      } catch (readErr) {
        errorBody = String(readErr?.message || '无法读取上游错误详情');
      }
      logger.error('[chat stream] upstream error:', response.status, errorBody);
      res.write('event: error\ndata: 上游请求失败\n\n');
      flushSseResponse(res);
      return res.end();
    }
    const decoder = new TextDecoder('utf-8');
    let buffer = '';

    while (true) {
      const { value, done } = await reader.read();
      if (value?.byteLength) {
        const chunk = decoder.decode(value, { stream: !done });
        buffer += chunk;
        const lines = buffer.split('\n');
        buffer = lines.pop() || '';

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed.startsWith('data:')) {
            continue;
          }

          const payload = trimmed.replace(/^data:\s*/i, '').trim();
          if (!payload || payload === '[DONE]') {
            continue;
          }

          let payloads = [];
          try {
            payloads = parseSsePayloads(payload);
          } catch {
            continue;
          }
          for (const json of payloads) {
            try {
              const delta = json?.choices?.[0]?.delta;
              if (!delta || typeof delta !== 'object') {
                continue;
              }
              const reasoning = delta.reasoning_content;
              if (typeof reasoning === 'string' && reasoning.length > 0) {
                res.write(`data: ${JSON.stringify({ type: 'thought', text: reasoning })}\n\n`);
                flushSseResponse(res);
              }
              const contentPiece = delta.content;
              if (typeof contentPiece === 'string' && contentPiece.length > 0) {
                res.write(`data: ${JSON.stringify({ type: 'delta', text: contentPiece })}\n\n`);
                flushSseResponse(res);
              }
            } catch (chunkErr) {
              console.warn('[chat stream] chunk skipped:', chunkErr?.message || chunkErr);
            }
          }
        }
      }

      if (done) {
        break;
      }
    }

    res.write(`data: ${JSON.stringify({ type: 'done' })}\n\n`);
    flushSseResponse(res);
    res.end();
  } catch (error) {
    logger.error('AI SSE 接口调用失败:', error);
    res.write('event: error\ndata: SSE 请求失败\n\n');
    flushSseResponse(res);
    res.end();
  }
});

app.post('/api/templates/upload', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, message: '未接收到文件' });
    }

    const ue = String(req.body?.userEmail || '').trim() || null;
    const stmt = db.prepare('INSERT INTO templates (filename, original_name, user_email, folder_path, tags_json) VALUES (?, ?, ?, ?, ?)');
    const info = stmt.run(req.file.filename, req.file.originalname, ue, '', '[]');

    res.json({
      success: true,
      template: {
        id: info.lastInsertRowid,
        filename: req.file.filename,
        originalName: req.file.originalname,
        user_email: ue,
      },
    });
  } catch (error) {
    logger.error('[template upload] failed', error);
    res.status(500).json({ success: false, message: '模板上传失败' });
  }
});

app.post('/api/templates/upload-folder', upload.array('files', 200), (req, res) => {
  try {
    const files = Array.isArray(req.files) ? req.files : [];
    if (files.length === 0) {
      return res.status(400).json({ success: false, message: '未接收到模板文件夹中的 .docx 文件' });
    }
    const ue = String(req.body?.userEmail || '').trim() || null;
    const insertStmt = db.prepare('INSERT INTO templates (filename, original_name, user_email, folder_path, tags_json) VALUES (?, ?, ?, ?, ?)');
    const items = [];
    for (const file of files) {
      const relativePath = String(file.originalname || file.filename || 'template.docx').replace(/\\/g, '/');
      const parts = relativePath.split('/').filter(Boolean);
      const displayName = parts.pop() || path.basename(relativePath || 'template.docx');
      const folderPath = parts.join('/').slice(0, 500);
      const info = insertStmt.run(file.filename, displayName, ue, folderPath, '[]');
      items.push({
        id: Number(info.lastInsertRowid),
        filename: file.filename,
        originalName: displayName,
        folderPath,
        tags: [],
        user_email: ue,
      });
    }
    logger.info('[templates/upload-folder]', { userEmail: ue, count: items.length });
    res.json({ success: true, items, count: items.length });
  } catch (error) {
    logger.error('[template upload-folder] failed', error);
    res.status(500).json({ success: false, message: '模板文件夹上传失败' });
  }
});

/** 当前用户的 Word 模板列表（不含「模板市场」占位） */
app.get('/api/templates/my', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const rows = db
      .prepare(
        `SELECT id, original_name AS originalName, folder_path AS folderPath, tags_json AS tagsJson,
                uploaded_at AS uploadedAt
         FROM templates WHERE user_email = ? ORDER BY folder_path ASC, id DESC LIMIT 200`,
      )
      .all(email)
      .map((row) => ({ ...row, tags: parseTemplateTagsJson(row.tagsJson) }));
    res.json({ success: true, items: rows });
  } catch (error) {
    logger.error('[templates/my]', error);
    res.status(500).json({ success: false, message: '读取模板列表失败' });
  }
});

/** 下载模板原文 .docx（校验归属） */
app.get('/api/templates/:id/file', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const id = Number.parseInt(String(req.params.id || ''), 10);
    if (!Number.isFinite(id)) {
      return res.status(400).json({ success: false, message: '无效 ID' });
    }
    const row = db.prepare('SELECT * FROM templates WHERE id = ?').get(id);
    if (!row) {
      return res.status(404).json({ success: false, message: '模板不存在' });
    }
    if (row.user_email && String(row.user_email) !== email) {
      return res.status(403).json({ success: false, message: '无权访问该模板' });
    }
    const filePath = path.join(uploadsDir, row.filename);
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ success: false, message: '模板文件缺失' });
    }
    const dlName = String(row.original_name || 'template.docx').replace(/[^\w.\-\u4e00-\u9fff]+/g, '_');
    res.download(filePath, dlName || 'template.docx');
  } catch (error) {
    logger.error('[templates/file]', error);
    res.status(500).json({ success: false, message: '下载失败' });
  }
});

app.put('/api/templates/:id', (req, res) => {
  try {
    const email = String(req.body?.email || '').trim();
    const originalName = String(req.body?.original_name || '').trim().slice(0, 240);
    const folderPath = String(req.body?.folder_path ?? req.body?.folderPath ?? '').trim().replace(/\\/g, '/').slice(0, 500);
    const tags = sanitizeTemplateTags(req.body?.tags ?? req.body?.tags_json ?? req.body?.tagsJson ?? []);
    if (!email || !originalName) {
      return res.status(400).json({ success: false, message: '需要 email 与 original_name' });
    }
    const id = Number.parseInt(String(req.params.id || ''), 10);
    if (!Number.isFinite(id)) {
      return res.status(400).json({ success: false, message: '无效 ID' });
    }
    const row = db.prepare('SELECT id, user_email FROM templates WHERE id = ?').get(id);
    if (!row || String(row.user_email || '') !== email) {
      return res.status(403).json({ success: false, message: '无权修改该模板' });
    }
    db.prepare('UPDATE templates SET original_name = ?, folder_path = ?, tags_json = ? WHERE id = ?').run(originalName, folderPath, JSON.stringify(tags), id);
    res.json({ success: true });
  } catch (error) {
    logger.error('[templates PUT]', error);
    res.status(500).json({ success: false, message: '更新失败' });
  }
});

app.delete('/api/templates/:id', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const id = Number.parseInt(String(req.params.id || ''), 10);
    if (!Number.isFinite(id)) {
      return res.status(400).json({ success: false, message: '无效 ID' });
    }
    const row = db.prepare('SELECT * FROM templates WHERE id = ?').get(id);
    if (!row || String(row.user_email || '') !== email) {
      return res.status(403).json({ success: false, message: '无权删除该模板' });
    }
    const filePath = path.join(uploadsDir, row.filename);
    try {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    } catch {
      /* ignore */
    }
    db.prepare('DELETE FROM templates WHERE id = ?').run(id);
    res.json({ success: true });
  } catch (error) {
    logger.error('[templates DELETE]', error);
    res.status(500).json({ success: false, message: '删除失败' });
  }
});

app.get('/api/templates/fill-records', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    if (!email) {
      return res.status(400).json({ success: false, message: '需要 email' });
    }
    const limit = Math.min(100, Math.max(1, Number.parseInt(String(req.query.limit || '30'), 10) || 30));
    const rows = db
      .prepare(
        `SELECT id, template_id AS templateId, template_name AS templateName, variables_json AS variablesJson,
                context_json AS contextJson, created_at AS createdAt
         FROM template_fill_records
         WHERE user_email = ?
         ORDER BY datetime(created_at) DESC, id DESC
         LIMIT ?`,
      )
      .all(email, limit)
      .map((row) => ({
        ...row,
        variables: (() => {
          try { return JSON.parse(row.variablesJson || '{}'); } catch { return {}; }
        })(),
      }));
    res.json({ success: true, items: rows });
  } catch (error) {
    logger.error('[templates/fill-records GET]', error);
    res.status(500).json({ success: false, message: '读取模板生成历史失败' });
  }
});

app.get('/api/templates/fill-records/:id/file', (req, res) => {
  try {
    const email = String(req.query.email || '').trim();
    const id = Number.parseInt(String(req.params.id || ''), 10);
    if (!email || !Number.isFinite(id)) {
      return res.status(400).json({ success: false, message: '参数无效' });
    }
    const row = db.prepare('SELECT * FROM template_fill_records WHERE id = ? AND user_email = ?').get(id, email);
    if (!row) {
      return res.status(404).json({ success: false, message: '历史记录不存在或无权访问' });
    }
    const filePath = path.join(uploadsDir, row.filled_filename);
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ success: false, message: '历史文件缺失' });
    }
    const dlName = `${String(row.template_name || 'filled-template').replace(/\.[^.]+$/i, '').replace(/[^\w.\-\u4e00-\u9fff]+/g, '_') || 'filled-template'}-已填充.docx`;
    res.download(filePath, dlName);
  } catch (error) {
    logger.error('[templates/fill-records/file]', error);
    res.status(500).json({ success: false, message: '下载历史文件失败' });
  }
});

app.post('/api/templates/render-pdf', async (req, res) => {
  const { templateId, variables } = req.body || {};
  if (!templateId) {
    return res.status(400).json({ success: false, message: '模板 ID 缺失' });
  }

  const stmt = db.prepare('SELECT * FROM templates WHERE id = ?');
  const template = stmt.get(templateId);
  if (!template) {
    return res.status(404).json({ success: false, message: '模板不存在' });
  }

  const filePath = path.join(uploadsDir, template.filename);
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ success: false, message: '模板文件缺失' });
  }

  const sofficePath = process.env.SOFFICE_PATH;
  if (!sofficePath) {
    return res.status(500).json({ success: false, message: 'SOFFICE_PATH 未配置' });
  }

  const args = [
    pythonScript,
    filePath,
    JSON.stringify(variables || {}),
    tempDir,
    sofficePath,
  ];

  const child = spawn(pythonBin, args, { windowsHide: true });
  let stdout = '';
  let stderr = '';
  let finished = false;

  const timeout = setTimeout(() => {
    if (finished) {
      return;
    }
    finished = true;
    child.kill('SIGKILL');
    res.status(504).json({ success: false, message: 'PDF 转换超时' });
  }, 60000);

  child.stdout.on('data', (data) => {
    stdout += data.toString();
  });

  child.stderr.on('data', (data) => {
    stderr += data.toString();
  });

  child.on('error', (error) => {
    if (finished) {
      return;
    }
    finished = true;
    clearTimeout(timeout);
    logger.error('[pdf render] spawn failed', error);
    res.status(500).json({ success: false, message: 'PDF 转换失败' });
  });

  child.on('close', (code) => {
    if (finished) {
      return;
    }
    finished = true;
    clearTimeout(timeout);

    if (code !== 0) {
      logger.error('[pdf render] non-zero exit', code, stderr);
      return res.status(500).json({ success: false, message: 'PDF 转换失败' });
    }

    let payload;
    try {
      payload = JSON.parse(stdout.trim());
    } catch (error) {
      logger.error('[pdf render] invalid output', stdout);
      return res.status(500).json({ success: false, message: 'PDF 转换失败' });
    }

    const pdfPath = payload?.pdfPath;
    if (!pdfPath || !fs.existsSync(pdfPath)) {
      return res.status(500).json({ success: false, message: 'PDF 文件缺失' });
    }

    res.sendFile(pdfPath, () => {
      fs.unlink(pdfPath, () => {});
    });
  });
});

function extractTemplateVariables(text) {
  if (!text) {
    return [];
  }

  const found = new Set();
  const normalized = String(text || '');
  const addName = (raw) => {
    const value = String(raw || '')
      .replace(/[{}]/g, '')
      .trim();
    if (value && value.length <= 120 && !/[\r\n]/.test(value)) {
      found.add(value);
    }
  };

  // 需求标准：半角双花括号 {{变量名}}
  const doubleRegex = /\{\{\s*([^{}]+?)\s*\}\}/g;
  let match;
  while ((match = doubleRegex.exec(normalized)) !== null) {
    addName(match[1]);
  }

  // 兼容历史模板：单花括号 {变量名}，但排除 JSON/大段文本导致的误识别。
  const singleRegex = /(?<!\{)\{\s*([^{}]+?)\s*\}(?!\})/g;
  while ((match = singleRegex.exec(normalized)) !== null) {
    addName(match[1]);
  }

  return Array.from(found);
}

function normalizeDocxTemplateBuffer(buffer) {
  const zip = new PizZip(buffer);
  const files = ['word/document.xml'];
  const rels = zip.file(/word\/(header|footer)\d+\.xml/);
  for (const f of rels) {
    files.push(f.name);
  }
  for (const name of files) {
    const file = zip.file(name);
    if (!file) {
      continue;
    }
    const xml = file.asText();
    const normalized = xml.replace(/\{\{\s*([^{}]+?)\s*\}\}/g, (_m, key) => `{${String(key).trim()}}`);
    if (normalized !== xml) {
      zip.file(name, normalized);
    }
  }
  return zip.generate({ type: 'nodebuffer' });
}

function persistTemplateFillRecord({ templateId, userEmail, templateName, variables, context, filledDocxBase64 }) {
  const ue = String(userEmail || '').trim();
  if (!ue || !filledDocxBase64) {
    return null;
  }
  const safeBase = String(templateName || 'filled-template')
    .replace(/\.[^.]+$/i, '')
    .replace(/[^\w\u4e00-\u9fff\-]+/g, '_')
    .slice(0, 80) || 'filled-template';
  const filledFilename = `${Date.now()}-${crypto.randomUUID()}-${safeBase}.docx`;
  const abs = path.join(uploadsDir, filledFilename);
  fs.writeFileSync(abs, Buffer.from(String(filledDocxBase64), 'base64'));
  const info = db
    .prepare(
      `INSERT INTO template_fill_records (template_id, user_email, template_name, variables_json, context_json, filled_filename)
       VALUES (?, ?, ?, ?, ?, ?)`,
    )
    .run(
      Number.isFinite(Number(templateId)) ? Number(templateId) : null,
      ue,
      String(templateName || '').slice(0, 240),
      JSON.stringify(variables || {}),
      JSON.stringify(context || {}),
      filledFilename,
    );
  return Number(info.lastInsertRowid);
}

function buildFillPrompt(text, variables, context) {
  const contextText = context ? JSON.stringify(context, null, 2) : '{}';
  return `请扮演专业文档助手。你将获得一份 Word 文档的全文和变量列表，请根据文档上下文为每个变量推断合理的填充值。\n\n` +
    `文档全文：\n${text}\n\n` +
    `变量列表：${variables.join(', ')}\n\n` +
    `当前上下文信息（JSON）：\n${contextText}\n\n` +
    `请只输出 JSON 对象，格式如下：{"变量名1":"填充值","变量名2":"填充值"}。不要输出多余说明。`;
}

async function requestBailianFill(text, variables, context) {
  const modelCfg = resolveTemplateModelConfig();
  if (!modelCfg) {
    throw new Error('未配置可用的模型（请在控制台启用一条含 API Key 的 aliyun/openai 配置，或设置环境变量）');
  }
  const url = chatCompletionsUrl(modelCfg);
  const headers = buildHeadersFromConfig(modelCfg.api_key);

  const response = await fetch(url, {
    method: 'POST',
    headers,
    body: JSON.stringify({
      model: normalizeModelId(modelCfg.model_name),
      messages: [
        { role: 'system', content: '你是专业文档助手，擅长根据上下文填写模板变量。' },
        { role: 'user', content: buildFillPrompt(text, variables, context) },
      ],
      temperature: 0.2,
    }),
  });

  if (!response.ok) {
    const errorBody = await response.text();
    logger.error('[Bailian] Error:', errorBody);
    throw new Error(`Bailian request failed: ${response.status}`);
  }

  const data = await response.json();
  const content = data?.choices?.[0]?.message?.content || '';
  return content;
}

async function requestBailianStream(text, variables, context, signal) {
  const modelCfg = resolveTemplateModelConfig();
  if (!modelCfg) {
    throw new Error('未配置可用的模型（请在控制台启用一条含 API Key 的 aliyun/openai 配置，或设置环境变量）');
  }
  const url = chatCompletionsUrl(modelCfg);
  const headers = buildHeadersFromConfig(modelCfg.api_key);
  if (modelCfg.type === 'aliyun') {
    headers['X-DashScope-SSE'] = 'enable';
  }

  const response = await fetch(url, {
    method: 'POST',
    headers,
    body: JSON.stringify({
      model: normalizeModelId(modelCfg.model_name),
      messages: [
        { role: 'system', content: '你是专业文档助手，擅长根据上下文填写模板变量。' },
        { role: 'user', content: buildFillPrompt(text, variables, context) },
      ],
      temperature: 0.2,
      stream: true,
    }),
    signal,
  });

  if (!response.ok || !response.body) {
    const errorBody = await response.text();
    logger.error('[Bailian stream] Error:', errorBody);
    throw new Error(`Bailian request failed: ${response.status}`);
  }

  return response.body.getReader();
}

function parseVariableResponse(text) {
  if (!text) {
    return {};
  }

  const jsonStart = text.indexOf('{');
  const jsonEnd = text.lastIndexOf('}');
  if (jsonStart >= 0 && jsonEnd > jsonStart) {
    const jsonText = text.slice(jsonStart, jsonEnd + 1);
    try {
      return JSON.parse(jsonText);
    } catch (error) {
      console.warn('[Bailian] JSON parse failed', error);
    }
  }

  const result = {};
  text.split(/\r?\n/).forEach((line) => {
    const match = line.match(/^\s*([^:：]+)[:：]\s*(.+)$/);
    if (match) {
      result[match[1].trim()] = match[2].trim();
    }
  });
  return result;
}

function buildFilledDocxBase64(buffer, values) {
  const normalizedBuffer = normalizeDocxTemplateBuffer(buffer);
  const zip = new PizZip(normalizedBuffer);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: '{', end: '}' },
    parser: (tag) => {
      const key = String(tag ?? '').trim();
      return {
        get(scope) {
          if (!key) {
            return undefined;
          }
          if (scope && Object.prototype.hasOwnProperty.call(scope, key)) {
            return scope[key];
          }
          return undefined;
        },
      };
    },
    nullGetter: (part) => `{{${part.value}}}`,
  });

  doc.render(values);
  const output = doc.getZip().generate({ type: 'nodebuffer' });
  return output.toString('base64');
}

app.post('/api/templates/fill-stream', async (req, res) => {
  res.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    Connection: 'keep-alive',
    'X-Accel-Filtering': 'no',
  });
  res.write(':ok\n\n');

  let closed = false;
  let abortController;
  req.on('close', () => {
    closed = true;
    if (abortController) {
      abortController.abort();
    }
  });

  try {
    const { templateId, context } = req.body || {};
    if (!templateId) {
      res.write(`event: error\ndata: 模板 ID 缺失\n\n`);
      return res.end();
    }

    const stmt = db.prepare('SELECT * FROM templates WHERE id = ?');
    const template = stmt.get(templateId);
    if (!template) {
      res.write(`event: error\ndata: 模板不存在\n\n`);
      return res.end();
    }

    const templateOwner = String(template.user_email || '').trim();
    const reqUe = String(req.body?.userEmail || '').trim();
    if (templateOwner && reqUe && templateOwner !== reqUe) {
      res.write(`event: error\ndata: 无权使用该模板\n\n`);
      return res.end();
    }

    const filePath = path.join(uploadsDir, template.filename);
    if (!fs.existsSync(filePath)) {
      res.write(`event: error\ndata: 模板文件缺失\n\n`);
      return res.end();
    }

    const originalBuffer = await fs.promises.readFile(filePath);
    const buffer = normalizeDocxTemplateBuffer(originalBuffer);
    const { value: rawText } = await mammoth.extractRawText({ buffer });
    const variables = extractTemplateVariables(rawText);

    if (variables.length === 0) {
      res.write(`event: done\ndata: {}\n\n`);
      return res.end();
    }

    abortController = new AbortController();
    const reader = await requestBailianStream(rawText, variables, context, abortController.signal);
    const decoder = new TextDecoder('utf-8');
    let streamBuffer = '';
    let fullText = '';

    while (true) {
      const { value, done } = await reader.read();
      if (done) {
        break;
      }

      const chunk = decoder.decode(value, { stream: true });
      logger.info('[RAW AI CHUNK]', chunk.toString());
      streamBuffer += chunk;
      const lines = streamBuffer.split('\n');
      streamBuffer = lines.pop() || '';

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed.startsWith('data:')) {
          continue;
        }

        const payload = trimmed.replace(/^data:\s*/i, '').trim();
        if (!payload || payload === '[DONE]') {
          continue;
        }

        let payloads = [];
        try {
          payloads = parseSsePayloads(payload);
        } catch {
          continue;
        }
        for (const json of payloads) {
          try {
            const delta = json?.choices?.[0]?.delta;
            if (!delta || typeof delta !== 'object') {
              continue;
            }
            // 深度思考模型的 reasoning 不参与模板 JSON 聚合
            if (typeof delta.reasoning_content === 'string' && delta.reasoning_content.length > 0) {
              continue;
            }
            const chunk = typeof delta.content === 'string' ? delta.content : '';
            if (chunk) {
              fullText += chunk;
              for (const char of chunk) {
                res.write(`data: ${JSON.stringify({ type: 'delta', text: char })}\n\n`);
              }
            }
          } catch {
            /* 截断 chunk 等异常跳过 */
          }
        }
      }
    }

    const values = parseVariableResponse(fullText);
    const currentValues = {};

    for (const variable of variables) {
      if (closed) {
        break;
      }

      const value = values?.[variable] || '';
      currentValues[variable] = value;

      let filledDocx = '';
      try {
        filledDocx = buildFilledDocxBase64(buffer, currentValues);
      } catch (error) {
        logger.error('[template fill] docx render failed', error);
      }

      const payload = JSON.stringify({ name: variable, value, filledDocx });
      res.write(`event: variable\ndata: ${payload}\n\n`);
    }

    let recordId = null;
    try {
      const finalDocx = buildFilledDocxBase64(buffer, currentValues);
      recordId = persistTemplateFillRecord({
        templateId,
        userEmail: reqUe,
        templateName: template.original_name,
        variables: currentValues,
        context,
        filledDocxBase64: finalDocx,
      });
    } catch (recordErr) {
      logger.error('[template fill] persist record failed', recordErr);
    }

    res.write(`event: done\ndata: ${JSON.stringify({ recordId })}\n\n`);
    res.end();
  } catch (error) {
    logger.error('[template fill] failed', error);
    res.write(`event: error\ndata: AI 填充失败\n\n`);
    res.end();
  }
});

/** 生产环境：若存在 frontend 构建产物则托管静态页（Docker 将 dist 拷至 ./public） */
const frontendPublic = path.join(__dirname, 'public');
if (fs.existsSync(path.join(frontendPublic, 'index.html'))) {
  app.use(express.static(frontendPublic, { index: false, fallthrough: true }));
  app.get(/.*/, (req, res, next) => {
    if (req.method !== 'GET' || req.path.startsWith('/api')) {
      return next();
    }
    res.sendFile(path.join(frontendPublic, 'index.html'));
  });
}

app.use((err, _req, res, _next) => {
  if (err?.message?.includes('.docx')) {
    return res.status(400).json({ success: false, message: err.message });
  }

  if (err?.code === 'LIMIT_FILE_SIZE') {
    return res.status(400).json({ success: false, message: '文件过大，请上传 20MB 以内的文件' });
  }

  logger.error('[server error]', err);
  res.status(500).json({ success: false, message: '服务器异常' });
});

const server = app.listen(port, () => {
  logger.info(`✅ 会意助手后端服务已启动: http://localhost:${port}`);
  logger.info(`[todos/webhooks] 定时推送已启用（间隔 ${TODO_PUSH_POLL_MS}ms，URL 来自 system_configs；未配置则静默跳过）`);
  setTimeout(() => {
    void runTodoWebhookPush();
  }, 3000);
  setInterval(() => {
    logger.info(`[调度器] setInterval 触发，系统时间：${new Date().toLocaleString('zh-CN', { hour12: false })}`);
    void runTodoWebhookPush();
  }, TODO_PUSH_POLL_MS);
});
attachRealtimeAsrWss(server);
/** 长耗时对话 / 纪要可能较慢，避免 Node 默认 ~2min 套接字超时过早断开 */
server.timeout = 900000;
server.headersTimeout = 910000;
try {
  server.requestTimeout = 900000;
} catch {
  /* 旧版 Node 无 requestTimeout */
}
