<script setup>
import { computed, onMounted, ref } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage, ElMessageBox } from 'element-plus'
import 'element-plus/es/components/message/style/css'
import 'element-plus/es/components/message-box/style/css'
import { useUserStore } from '../store/user.js'
import { formatLocalDateTime } from '../utils/dateDisplay.js'

const router = useRouter()
const userStore = useUserStore()

const tab = ref('overview')
const loading = ref(false)
const errorText = ref('')

const overview = ref(null)
const users = ref([])
const userQuery = ref('')
const system = ref(null)

const adminName = computed(() => userStore.user?.name || userStore.user?.email || '管理员')

/** 统一带令牌请求；401 说明令牌缺失/过期，403 说明不是管理员 */
async function adminApi(path, options = {}) {
  const res = await fetch(path, {
    ...options,
    headers: {
      ...(options.body ? { 'Content-Type': 'application/json' } : {}),
      ...(userStore.token ? { Authorization: `Bearer ${userStore.token}` } : {}),
      ...(options.headers || {}),
    },
  })
  let data = {}
  try {
    data = await res.json()
  } catch {
    data = {}
  }
  if (res.status === 401) {
    const hint = userStore.token
      ? '登录状态已失效，请重新登录'
      : '当前登录态没有管理令牌，请退出后重新登录'
    ElMessage.error(hint)
    throw new Error(hint)
  }
  if (res.status === 403) {
    ElMessage.error('该账号不是管理员')
    router.replace({ name: 'dashboard' })
    throw new Error('forbidden')
  }
  if (!res.ok || data.success === false) {
    throw new Error(data.message || `HTTP ${res.status}`)
  }
  return data
}

async function loadOverview() {
  const data = await adminApi('/api/admin/overview')
  overview.value = data
}

async function loadUsers() {
  const q = userQuery.value.trim()
  const data = await adminApi(`/api/admin/users${q ? `?q=${encodeURIComponent(q)}` : ''}`)
  users.value = Array.isArray(data.items) ? data.items : []
}

async function loadSystem() {
  const data = await adminApi('/api/admin/system')
  system.value = data
}

async function refresh() {
  loading.value = true
  errorText.value = ''
  try {
    if (tab.value === 'overview') {
      await loadOverview()
    } else if (tab.value === 'users') {
      await loadUsers()
    } else {
      await loadSystem()
    }
  } catch (e) {
    errorText.value = String(e?.message || e || '加载失败')
  } finally {
    loading.value = false
  }
}

const tabs = [
  { key: 'overview', label: '概览' },
  { key: 'users', label: '用户管理' },
  { key: 'system', label: '系统信息' },
]

async function switchTab(key) {
  if (tab.value === key) {
    return
  }
  tab.value = key
  await refresh()
}

function statCards() {
  const c = overview.value?.counts || {}
  return [
    { label: '注册用户', value: c.users, hint: `管理员 ${c.admins || 0} 人` },
    { label: '对话会话', value: c.sessions, hint: `消息 ${c.messages || 0} 条` },
    { label: '待办事项', value: c.todos, hint: '含已完成' },
    { label: '模板', value: c.templates, hint: '全部用户合计' },
    { label: '资料库条目', value: c.knowledge, hint: '文档 / 纪要' },
    { label: 'Skill', value: c.skills, hint: '可复用流程' },
  ]
}

async function changeRole(row, role) {
  if (row.role === role) {
    return
  }
  const label = role === 'admin' ? '设为管理员' : '取消管理员'
  try {
    await ElMessageBox.confirm(`确定对「${row.email}」执行${label}？`, '角色变更', {
      type: 'warning',
      confirmButtonText: '确定',
      cancelButtonText: '取消',
    })
  } catch {
    return
  }
  try {
    await adminApi(`/api/admin/users/${row.id}/role`, {
      method: 'PATCH',
      body: JSON.stringify({ role }),
    })
    ElMessage.success('角色已更新')
    await loadUsers()
  } catch (e) {
    ElMessage.error(String(e?.message || e || '更新失败'))
  }
}

async function resetPassword(row) {
  let value = ''
  try {
    const result = await ElMessageBox.prompt(`为「${row.email}」设置新密码（至少 6 位）`, '重置密码', {
      confirmButtonText: '重置',
      cancelButtonText: '取消',
      inputType: 'text',
      inputPlaceholder: '输入新密码',
      inputValidator: (v) => (String(v || '').trim().length >= 6 ? true : '至少 6 位'),
    })
    value = String(result?.value || '').trim()
  } catch {
    return
  }
  if (!value) {
    return
  }
  try {
    await adminApi(`/api/admin/users/${row.id}/password`, {
      method: 'POST',
      body: JSON.stringify({ newPassword: value }),
    })
    ElMessage.success('密码已重置，该用户的旧登录态已失效')
  } catch (e) {
    ElMessage.error(String(e?.message || e || '重置失败'))
  }
}

async function removeUser(row) {
  let typed = ''
  try {
    const result = await ElMessageBox.prompt(
      `将删除「${row.email}」的账号及其全部会话、待办、模板、资料与技能，不可恢复。\n请输入该账号以确认：`,
      '删除用户',
      {
        type: 'error',
        confirmButtonText: '确认删除',
        cancelButtonText: '取消',
        inputPlaceholder: row.email,
      },
    )
    typed = String(result?.value || '').trim()
  } catch {
    return
  }
  if (typed !== row.email) {
    ElMessage.warning('账号输入不一致，已取消')
    return
  }
  try {
    await adminApi(`/api/admin/users/${row.id}`, { method: 'DELETE' })
    ElMessage.success('用户已删除')
    await loadUsers()
  } catch (e) {
    ElMessage.error(String(e?.message || e || '删除失败'))
  }
}

function formatBytes(n) {
  const v = Number(n || 0)
  if (!v) {
    return '—'
  }
  if (v < 1024) {
    return `${v} B`
  }
  if (v < 1024 * 1024) {
    return `${(v / 1024).toFixed(1)} KB`
  }
  return `${(v / 1024 / 1024).toFixed(2)} MB`
}

function formatUptime(seconds) {
  const s = Number(seconds || 0)
  const d = Math.floor(s / 86400)
  const h = Math.floor((s % 86400) / 3600)
  const m = Math.floor((s % 3600) / 60)
  if (d > 0) {
    return `${d} 天 ${h} 小时`
  }
  if (h > 0) {
    return `${h} 小时 ${m} 分`
  }
  return `${m} 分`
}

function goWorkspace() {
  router.push({ name: 'dashboard' })
}

function logout() {
  userStore.clearUser()
  router.push({ name: 'login' })
}

onMounted(() => {
  void refresh()
})
</script>

<template>
  <div class="admin-page">
    <header class="admin-header">
      <div class="admin-brand">
        <span class="admin-brand-mark">会</span>
        <span class="admin-brand-text">后台管理</span>
      </div>
      <nav class="admin-tabs" aria-label="后台导航">
        <button
          v-for="t in tabs"
          :key="t.key"
          type="button"
          class="admin-tab"
          :class="{ 'admin-tab--active': tab === t.key }"
          @click="switchTab(t.key)"
        >
          {{ t.label }}
        </button>
      </nav>
      <div class="admin-account">
        <span class="admin-account-name">{{ adminName }}</span>
        <span class="admin-account-badge">管理员</span>
        <button type="button" class="admin-link" @click="goWorkspace">返回工作台</button>
        <button type="button" class="admin-link admin-link--danger" @click="logout">退出</button>
      </div>
    </header>

    <main class="admin-main">
      <div class="admin-toolbar">
        <h1 class="admin-title">{{ tabs.find((t) => t.key === tab)?.label }}</h1>
        <div class="admin-toolbar-actions">
          <input
            v-if="tab === 'users'"
            v-model="userQuery"
            class="admin-search"
            type="search"
            placeholder="搜索账号或用户名"
            @keydown.enter.prevent="refresh"
          />
          <button v-if="tab === 'users'" type="button" class="btn btn-outline btn-sm" @click="refresh">搜索</button>
          <button type="button" class="btn btn-outline btn-sm" :disabled="loading" @click="refresh">
            {{ loading ? '刷新中…' : '刷新' }}
          </button>
        </div>
      </div>

      <p v-if="errorText" class="admin-error">{{ errorText }}</p>

      <!-- 概览 -->
      <section v-if="tab === 'overview'">
        <div class="admin-stats">
          <div v-for="c in statCards()" :key="c.label" class="admin-stat">
            <div class="admin-stat-value">{{ c.value ?? '—' }}</div>
            <div class="admin-stat-label">{{ c.label }}</div>
            <div class="admin-stat-hint">{{ c.hint }}</div>
          </div>
        </div>

        <div class="admin-split">
          <div class="admin-card">
            <h2 class="admin-card-title">最近注册</h2>
            <table class="admin-table">
              <thead>
                <tr>
                  <th>账号</th>
                  <th>用户名</th>
                  <th>角色</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="u in overview?.recentUsers || []" :key="u.id">
                  <td class="admin-mono">{{ u.email }}</td>
                  <td>{{ u.name || '—' }}</td>
                  <td>
                    <span class="admin-chip" :class="u.role === 'admin' ? 'admin-chip--admin' : ''">
                      {{ u.role === 'admin' ? '管理员' : '普通用户' }}
                    </span>
                  </td>
                </tr>
                <tr v-if="!(overview?.recentUsers || []).length">
                  <td colspan="3" class="admin-empty">暂无数据</td>
                </tr>
              </tbody>
            </table>
          </div>

          <div class="admin-card">
            <h2 class="admin-card-title">近 7 天</h2>
            <div class="admin-mini">
              <div class="admin-mini-item">
                <span class="admin-mini-value">{{ overview?.activeUsers7d ?? '—' }}</span>
                <span class="admin-mini-label">活跃用户</span>
              </div>
              <div class="admin-mini-item">
                <span class="admin-mini-value">{{ overview?.messages7d ?? '—' }}</span>
                <span class="admin-mini-label">对话消息</span>
              </div>
            </div>
            <h2 class="admin-card-title" style="margin-top: 20px">最近会话</h2>
            <ul class="admin-list">
              <li v-for="s in overview?.recentSessions || []" :key="s.id" class="admin-list-item">
                <span class="admin-list-title">{{ s.title || '未命名会话' }}</span>
                <span class="admin-list-meta">{{ s.userEmail }} · {{ s.messageCount }} 条 · {{ formatLocalDateTime(s.updatedAt) }}</span>
              </li>
              <li v-if="!(overview?.recentSessions || []).length" class="admin-empty">暂无会话</li>
            </ul>
          </div>
        </div>
      </section>

      <!-- 用户管理 -->
      <section v-if="tab === 'users'" class="admin-card">
        <table class="admin-table admin-table--wide">
          <thead>
            <tr>
              <th>#</th>
              <th>账号</th>
              <th>用户名</th>
              <th>角色</th>
              <th>会话</th>
              <th>待办</th>
              <th>模板</th>
              <th>资料</th>
              <th>技能</th>
              <th>最近活跃</th>
              <th>操作</th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="row in users" :key="row.id">
              <td class="admin-muted">{{ row.id }}</td>
              <td class="admin-mono">{{ row.email }}</td>
              <td>{{ row.name || '—' }}</td>
              <td>
                <span class="admin-chip" :class="row.role === 'admin' ? 'admin-chip--admin' : ''">
                  {{ row.role === 'admin' ? '管理员' : '普通用户' }}
                </span>
              </td>
              <td>{{ row.sessionCount }}</td>
              <td>{{ row.todoCount }}</td>
              <td>{{ row.templateCount }}</td>
              <td>{{ row.libraryCount }}</td>
              <td>{{ row.skillCount }}</td>
              <td class="admin-muted">{{ row.lastActiveAt ? formatLocalDateTime(row.lastActiveAt) : '—' }}</td>
              <td>
                <div class="admin-row-actions">
                  <button
                    v-if="row.role !== 'admin'"
                    type="button"
                    class="admin-action"
                    @click="changeRole(row, 'admin')"
                  >
                    设为管理员
                  </button>
                  <button v-else type="button" class="admin-action" @click="changeRole(row, 'user')">取消管理员</button>
                  <button type="button" class="admin-action" @click="resetPassword(row)">重置密码</button>
                  <button type="button" class="admin-action admin-action--danger" @click="removeUser(row)">删除</button>
                </div>
              </td>
            </tr>
            <tr v-if="!users.length">
              <td colspan="11" class="admin-empty">{{ loading ? '加载中…' : '没有匹配的用户' }}</td>
            </tr>
          </tbody>
        </table>
      </section>

      <!-- 系统信息 -->
      <section v-if="tab === 'system'" class="admin-split">
        <div class="admin-card">
          <h2 class="admin-card-title">模型配置</h2>
          <table class="admin-table">
            <thead>
              <tr>
                <th>名称</th>
                <th>模型</th>
                <th>Key</th>
                <th>优先级</th>
                <th>状态</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="m in system?.models || []" :key="m.id">
                <td>{{ m.name || '—' }}</td>
                <td class="admin-mono">{{ m.modelName }}</td>
                <td class="admin-mono">{{ m.apiKey || '—' }}</td>
                <td>{{ m.priority }}</td>
                <td>
                  <span class="admin-chip" :class="m.enabled ? 'admin-chip--ok' : ''">
                    {{ m.enabled ? '启用' : '停用' }}
                  </span>
                </td>
              </tr>
              <tr v-if="!(system?.models || []).length">
                <td colspan="5" class="admin-empty">还没有配置模型</td>
              </tr>
            </tbody>
          </table>

          <h2 class="admin-card-title" style="margin-top: 22px">外联推送</h2>
          <ul class="admin-kv">
            <li><span>企业微信 Webhook</span><b>{{ system?.webhooks?.wecom ? '已配置' : '未配置' }}</b></li>
            <li><span>飞书 Webhook</span><b>{{ system?.webhooks?.feishu ? '已配置' : '未配置' }}</b></li>
            <li><span>轮询间隔</span><b>{{ system?.webhooks?.pollMs ? `${system.webhooks.pollMs} ms` : '—' }}</b></li>
          </ul>
        </div>

        <div class="admin-card">
          <h2 class="admin-card-title">运行环境</h2>
          <ul class="admin-kv">
            <li><span>Node</span><b>{{ system?.runtime?.nodeVersion || '—' }}</b></li>
            <li><span>平台</span><b>{{ system?.runtime?.platform || '—' }}</b></li>
            <li><span>已运行</span><b>{{ formatUptime(system?.runtime?.uptimeSeconds) }}</b></li>
            <li><span>数据库大小</span><b>{{ formatBytes(system?.runtime?.dbSizeBytes) }}</b></li>
            <li><span>日志级别</span><b>{{ system?.runtime?.logLevel || '—' }}</b></li>
          </ul>
          <h2 class="admin-card-title" style="margin-top: 22px">路径</h2>
          <p class="admin-path">{{ system?.runtime?.dbPath || '—' }}</p>
          <p class="admin-path">{{ system?.runtime?.dataDir || '—' }}</p>
        </div>
      </section>
    </main>
  </div>
</template>

<style scoped>
.admin-page {
  min-height: 100vh;
  background: var(--bg, #f5f3f0);
  color: var(--text, #1a1a1a);
}

.admin-header {
  display: flex;
  align-items: center;
  gap: 20px;
  padding: 0 24px;
  height: 60px;
  background: var(--surface, #fff);
  border-bottom: 1px solid var(--border, #e5e1db);
  position: sticky;
  top: 0;
  z-index: 10;
}

.admin-brand {
  display: inline-flex;
  align-items: center;
  gap: 10px;
  flex-shrink: 0;
}

.admin-brand-mark {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 30px;
  height: 30px;
  border-radius: 9px;
  background: var(--primary, #4a6cf7);
  color: #fff;
  font-size: 15px;
  font-weight: 700;
}

.admin-brand-text {
  font-size: 0.95rem;
  font-weight: 700;
}

.admin-tabs {
  display: inline-flex;
  gap: 4px;
  flex: 1;
}

.admin-tab {
  padding: 7px 14px;
  border: none;
  border-radius: 9px;
  background: transparent;
  color: var(--text-secondary, #6b6560);
  font-size: 0.86rem;
  font-family: inherit;
  cursor: pointer;
}

.admin-tab:hover {
  background: var(--surface-alt, #faf9f7);
}

.admin-tab--active {
  background: var(--primary-light, #eef1fe);
  color: var(--primary, #4a6cf7);
  font-weight: 600;
}

.admin-account {
  display: inline-flex;
  align-items: center;
  gap: 10px;
  flex-shrink: 0;
  font-size: 0.82rem;
}

.admin-account-name {
  font-weight: 600;
}

.admin-account-badge {
  padding: 3px 8px;
  border-radius: 999px;
  background: var(--primary-light, #eef1fe);
  color: var(--primary, #4a6cf7);
  font-size: 0.72rem;
  font-weight: 600;
}

.admin-link {
  border: none;
  background: none;
  color: var(--text-secondary, #6b6560);
  font-size: 0.82rem;
  font-family: inherit;
  cursor: pointer;
  padding: 4px 2px;
}

.admin-link:hover {
  color: var(--primary, #4a6cf7);
}

.admin-link--danger:hover {
  color: var(--danger, #ef4444);
}

.admin-main {
  max-width: 1180px;
  margin: 0 auto;
  padding: 24px;
}

.admin-toolbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  margin-bottom: 18px;
  flex-wrap: wrap;
}

.admin-title {
  font-size: 1.15rem;
  font-weight: 700;
  margin: 0;
}

.admin-toolbar-actions {
  display: inline-flex;
  align-items: center;
  gap: 10px;
}

.admin-search {
  height: 34px;
  min-width: 200px;
  padding: 0 12px;
  border: 1px solid var(--border, #e5e1db);
  border-radius: 9px;
  background: var(--surface, #fff);
  font-family: inherit;
  font-size: 0.85rem;
  color: inherit;
}

.admin-search:focus {
  outline: none;
  border-color: var(--primary, #4a6cf7);
}

.admin-error {
  margin: 0 0 14px;
  padding: 10px 14px;
  border-radius: 10px;
  background: var(--danger-light, #fef2f2);
  color: var(--danger, #ef4444);
  font-size: 0.84rem;
}

.admin-stats {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
  gap: 14px;
  margin-bottom: 20px;
}

.admin-stat {
  padding: 16px 18px;
  border-radius: 14px;
  border: 1px solid var(--border, #e5e1db);
  background: var(--surface, #fff);
}

.admin-stat-value {
  font-size: 1.7rem;
  font-weight: 700;
  line-height: 1.15;
}

.admin-stat-label {
  margin-top: 4px;
  font-size: 0.84rem;
  font-weight: 600;
  color: var(--text, #1a1a1a);
}

.admin-stat-hint {
  margin-top: 2px;
  font-size: 0.76rem;
  color: var(--text-muted, #9a9490);
}

.admin-split {
  display: grid;
  grid-template-columns: minmax(0, 1.25fr) minmax(0, 1fr);
  gap: 18px;
  align-items: start;
}

.admin-card {
  padding: 18px 20px;
  border-radius: 14px;
  border: 1px solid var(--border, #e5e1db);
  background: var(--surface, #fff);
  min-width: 0;
}

.admin-card-title {
  margin: 0 0 12px;
  font-size: 0.92rem;
  font-weight: 700;
}

.admin-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 0.83rem;
}

.admin-table th {
  text-align: left;
  font-weight: 600;
  color: var(--text-muted, #9a9490);
  padding: 8px 10px;
  border-bottom: 1px solid var(--border-light, #f0ede8);
  white-space: nowrap;
}

.admin-table td {
  padding: 10px;
  border-bottom: 1px solid var(--border-light, #f0ede8);
  vertical-align: middle;
}

.admin-table tbody tr:hover {
  background: var(--surface-alt, #faf9f7);
}

.admin-table--wide {
  min-width: 900px;
}

.admin-mono {
  font-family: var(--font-mono, ui-monospace, Menlo, Consolas, monospace);
  font-size: 0.8rem;
}

.admin-muted {
  color: var(--text-muted, #9a9490);
}

.admin-empty {
  text-align: center;
  color: var(--text-muted, #9a9490);
  padding: 22px 10px;
}

.admin-chip {
  display: inline-block;
  padding: 3px 9px;
  border-radius: 999px;
  background: var(--surface-alt, #faf9f7);
  border: 1px solid var(--border, #e5e1db);
  font-size: 0.74rem;
  color: var(--text-secondary, #6b6560);
  white-space: nowrap;
}

.admin-chip--admin {
  background: var(--primary-light, #eef1fe);
  border-color: #c7d2fe;
  color: var(--primary, #4a6cf7);
  font-weight: 600;
}

.admin-chip--ok {
  background: var(--success-light, #ecfdf5);
  border-color: #a7f3d0;
  color: var(--success, #10b981);
}

.admin-row-actions {
  display: inline-flex;
  gap: 6px;
  flex-wrap: wrap;
}

.admin-action {
  border: 1px solid var(--border, #e5e1db);
  background: var(--surface, #fff);
  border-radius: 8px;
  padding: 4px 9px;
  font-size: 0.76rem;
  font-family: inherit;
  color: var(--text-secondary, #6b6560);
  cursor: pointer;
  white-space: nowrap;
}

.admin-action:hover {
  border-color: var(--primary, #4a6cf7);
  color: var(--primary, #4a6cf7);
}

.admin-action--danger:hover {
  border-color: var(--danger, #ef4444);
  color: var(--danger, #ef4444);
}

.admin-mini {
  display: flex;
  gap: 28px;
}

.admin-mini-item {
  display: flex;
  flex-direction: column;
  gap: 2px;
}

.admin-mini-value {
  font-size: 1.4rem;
  font-weight: 700;
}

.admin-mini-label {
  font-size: 0.78rem;
  color: var(--text-muted, #9a9490);
}

.admin-list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.admin-list-item {
  display: flex;
  flex-direction: column;
  gap: 2px;
  padding-bottom: 8px;
  border-bottom: 1px dashed var(--border-light, #f0ede8);
}

.admin-list-title {
  font-size: 0.85rem;
  font-weight: 500;
}

.admin-list-meta {
  font-size: 0.75rem;
  color: var(--text-muted, #9a9490);
}

.admin-kv {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.admin-kv li {
  display: flex;
  justify-content: space-between;
  gap: 12px;
  font-size: 0.83rem;
  color: var(--text-secondary, #6b6560);
}

.admin-kv b {
  color: var(--text, #1a1a1a);
  font-weight: 600;
}

.admin-path {
  margin: 0 0 6px;
  font-family: var(--font-mono, ui-monospace, Menlo, Consolas, monospace);
  font-size: 0.75rem;
  color: var(--text-muted, #9a9490);
  overflow-wrap: anywhere;
}

@media (max-width: 900px) {
  .admin-split {
    grid-template-columns: minmax(0, 1fr);
  }
  .admin-tabs {
    display: none;
  }
  .admin-main {
    padding: 16px;
  }
  .admin-card {
    overflow-x: auto;
  }
}
</style>
