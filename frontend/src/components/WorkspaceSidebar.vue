<script setup>
import { computed } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { ElMessage, ElIcon } from 'element-plus'
import 'element-plus/es/components/icon/style/css'
import {
  ChatDotRound,
  DArrowLeft,
  DArrowRight,
  DataBoard,
  Document,
  FolderOpened,
  DocumentCopy,
  Link,
  Setting,
  Star,
  Switch,
  SwitchButton,
  TrendCharts,
  Trophy,
} from '@element-plus/icons-vue'
import { useUserStore } from '../store/user.js'

defineProps({
  collapsed: {
    type: Boolean,
    default: false,
  },
})

const emit = defineEmits(['toggle-collapse'])

const route = useRoute()
const router = useRouter()
const userStore = useUserStore()

function goDashboardTab(tab) {
  if (tab === 'chat') {
    router.push({ name: 'dashboard-chat' })
    return
  }
  if (tab === 'companion') {
    router.push({ name: 'dashboard-companion' })
    return
  }
  router.push({ name: 'dashboard', query: { tab } })
}

function isDashTab(tab) {
  if (route.name === 'dashboard-chat') {
    return tab === 'chat'
  }
  if (route.name === 'dashboard-companion') {
    return tab === 'companion'
  }
  if (route.name !== 'dashboard') {
    return false
  }
  const q = route.query.tab
  if (tab === 'overview') {
    return !q || q === 'overview'
  }
  return q === tab
}

function isCompanionWarmNav() {
  return route.name === 'dashboard-companion' || (route.name === 'dashboard' && route.query.tab === 'companion')
}

function goSkills() {
  if (!userStore.user?.email) {
    ElMessage.warning('请先登录后使用技能广场')
    return
  }
  router.push({ name: 'skills' })
}

function goGrowth() {
  router.push({ name: 'growth' })
}

function goAdmin() {
  router.push({ name: 'admin' })
}

const isAdmin = computed(() => String(userStore.user?.role || '') === 'admin')

function handleLogout() {
  userStore.clearUser()
  router.push('/login')
}

</script>

<template>
  <aside class="sidebar" :class="{ 'sidebar--warm': isCompanionWarmNav(), 'sidebar--collapsed': collapsed }">
    <div class="sidebar-header">
      <div class="logo-icon-sm">会</div>
      <span class="logo-text">会意助手</span>
    </div>
    <nav class="sidebar-nav">
      <div class="nav-section">主菜单</div>
      <button type="button" class="nav-item" :class="{ active: isDashTab('overview') }" @click="goDashboardTab('overview')">
        <span class="nav-item-inner">
          <el-icon class="nav-item-icon" :size="18"><DataBoard /></el-icon>
          <span>工作台概览</span>
        </span>
      </button>
      <button type="button" class="nav-item" :class="{ active: isDashTab('chat') }" @click="goDashboardTab('chat')">
        <span class="nav-item-inner">
          <el-icon class="nav-item-icon" :size="18"><ChatDotRound /></el-icon>
          <span>对话助手</span>
          <span class="nav-badge hot">NEW</span>
        </span>
      </button>
      <button type="button" class="nav-item" :class="{ active: route.name === 'skills' }" @click="goSkills">
        <span class="nav-item-inner">
          <el-icon class="nav-item-icon" :size="18"><Trophy /></el-icon>
          <span>技能广场</span>
        </span>
      </button>
      <button type="button" class="nav-item" :class="{ active: isDashTab('templates') }" @click="goDashboardTab('templates')">
        <span class="nav-item-inner">
          <el-icon class="nav-item-icon" :size="18"><DocumentCopy /></el-icon>
          <span>智能模板引擎</span>
          <span class="nav-badge hot">HOT</span>
        </span>
      </button>
      <button type="button" class="nav-item" :class="{ active: isDashTab('library') }" @click="goDashboardTab('library')">
        <span class="nav-item-inner">
          <el-icon class="nav-item-icon" :size="18"><FolderOpened /></el-icon>
          <span>个人资料库</span>
        </span>
      </button>
      <button v-if="false" type="button" class="nav-item" :class="{ active: isDashTab('meeting') }" @click="goDashboardTab('meeting')">
        <span class="nav-item-inner">
          <el-icon class="nav-item-icon" :size="18"><Document /></el-icon>
          <span>智能会议纪要</span>
        </span>
      </button>
      <button type="button" class="nav-item" :class="{ active: isDashTab('paper') }" @click="goDashboardTab('paper')">
        <span class="nav-item-inner">
          <el-icon class="nav-item-icon" :size="18"><Document /></el-icon>
          <span>智能长文速读</span>
        </span>
      </button>
      <button v-if="false" type="button" class="nav-item" :class="{ active: isDashTab('ppt') }" @click="goDashboardTab('ppt')">
        <span class="nav-item-inner">
          <el-icon class="nav-item-icon" :size="18"><Document /></el-icon>
          <span>PPT大纲生成</span>
        </span>
      </button>
      <div class="nav-section">更多</div>
      <button type="button" class="nav-item" :class="{ active: isDashTab('companion') }" @click="goDashboardTab('companion')">
        <span class="nav-item-inner">
          <el-icon class="nav-item-icon" :size="18"><Star /></el-icon>
          <span>情感陪伴对话</span>
        </span>
      </button>
      <button type="button" class="nav-item" :class="{ active: route.name === 'growth' }" @click="goGrowth">
        <span class="nav-item-inner">
          <el-icon class="nav-item-icon" :size="18" aria-hidden="true"><TrendCharts /></el-icon>
          <span>情绪与成长看板</span>
        </span>
      </button>
      <button type="button" class="nav-item" :class="{ active: isDashTab('identity') }" @click="goDashboardTab('identity')">
        <span class="nav-item-inner">
          <el-icon class="nav-item-icon" :size="18"><Switch /></el-icon>
          <span>多身份切换</span>
        </span>
      </button>
      <button type="button" class="nav-item" :class="{ active: isDashTab('reminders') }" @click="goDashboardTab('reminders')">
        <span class="nav-item-inner">
          <span class="nav-item-icon nav-item-icon--emoji" aria-hidden="true">🔔</span>
          <span>主动提醒推送</span>
        </span>
      </button>
      <div class="nav-section">设置</div>
      <button type="button" class="nav-item" :class="{ active: isDashTab('apiConfig') }" @click="goDashboardTab('apiConfig')">
        <span class="nav-item-inner">
          <el-icon class="nav-item-icon" :size="18"><Setting /></el-icon>
          <span>API &amp; 模型配置</span>
        </span>
      </button>
      <button type="button" class="nav-item" :class="{ active: isDashTab('webhookConfig') }" @click="goDashboardTab('webhookConfig')">
        <span class="nav-item-inner">
          <el-icon class="nav-item-icon" :size="18" aria-hidden="true"><Link /></el-icon>
          <span>外联推送配置</span>
        </span>
      </button>
      <button v-if="isAdmin" type="button" class="nav-item" :class="{ active: route.name === 'admin' }" @click="goAdmin">
        <span class="nav-item-inner">
          <el-icon class="nav-item-icon" :size="18"><Setting /></el-icon>
          <span>后台管理</span>
          <span class="nav-badge hot">ADMIN</span>
        </span>
      </button>
    </nav>
    <div class="sidebar-footer">
      <div class="user-avatar" id="sidebarAvatar">{{ (userStore.user?.name || 'U').charAt(0).toUpperCase() }}</div>
      <div class="user-info">
        <div class="user-name" id="sidebarUserName">{{ userStore.user?.name || '用户' }}</div>
        <div class="user-role">个人版</div>
      </div>
      <div class="sidebar-footer-actions">
        <button
          type="button"
          class="sidebar-collapse-btn sidebar-collapse-btn--footer"
          :title="collapsed ? '展开侧栏' : '收起侧栏'"
          :aria-expanded="!collapsed"
          @click="emit('toggle-collapse')"
        >
          <el-icon :size="18">
            <DArrowLeft v-if="!collapsed" />
            <DArrowRight v-else />
          </el-icon>
        </button>
        <button type="button" class="btn-ghost sidebar-logout-btn" title="退出登录" @click="handleLogout">
          <el-icon :size="20"><SwitchButton /></el-icon>
        </button>
      </div>
    </div>
  </aside>
</template>

<style scoped>
.nav-item-inner {
  display: inline-flex;
  align-items: center;
  gap: 10px;
  width: 100%;
}

.nav-item-icon {
  flex-shrink: 0;
  color: inherit;
  opacity: 0.92;
}

.nav-item-icon--emoji {
  font-size: 18px;
  line-height: 1;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 1.25em;
}

.sidebar-logout-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  padding: 6px 8px;
}

.sidebar-footer-actions {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  flex-shrink: 0;
  margin-left: auto;
}

.sidebar-collapse-btn {
  flex-shrink: 0;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 32px;
  height: 32px;
  padding: 0;
  border: 1px solid var(--border-light, #e5e7eb);
  border-radius: var(--radius-sm, 8px);
  background: var(--surface-alt, #f8fafc);
  color: var(--text-secondary, #475569);
  cursor: pointer;
  transition:
    background 0.2s ease,
    border-color 0.2s ease,
    color 0.2s ease,
    transform 0.18s ease,
    box-shadow 0.2s ease;
}

.sidebar-collapse-btn:hover {
  background: var(--primary-light, #eef2ff);
  border-color: #c7d2fe;
  color: var(--primary, #4f46e5);
  transform: translateX(-1px);
  box-shadow: 0 2px 8px rgba(79, 70, 229, 0.12);
}

.sidebar-collapse-btn--footer:hover {
  transform: scale(1.05);
}

.sidebar.sidebar--collapsed .sidebar-footer-actions {
  margin-left: 0;
  flex-direction: column;
  gap: 8px;
}
</style>
