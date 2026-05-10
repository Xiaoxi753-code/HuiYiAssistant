<script setup>
import { useRoute, useRouter } from 'vue-router'
import { ElMessage, ElMessageBox } from 'element-plus'
import { useUserStore } from '../store/user.js'

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

function handleLogout() {
  userStore.clearUser()
  router.push('/login')
}

async function confirmDemoPresentationReset() {
  const email = userStore.user?.email
  if (!email) {
    ElMessage.warning('请先登录后再使用演示重置')
    return
  }
  try {
    await ElMessageBox.confirm(
      '将清空当前对话与服务器上的「眼中的你」画像缓存（Skill 库保留）。适合答辩反复演示。',
      '演示重置',
      { confirmButtonText: '重置', cancelButtonText: '取消', type: 'warning' },
    )
  } catch {
    return
  }
  try {
    const r = await fetch('/api/user/demo-reset', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email }),
    })
    let data = {}
    try {
      data = await r.json()
    } catch {
      data = {}
    }
    if (!r.ok || !data.success) {
      throw new Error(data.message || `HTTP ${r.status}`)
    }
    ElMessage.success('演示状态已重置')
    void router.push({ name: 'dashboard', query: { tab: 'overview' } })
  } catch (e) {
    ElMessage.error(String(e?.message || e || '重置失败'))
  }
}
</script>

<template>
  <aside class="sidebar" :class="{ 'sidebar--warm': isCompanionWarmNav() }">
    <div class="sidebar-header">
      <div class="logo-icon-sm">OC</div>
      <span class="logo-text">OpenClaw</span>
    </div>
    <nav class="sidebar-nav">
      <div class="nav-section">主菜单</div>
      <button type="button" class="nav-item" :class="{ active: isDashTab('overview') }" @click="goDashboardTab('overview')">
        <span class="nav-item-icon-text">&#9681;</span> <span>工作台概览</span>
      </button>
      <button type="button" class="nav-item" :class="{ active: isDashTab('chat') }" @click="goDashboardTab('chat')">
        <span class="nav-item-icon-text">&#9654;</span> <span>对话助手</span><span class="nav-badge hot">NEW</span>
      </button>
      <button type="button" class="nav-item" :class="{ active: route.name === 'skills' }" @click="goSkills">
        <span class="nav-item-icon-text">&#10022;</span> <span>技能广场</span>
      </button>
      <button type="button" class="nav-item" :class="{ active: isDashTab('templates') }" @click="goDashboardTab('templates')">
        <span class="nav-item-icon-text">&#9638;</span> <span>智能模板引擎</span><span class="nav-badge hot">HOT</span>
      </button>
      <button type="button" class="nav-item" :class="{ active: isDashTab('library') }" @click="goDashboardTab('library')">
        <span class="nav-item-icon-text">&#128193;</span> <span>个人资料库</span>
      </button>
      <button type="button" class="nav-item" :class="{ active: isDashTab('meeting') }" @click="goDashboardTab('meeting')">
        <span class="nav-item-icon-text">&#9679;</span> <span>智能会议纪要</span>
      </button>
      <button type="button" class="nav-item" :class="{ active: isDashTab('paper') }" @click="goDashboardTab('paper')">
        <span class="nav-item-icon-text">&#9633;</span> <span>文献论文摘要</span>
      </button>
      <button type="button" class="nav-item" :class="{ active: isDashTab('ppt') }" @click="goDashboardTab('ppt')">
        <span class="nav-item-icon-text">&#9650;</span> <span>PPT大纲生成</span>
      </button>
      <div class="nav-section">更多</div>
      <button type="button" class="nav-item" :class="{ active: isDashTab('companion') }" @click="goDashboardTab('companion')">
        <span class="nav-item-icon-text">&#9825;</span> <span>情感陪伴对话</span>
      </button>
      <button type="button" class="nav-item" :class="{ active: route.name === 'growth' }" @click="goGrowth">
        <span class="nav-item-icon-text" aria-hidden="true">&#128202;</span> <span>情绪与成长看板</span>
      </button>
      <button type="button" class="nav-item" :class="{ active: isDashTab('identity') }" @click="goDashboardTab('identity')">
        <span class="nav-item-icon-text">&#8644;</span> <span>多身份切换</span>
      </button>
      <button type="button" class="nav-item" :class="{ active: isDashTab('reminders') }" @click="goDashboardTab('reminders')">
        <span class="nav-item-icon-text">&#9200;</span> <span>主动提醒推送</span>
      </button>
      <div class="nav-section">设置</div>
      <button type="button" class="nav-item" :class="{ active: isDashTab('apiConfig') }" @click="goDashboardTab('apiConfig')">
        <span class="nav-item-icon-text">&#9881;</span> <span>API &amp; 模型配置</span>
      </button>
    </nav>
    <div class="sidebar-footer">
      <div class="user-avatar" id="sidebarAvatar">{{ (userStore.user?.name || 'U').charAt(0).toUpperCase() }}</div>
      <div class="user-info">
        <div class="user-name" id="sidebarUserName">{{ userStore.user?.name || '用户' }}</div>
        <div class="user-role">个人版</div>
        <button
          v-if="userStore.user?.email"
          type="button"
          class="demo-reset-link"
          title="答辩演示：清空对话并重置画像（Skill 保留）"
          @click="confirmDemoPresentationReset"
        >
          演示重置
        </button>
      </div>
      <button type="button" class="btn-ghost" title="退出登录" @click="handleLogout" style="font-size: 1.1rem; font-weight: 700">
        &#10149;
      </button>
    </div>
  </aside>
</template>
