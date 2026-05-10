<script setup>
import { ref, onMounted } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage, ElMessageBox, ElDialog, ElButton, ElInput } from 'element-plus'
import 'element-plus/es/components/dialog/style/css'
import 'element-plus/es/components/button/style/css'
import 'element-plus/es/components/input/style/css'
import { useUserStore } from '../store/user.js'

const router = useRouter()
const userStore = useUserStore()

const loading = ref(false)
const items = ref([])

const editOpen = ref(false)
const creating = ref(false)
const form = ref({
  id: null,
  name: '',
  description: '',
  system_prompt: '',
  emoji: '✨',
  wake_word: '',
})

async function loadList() {
  const email = userStore.user?.email
  if (!email) {
    items.value = []
    return
  }
  loading.value = true
  try {
    const r = await fetch(`/api/skills?email=${encodeURIComponent(email)}&detail=1`)
    const d = await r.json().catch(() => ({}))
    items.value = d.success && Array.isArray(d.items) ? d.items : []
  } catch (e) {
    console.error(e)
    ElMessage.error('加载技能列表失败')
  } finally {
    loading.value = false
  }
}

function openCreate() {
  creating.value = true
  form.value = {
    id: null,
    name: '',
    description: '',
    system_prompt: '',
    emoji: '✨',
    wake_word: '',
  }
  editOpen.value = true
}

function openEdit(row) {
  creating.value = false
  form.value = {
    id: row.id,
    name: row.name || '',
    description: row.description || '',
    system_prompt: row.system_prompt || '',
    emoji: row.emoji || '✨',
    wake_word: row.wake_word || '',
  }
  editOpen.value = true
}

async function saveForm() {
  const email = userStore.user?.email
  if (!email) {
    ElMessage.warning('请先登录')
    return
  }
  const name = String(form.value.name || '').trim()
  const sp = String(form.value.system_prompt || '').trim()
  if (!name || !sp) {
    ElMessage.warning('请填写名称与 System Prompt')
    return
  }
  try {
    if (creating.value) {
      const r = await fetch('/api/skills', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          email,
          name,
          description: String(form.value.description || '').trim(),
          system_prompt: sp,
          emoji: String(form.value.emoji || '✨').trim(),
          wake_word: String(form.value.wake_word || '').trim(),
        }),
      })
      const d = await r.json().catch(() => ({}))
      if (!r.ok || !d.success) {
        throw new Error(d.message || `HTTP ${r.status}`)
      }
      ElMessage.success('已创建 Skill')
    } else {
      const r = await fetch(`/api/skills/${form.value.id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          email,
          name,
          description: String(form.value.description || '').trim(),
          system_prompt: sp,
          emoji: String(form.value.emoji || '✨').trim(),
          wake_word: String(form.value.wake_word || '').trim(),
        }),
      })
      const d = await r.json().catch(() => ({}))
      if (!r.ok || !d.success) {
        throw new Error(d.message || `HTTP ${r.status}`)
      }
      ElMessage.success('已保存')
    }
    editOpen.value = false
    await loadList()
  } catch (e) {
    ElMessage.error(String(e?.message || e || '保存失败'))
  }
}

async function remove(row) {
  const email = userStore.user?.email
  if (!email) {
    return
  }
  try {
    await ElMessageBox.confirm(`确定删除技能「${row.name}」吗？`, '删除 Skill', {
      type: 'warning',
      confirmButtonText: '删除',
      cancelButtonText: '取消',
    })
  } catch {
    return
  }
  try {
    const r = await fetch(`/api/skills/${row.id}?email=${encodeURIComponent(email)}`, { method: 'DELETE' })
    const d = await r.json().catch(() => ({}))
    if (!r.ok || !d.success) {
      throw new Error(d.message || `HTTP ${r.status}`)
    }
    ElMessage.success('已删除')
    await loadList()
  } catch (e) {
    ElMessage.error(String(e?.message || e || '删除失败'))
  }
}

function runSkill(sk) {
  const w = String(sk.wake_word || '').trim()
  router.push({
    name: 'dashboard',
    query: {
      tab: 'chat',
      skillId: String(sk.id),
      wake: w || '',
    },
  })
}

onMounted(() => {
  void loadList()
})
</script>

<template>
  <div class="main-content skill-library-page">
    <div class="topbar">
      <div>
        <h1 class="page-title">技能广场 · Skill Library</h1>
        <span class="breadcrumb">首页 / 技能广场</span>
      </div>
      <div class="topbar-actions">
        <button class="btn btn-primary btn-sm" :disabled="!userStore.user?.email" @click="openCreate">＋ 新建 Skill</button>
        <button class="btn btn-outline btn-sm" @click="loadList">刷新</button>
      </div>
    </div>

    <p class="skill-page-lead">管理从对话中提炼或手动创建的可复用流程；在对话助手中输入 @ 可快速挂载。</p>

    <div v-if="loading" class="skill-page-loading">加载中…</div>
    <div v-else class="skill-library-grid skill-library-grid--page">
      <div v-for="sk in items" :key="sk.id" class="skill-library-card skill-library-card--page">
        <div class="skill-library-card-icon" aria-hidden="true">{{ sk.emoji || '✨' }}</div>
        <div class="skill-library-card-body">
          <div class="skill-library-card-title">{{ sk.name }}</div>
          <p class="skill-library-card-desc">{{ sk.description || '（无描述）' }}</p>
          <div class="skill-library-card-meta">已使用 {{ sk.use_count ?? 0 }} 次</div>
        </div>
        <div class="skill-library-card-actions">
          <button type="button" class="btn btn-primary btn-xs" @click="runSkill(sk)">带入对话</button>
          <button type="button" class="btn btn-outline btn-xs" @click="openEdit(sk)">编辑</button>
          <button type="button" class="btn btn-ghost btn-xs" @click="remove(sk)">删除</button>
        </div>
      </div>
      <p v-if="!loading && items.length === 0" class="skill-library-empty skill-library-empty--page">
        暂无 Skill。点击「新建 Skill」，或在对话助手内使用「打包为 Skill」。
      </p>
    </div>

    <el-dialog v-model="editOpen" :title="creating ? '新建 Skill' : '编辑 Skill'" width="520px" destroy-on-close align-center>
      <div class="skill-form">
        <label class="mcm-label">名称</label>
        <el-input v-model="form.name" placeholder="简洁命名" />
        <label class="mcm-label">图标 Emoji</label>
        <el-input v-model="form.emoji" maxlength="8" />
        <label class="mcm-label">唤醒词（可选）</label>
        <el-input v-model="form.wake_word" placeholder="如 /周报助手" />
        <label class="mcm-label">描述</label>
        <el-input v-model="form.description" type="textarea" :rows="2" />
        <label class="mcm-label">System Prompt</label>
        <el-input v-model="form.system_prompt" type="textarea" :rows="10" placeholder="注入模型侧的系统指令" />
      </div>
      <template #footer>
        <el-button @click="editOpen = false">取消</el-button>
        <el-button type="primary" @click="saveForm">保存</el-button>
      </template>
    </el-dialog>
  </div>
</template>

<style scoped>
.skill-library-page {
  flex: 1;
  min-width: 0;
  display: flex;
  flex-direction: column;
}
.skill-page-lead {
  color: var(--text-secondary);
  font-size: 0.88rem;
  margin: 0 0 16px;
  line-height: 1.55;
}
.skill-page-loading {
  padding: 24px;
  color: var(--text-muted);
}
.skill-library-grid--page {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
  gap: 14px;
}
.skill-library-card--page {
  display: flex;
  flex-direction: column;
  gap: 10px;
  padding: 16px;
  border-radius: var(--radius-lg);
  border: 1px solid var(--border-light);
  background: var(--surface);
  box-shadow: var(--shadow-sm);
}
.skill-library-empty--page {
  grid-column: 1 / -1;
  text-align: center;
  padding: 40px 16px;
  color: var(--text-muted);
}
.skill-form {
  display: flex;
  flex-direction: column;
  gap: 10px;
}
.mcm-label {
  font-size: 0.75rem;
  font-weight: 600;
  color: var(--text-secondary);
}
</style>
