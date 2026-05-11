<script setup>
import { ref, onMounted } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage } from 'element-plus'
import 'element-plus/es/components/message/style/css'
import { useUserStore } from '../store/user.js'

const REMEMBER_KEY = 'openclaw_remember_me'
const SAVED_ACCOUNT_KEY = 'openclaw_saved_login_account'
/** 兼容旧版：仅用于迁移本地记住的登录名 */
const LEGACY_SAVED_EMAIL_KEY = 'openclaw_saved_login_email'

const router = useRouter()
const userStore = useUserStore()
const currentView = ref('login')

const loginAccount = ref('')
const loginPassword = ref('')
const rememberMe = ref(false)

onMounted(() => {
  rememberMe.value = localStorage.getItem(REMEMBER_KEY) === '1'
  if (rememberMe.value) {
    let saved = localStorage.getItem(SAVED_ACCOUNT_KEY)
    if (!saved) {
      saved = localStorage.getItem(LEGACY_SAVED_EMAIL_KEY)
      if (saved) {
        try {
          localStorage.setItem(SAVED_ACCOUNT_KEY, saved)
          localStorage.removeItem(LEGACY_SAVED_EMAIL_KEY)
        } catch {
          /* ignore */
        }
      }
    }
    if (saved) {
      loginAccount.value = saved
    }
  }
})
const regName = ref('')
const regAccount = ref('')
const regPassword = ref('')
const regConfirmPassword = ref('')

function setView(view) {
  currentView.value = view
}

async function parseJsonResponse(res) {
  try {
    return await res.json()
  } catch {
    return {}
  }
}

async function loginWithCredentials(account, password, persistentSession) {
  const res = await fetch('/api/auth/login', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ account, password }),
  })
  const data = await parseJsonResponse(res)
  if (!res.ok) {
    if (res.status === 404 && data.error === 'ACCOUNT_NOT_FOUND') {
      ElMessage.warning('该账号暂未注册，请先注册')
      setView('register')
      regAccount.value = String(account || '').trim()
      return false
    }
    if (res.status === 401 || res.status === 400) {
      ElMessage.error(data.message || '登录失败')
    } else {
      ElMessage.error(data.message || '登录失败，请稍后重试')
    }
    return false
  }
  if (data.user) {
    try {
      localStorage.setItem(REMEMBER_KEY, persistentSession ? '1' : '0')
      const trimmed = String(account || '').trim()
      if (persistentSession) {
        localStorage.setItem(SAVED_ACCOUNT_KEY, trimmed)
      } else {
        localStorage.removeItem(SAVED_ACCOUNT_KEY)
      }
    } catch {
      /* ignore */
    }
    userStore.setUser(data.user, { persistent: persistentSession })
  }
  router.push('/dashboard')
  return true
}

async function handleLogin() {
  const account = loginAccount.value.trim()
  const password = loginPassword.value
  if (!account || !password) {
    ElMessage.warning('请填写账号和密码')
    return
  }
  await loginWithCredentials(account, password, rememberMe.value)
}

async function handleRegisterSubmit() {
  const name = regName.value.trim()
  const account = regAccount.value.trim()
  const password = regPassword.value
  const confirm = regConfirmPassword.value
  if (!name || !account || !password) {
    ElMessage.warning('请填写完整注册信息')
    return
  }
  if (password !== confirm) {
    ElMessage.warning('两次密码不一致')
    return
  }
  const res = await fetch('/api/auth/register', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ account, password, name }),
  })
  const data = await parseJsonResponse(res)
  if (!res.ok) {
    if (res.status === 401 || res.status === 400) {
      ElMessage.error(data.message || '注册失败')
    } else {
      ElMessage.error(data.message || '注册失败，请稍后重试')
    }
    return
  }
  await loginWithCredentials(account, password, rememberMe.value)
}

function handleForgotSubmit() {
  setView('login')
}
</script>

<template>
  <div class="bg-decor shape-1"></div>
  <div class="bg-decor shape-2"></div>
  <div class="bg-decor shape-3"></div>

  <div v-show="currentView === 'login'" class="auth-container" id="pageLogin">
    <div class="auth-card">
      <div class="auth-logo">
        <div class="logo-icon">会</div>
        会意助手
      </div>
      <h2>欢迎回来</h2>
      <p class="auth-subtitle">登录你的智能办公空间，继续高效工作</p>
      <form id="loginForm" @submit.prevent="handleLogin">
        <div class="form-group">
          <label>账号</label>
          <input
            v-model="loginAccount"
            type="text"
            id="loginAccount"
            placeholder="字母或数字等，任意账号名"
            required
            autocomplete="username"
          />
        </div>
        <div class="form-group">
          <label>密码</label>
          <input
            v-model="loginPassword"
            type="password"
            id="loginPassword"
            placeholder="输入密码"
            required
            minlength="6"
            autocomplete="current-password"
          />
          <p class="error-msg" id="loginError">账号或密码错误，请重试</p>
        </div>
        <div class="form-row">
          <label style="display: flex; align-items: center; gap: 6px; cursor: pointer; font-weight: 400">
            <input v-model="rememberMe" type="checkbox" style="accent-color: var(--primary)" />
            记住我
          </label>
          <a href="#" @click.prevent="setView('forgotPassword')">忘记密码？</a>
        </div>
        <button type="submit" class="btn btn-primary">登录</button>
      </form>
      <p class="auth-footer">
        还没有账号？<a href="#" @click.prevent="setView('register')">立即注册</a>
      </p>
    </div>
  </div>

  <div v-show="currentView === 'register'" class="auth-container" id="pageRegister">
    <div class="auth-card">
      <div class="auth-logo">
        <div class="logo-icon">会</div>
        会意助手
      </div>
      <h2>创建账号</h2>
      <p class="auth-subtitle">注册即可解锁全部智能办公功能</p>
      <form id="registerForm" @submit.prevent="handleRegisterSubmit">
        <div class="form-group">
          <label>用户名</label>
          <input v-model="regName" type="text" id="regName" placeholder="你的名字" required minlength="2" />
        </div>
        <div class="form-group">
          <label>账号</label>
          <input
            v-model="regAccount"
            type="text"
            id="regAccount"
            placeholder="字母、数字等，用于登录"
            required
            autocomplete="username"
          />
        </div>
        <div class="form-group">
          <label>密码</label>
          <input v-model="regPassword" type="password" id="regPassword" placeholder="至少6位密码" required minlength="6" />
        </div>
        <div class="form-group">
          <label>确认密码</label>
          <input
            v-model="regConfirmPassword"
            type="password"
            id="regConfirmPassword"
            placeholder="再次输入密码"
            required
            minlength="6"
          />
          <p class="error-msg" id="regError">两次密码不一致</p>
        </div>
        <button type="submit" class="btn btn-primary">注册</button>
      </form>
      <p class="auth-footer">已有账号？<a href="#" @click.prevent="setView('login')">去登录</a></p>
    </div>
  </div>

  <div v-show="currentView === 'forgotPassword'" class="auth-container" id="pageForgot">
    <div class="auth-card">
      <div class="auth-logo">
        <div class="logo-icon">会</div>
        会意助手
      </div>
      <h2>忘记密码</h2>
      <p class="auth-subtitle">请联系管理员或通过注册邮箱找回（演示环境请直接注册新账号）。</p>
      <button type="button" class="btn btn-primary" @click="handleForgotSubmit">返回登录</button>
    </div>
  </div>
</template>
