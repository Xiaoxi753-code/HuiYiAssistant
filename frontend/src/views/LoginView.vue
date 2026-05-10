<script setup>
import { ref, onMounted } from 'vue'
import { useRouter } from 'vue-router'
import { useUserStore } from '../store/user.js'

const REMEMBER_KEY = 'openclaw_remember_me'
const SAVED_EMAIL_KEY = 'openclaw_saved_login_email'

const router = useRouter()
const userStore = useUserStore()
const currentView = ref('login')

const loginEmail = ref('')
const loginPassword = ref('')
const rememberMe = ref(false)

onMounted(() => {
  rememberMe.value = localStorage.getItem(REMEMBER_KEY) === '1'
  if (rememberMe.value) {
    const saved = localStorage.getItem(SAVED_EMAIL_KEY)
    if (saved) {
      loginEmail.value = saved
    }
  }
})
const regName = ref('')
const regEmail = ref('')
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

async function loginWithCredentials(email, password, persistentSession) {
  const res = await fetch('/api/auth/login', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password }),
  })
  const data = await parseJsonResponse(res)
  if (!res.ok) {
    if (res.status === 401 || res.status === 400) {
      alert(data.message || '登录失败')
    } else {
      alert(data.message || '登录失败，请稍后重试')
    }
    return false
  }
  if (data.user) {
    try {
      localStorage.setItem(REMEMBER_KEY, persistentSession ? '1' : '0')
      if (persistentSession) {
        localStorage.setItem(SAVED_EMAIL_KEY, email.trim())
      } else {
        localStorage.removeItem(SAVED_EMAIL_KEY)
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
  const email = loginEmail.value.trim()
  const password = loginPassword.value
  if (!email || !password) {
    alert('请填写邮箱和密码')
    return
  }
  await loginWithCredentials(email, password, rememberMe.value)
}

async function handleRegisterSubmit() {
  const name = regName.value.trim()
  const email = regEmail.value.trim()
  const password = regPassword.value
  const confirm = regConfirmPassword.value
  if (!name || !email || !password) {
    alert('请填写完整注册信息')
    return
  }
  if (password !== confirm) {
    alert('两次密码不一致')
    return
  }
  const res = await fetch('/api/auth/register', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password, name }),
  })
  const data = await parseJsonResponse(res)
  if (!res.ok) {
    if (res.status === 401 || res.status === 400) {
      alert(data.message || '注册失败')
    } else {
      alert(data.message || '注册失败，请稍后重试')
    }
    return
  }
  await loginWithCredentials(email, password, rememberMe.value)
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
        <div class="logo-icon">OC</div>
        OpenClaw
      </div>
      <h2>欢迎回来</h2>
      <p class="auth-subtitle">登录你的智能办公空间，继续高效工作</p>
      <form id="loginForm" @submit.prevent="handleLogin">
        <div class="form-group">
          <label>邮箱地址</label>
          <input v-model="loginEmail" type="email" id="loginEmail" placeholder="your@email.com" required />
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
          />
          <p class="error-msg" id="loginError">邮箱或密码错误，请重试</p>
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
        <div class="logo-icon">OC</div>
        OpenClaw
      </div>
      <h2>创建账号</h2>
      <p class="auth-subtitle">注册即可解锁全部智能办公功能</p>
      <form id="registerForm" @submit.prevent="handleRegisterSubmit">
        <div class="form-group">
          <label>用户名</label>
          <input v-model="regName" type="text" id="regName" placeholder="你的名字" required minlength="2" />
        </div>
        <div class="form-group">
          <label>邮箱地址</label>
          <input v-model="regEmail" type="email" id="regEmail" placeholder="your@email.com" required />
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

  <div v-show="currentView === 'forgotPassword'" class="auth-container" id="pageForgotPassword">
    <div class="auth-card">
      <div class="auth-logo">
        <div class="logo-icon">OC</div>
        OpenClaw
      </div>
      <h2>找回密码</h2>
      <p class="auth-subtitle">输入邮箱，我们会发送重置链接</p>
      <form id="forgotForm" @submit.prevent="handleForgotSubmit">
        <div class="form-group">
          <label>注册邮箱</label>
          <input type="email" id="forgotEmail" placeholder="your@email.com" required />
        </div>
        <button type="submit" class="btn btn-primary">发送重置链接</button>
      </form>
      <p class="auth-footer">
        <a href="#" @click.prevent="setView('login')">返回登录</a>
      </p>
    </div>
  </div>
</template>
