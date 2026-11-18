import { defineStore } from 'pinia'

const LS_USER = 'openclaw_user'
const SS_USER = 'openclaw_user_session'
const LS_TOKEN = 'openclaw_auth_token'
const SS_TOKEN = 'openclaw_auth_token_session'

function readStoredUser() {
  try {
    const sess = sessionStorage.getItem(SS_USER)
    if (sess != null && sess !== '') {
      return JSON.parse(sess)
    }
    const raw = localStorage.getItem(LS_USER)
    if (raw == null || raw === '') {
      return null
    }
    return JSON.parse(raw)
  } catch {
    return null
  }
}

function readStoredToken() {
  try {
    const sess = sessionStorage.getItem(SS_TOKEN)
    if (sess != null && sess !== '') {
      return sess
    }
    return localStorage.getItem(LS_TOKEN) || ''
  } catch {
    return ''
  }
}

export const useUserStore = defineStore('user', {
  state: () => ({
    user: readStoredUser(),
    /** 服务端下发的会话令牌；管理接口用它鉴权，和业务 email 参数分开存放 */
    token: readStoredToken(),
  }),
  getters: {
    isAdmin: (state) => String(state.user?.role || '') === 'admin',
  },
  actions: {
    /**
     * @param {object} userData
     * @param {{ persistent?: boolean, token?: string }} [opts] persistent=true 写入 localStorage（关闭浏览器仍保持）；false 仅 sessionStorage（关标签即失效）
     */
    setUser(userData, opts = {}) {
      const persistent = typeof opts.persistent === 'boolean' ? opts.persistent : true
      this.user = userData
      try {
        if (persistent) {
          localStorage.setItem(LS_USER, JSON.stringify(userData))
          sessionStorage.removeItem(SS_USER)
        } else {
          sessionStorage.setItem(SS_USER, JSON.stringify(userData))
          localStorage.removeItem(LS_USER)
        }
      } catch {
        /* ignore quota / private mode */
      }
      if (typeof opts.token === 'string' && opts.token) {
        this.setToken(opts.token, persistent)
      }
    },
    setToken(token, persistent = true) {
      this.token = token || ''
      try {
        if (this.token) {
          if (persistent) {
            localStorage.setItem(LS_TOKEN, this.token)
            sessionStorage.removeItem(SS_TOKEN)
          } else {
            sessionStorage.setItem(SS_TOKEN, this.token)
            localStorage.removeItem(LS_TOKEN)
          }
        } else {
          localStorage.removeItem(LS_TOKEN)
          sessionStorage.removeItem(SS_TOKEN)
        }
      } catch {
        /* ignore */
      }
    },
    clearUser() {
      this.user = null
      this.token = ''
      try {
        localStorage.removeItem(LS_USER)
        sessionStorage.removeItem(SS_USER)
        localStorage.removeItem(LS_TOKEN)
        sessionStorage.removeItem(SS_TOKEN)
      } catch {
        /* ignore */
      }
    },
  },
})
