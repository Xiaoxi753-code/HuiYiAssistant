import { defineStore } from 'pinia'

const LS_USER = 'openclaw_user'
const SS_USER = 'openclaw_user_session'

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

export const useUserStore = defineStore('user', {
  state: () => ({
    user: readStoredUser(),
  }),
  actions: {
    /**
     * @param {object} userData
     * @param {{ persistent?: boolean }} [opts] persistent=true 写入 localStorage（关闭浏览器仍保持）；false 仅 sessionStorage（关标签即失效）
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
    },
    clearUser() {
      this.user = null
      try {
        localStorage.removeItem(LS_USER)
        sessionStorage.removeItem(SS_USER)
      } catch {
        /* ignore */
      }
    },
  },
})
