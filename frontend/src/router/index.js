import { createRouter, createWebHistory } from 'vue-router'
import { ElMessage } from 'element-plus'
import { useUserStore } from '../store/user.js'
import LoginView from '../views/LoginView.vue'
import WorkspaceLayout from '../layouts/WorkspaceLayout.vue'
import DashboardView from '../views/DashboardView.vue'
import SkillLibraryView from '../views/SkillLibraryView.vue'
import GrowthDashboardView from '../views/GrowthDashboardView.vue'
import AdminView from '../views/AdminView.vue'

const router = createRouter({
  history: createWebHistory(import.meta.env.BASE_URL),
  routes: [
    {
      path: '/login',
      name: 'login',
      component: LoginView,
    },
    {
      path: '/',
      component: WorkspaceLayout,
      meta: { requiresAuth: true },
      children: [
        { path: '', redirect: { name: 'dashboard' } },
        {
          path: 'dashboard',
          name: 'dashboard',
          component: DashboardView,
        },
        {
          path: 'dashboard/chat',
          name: 'dashboard-chat',
          component: DashboardView,
          meta: { hideHeader: true },
        },
        {
          path: 'dashboard/companion',
          name: 'dashboard-companion',
          component: DashboardView,
          meta: { hideHeader: true },
        },
        {
          path: 'skills',
          name: 'skills',
          component: SkillLibraryView,
        },
        {
          path: 'growth',
          name: 'growth',
          component: GrowthDashboardView,
        },
      ],
    },
    {
      path: '/admin',
      name: 'admin',
      component: AdminView,
      meta: { requiresAuth: true, requiresAdmin: true },
    },
    {
      path: '/:pathMatch(.*)*',
      name: 'notFound',
      redirect: { name: 'login' },
    },
  ],
})

router.beforeEach((to, _from, next) => {
  const store = useUserStore()
  if (to.meta.requiresAuth && !store.user) {
    ElMessage.warning('请先登录或注册，以使用完整功能')
    next({ name: 'login' })
    return
  }
  if (to.meta.requiresAdmin && String(store.user?.role || '') !== 'admin') {
    ElMessage.warning('该页面仅管理员可访问')
    next({ name: 'dashboard' })
    return
  }
  if (to.name === 'login' && store.user) {
    next({ name: 'dashboard' })
    return
  }
  if (to.name === 'dashboard' && to.query.tab === 'chat') {
    const { tab: _t, ...rest } = to.query
    next({ name: 'dashboard-chat', query: rest })
    return
  }
  if (to.name === 'dashboard' && to.query.tab === 'companion') {
    const { tab: _t2, ...rest2 } = to.query
    next({ name: 'dashboard-companion', query: rest2 })
    return
  }
  next()
})

export default router
