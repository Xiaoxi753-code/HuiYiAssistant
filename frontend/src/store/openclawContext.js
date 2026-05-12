import { defineStore } from 'pinia'

/**
 * 会意助手工作台全局上下文：与「智能长文速读」自适应卡片等模块共享，供对话助手读取。
 */
export const useOpenclawContextStore = defineStore('openclawContext', {
  state: () => ({
    /** @type {null | { content_type: string, dimensions: Record<string, string>, documentTitle: string, syncedAt: number }} */
    literatureCard: null,
  }),
  actions: {
    /**
     * @param {{ content_type: string, dimensions: Record<string, string>, documentTitle?: string }} payload
     */
    setLiteratureCard(payload) {
      if (!payload || typeof payload !== 'object') {
        this.literatureCard = null
        return
      }
      const documentTitle = String(payload.documentTitle || '').trim()
      this.literatureCard = {
        content_type: String(payload.content_type || '').trim() || '未分类',
        dimensions:
          payload.dimensions && typeof payload.dimensions === 'object' && !Array.isArray(payload.dimensions)
            ? { ...payload.dimensions }
            : {},
        documentTitle,
        syncedAt: Date.now(),
      }
    },
    clearLiteratureCard() {
      this.literatureCard = null
    },
  },
})
