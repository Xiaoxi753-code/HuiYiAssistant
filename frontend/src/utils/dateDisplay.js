/** 浏览器本地时区下的可读时间展示 */

function pad2(n) {
  return String(n).padStart(2, '0')
}

/**
 * @param {unknown} value ISO 或其它 Date.parse 可识别的字符串
 * @returns {number}
 */
export function parseDateMs(value) {
  if (value == null || value === '') {
    return NaN
  }
  const ms = Date.parse(String(value).trim())
  return Number.isFinite(ms) ? ms : NaN
}

/**
 * @param {unknown} value
 * @returns {string} YYYY-MM-DD HH:mm（本地）
 */
export function formatLocalDateTime(value) {
  const ms = parseDateMs(value)
  if (!Number.isFinite(ms)) {
    return '—'
  }
  const d = new Date(ms)
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`
}

function startOfLocalDay(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime()
}

/**
 * @param {unknown} value
 * @returns {string} 相对日历提示，如「今天 11:50」；无法计算时为空串
 */
export function formatRelativeCalendarHint(value) {
  const ms = parseDateMs(value)
  if (!Number.isFinite(ms)) {
    return ''
  }
  const target = new Date(ms)
  const now = new Date()
  const diffDays = Math.round((startOfLocalDay(target) - startOfLocalDay(now)) / 86400000)
  const hm = `${pad2(target.getHours())}:${pad2(target.getMinutes())}`
  if (diffDays === 0) {
    return `今天 ${hm}`
  }
  if (diffDays === 1) {
    return `明天 ${hm}`
  }
  if (diffDays === -1) {
    return `昨天 ${hm}`
  }
  if (diffDays === 2) {
    return `后天 ${hm}`
  }
  if (diffDays === -2) {
    return `前天 ${hm}`
  }
  return ''
}

/**
 * @param {unknown} value
 * @returns {string} 本地时间 + 可选相对提示
 */
export function formatLocalDateTimeWithRelative(value) {
  const base = formatLocalDateTime(value)
  if (base === '—') {
    return '—'
  }
  const rel = formatRelativeCalendarHint(value)
  return rel ? `${base}（${rel}）` : base
}
