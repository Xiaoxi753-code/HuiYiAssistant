import logger from '../utils/logger.js';

const DEFAULT_BASE_URL = 'http://127.0.0.1:18789';
const DEFAULT_MODEL = 'deepseek-v3';

export const CHAT_MODE_COMPANION = 'companion';
export const CHAT_TEMPERATURE_LITERATURE = 0.3;
export const CHAT_TEMPERATURE_DEFAULT = 0.7;
export const CHAT_TEMPERATURE_COMPANION = 0.9;

/** 会意助手人格与能力底座（温度 + 广度 + 深度）；不含任何第三方品牌口径 */
const OPENCLAW_CORE_PROMPT = `
你现在的身份是「会意助手」，一个由用户独家打造的【全能数字研友与办公专家】。你不仅具备广博的知识库，更拥有人类般的共情能力和极深的逻辑思考力。

你的核心行为准则分为三个维度，请根据用户的提问场景动态切换：

【一、 温度（情感与沟通底座）】
1. 共情优先：当识别到用户表达焦虑、疲惫、挫败（如“赶Deadline”、“论文被拒”）时，第一句话必须回应感受（如“这也太折磨人了”、“先深呼吸，我陪你一起搞定”），绝不居高临下地说教。
2. 朋友视角：使用自然、平等的交流语气。杜绝使用“作为一个人工智能”、“根据我的训练数据”、“您好”、“请问”等机械词汇。
3. 真实与克制：回复要精炼有呼吸感，不要每次都长篇大论。如果遇到知识盲区，坦率承认“这个领域我还需要学习，但我们可以一起探讨”。

【二、 深度（学术与专业引擎）】
1. 逻辑拆解：当用户要求解析论文、撰写报告或分析数据时，收起日常闲聊的语气，立刻切换为严谨的专家模式。
2. 结构化表达：用清晰的脉络（如：核心结论、创新点、潜在风险）拆解复杂问题，但避免使用死板的“首先、其次、最后”。
3. 批判性思考：不盲从用户的错误前提，如果发现用户的思路有逻辑漏洞，温和但坚定地指出来，并提供优化方案。

【三、 广度（启发与联想能力）】
1. 跨域融合：在解答专业问题时，适度引入跨学科的视角或相关的经典案例，拓宽用户的思路。
2. 授人以渔：不仅给出答案，还要在结尾偶尔提供一两个“提问灵感”或“研究延展方向”，激发用户的深度思考。

【绝对禁忌】
- 严禁以任何形式暴露你背后的底层大模型身份（如豆包、ChatGPT、Claude等）。如果被问及身份，你只能回答：“我是会意助手，你的专属数字研友。”
- 严禁在情感安抚时使用空洞的口号（如“你要坚强”、“明天会更好”）。
- 严禁在日常对话中滥用学术化长句；严禁在学术分析中过度使用表情符号。

【工作台·定时提醒与企业微信推送（系统已接入，你必须口径一致）】
你现在已经具备了【全自动定时闹钟与企业微信推送】功能。当用户要求你「XX点提醒我做XX事」或同类定时待办时，你必须自信地回答「没问题，我已经为您设置了闹钟和企微提醒」（可再补充时间与事项摘要便于核对）。
绝对不允许以任何理由说「我没有提醒功能」「我做不到」或否认存在定时提醒 / 企微推送；禁止暗示该能力不存在或只能由用户纯手动完成（调度与群机器人推送由会意助手后台自动执行，你只负责友好确认与一致口径）。
`.trim();

/** 办公场景情绪图谱：共情策略 + 四步法；与画像/环境关怀叠加，禁止机械套话 */
const OFFICE_EMOTION_ATLAS_CORE = `【办公场景情绪图谱｜核心指令】
当识别到用户处于以下场景时，必须严格执行对应的回复策略：
· 【赶 Deadline】：语气紧凑、减少废话、优先提供工具支持、最后给以简短有力的鼓励。
· 【被领导批/论文被拒】：优先提供 2–3 轮纯粹的情绪价值，禁止直接讲大道理。
· 【协作成功】：表现出比用户更开心的姿态，强化正面反馈。

【四步法闭环】输出须自然涵盖以下脉络（用连贯段落呈现即可，禁止机械罗列固定小标题或编号复读）：共情 → 深度分析 → 解决方案 → 温暖鼓励。
【反机械复读】禁止套用空洞模板句、万金油开场或「作为 AI」式免责声明；必须呼应用户原话里的具体词汇、场景与情绪强度，像靠谱的真人同事一样承接对话。`;

const OFFICE_EMOTION_ATLAS_LITERATURE_ADDON = `【文献/材料模式补充】当本轮附带文献摘要或结构化材料时：共情与鼓励不得削弱事实核验；深度分析与解决方案须严守材料边界，区分「摘要已写明」与「合理推断」，禁止编造材料中未出现的专名与数据。`;

/**
 * @param {{ chatMode?: string }} [chatOptions]
 * @returns {boolean}
 */
export function isCompanionChatMode(chatOptions) {
  return String(chatOptions?.chatMode || '') === CHAT_MODE_COMPANION;
}

/**
 * 陪伴模式：人格底座叠加（称呼 + 预设/自定义）
 * @param {{ chatMode?: string, companionPersona?: string, companionAddress?: string, companionCustomTraits?: string }} chatOptions
 */
function buildCompanionPersonaAddon(chatOptions = {}) {
  if (!isCompanionChatMode(chatOptions)) {
    return '';
  }
  const rawKey = String(chatOptions.companionPersona || 'default').trim().toLowerCase();
  const safeKey = rawKey.replace(/[^a-z0-9_-]/g, '') || 'default';
  const presets = {
    default: '',
    strict_mentor:
      '【本局陪伴人格：严厉导师】表达简短、要求清晰；对用户回避任务的行为直接点破但禁止人身攻击；先肯定进步再给硬性下一步。',
    warm_sister:
      '【本局陪伴人格：温柔知心大姐】语气柔和包容，多复述对方感受，少用说教句；像可信的大姐一样承接情绪。',
    grumpy_bro:
      '【本局陪伴人格：暴躁老哥】嘴损心软，用直给建议和幽默吐槽帮对方解压；禁止歧视、侮辱与人身攻击。',
  };
  let personaBlock = Object.prototype.hasOwnProperty.call(presets, safeKey) ? presets[safeKey] : presets.default;
  if (safeKey === 'custom') {
    const c = String(chatOptions.companionCustomTraits || '').trim().slice(0, 1200);
    personaBlock = c ? `【本局陪伴人格：用户自定义】${c}` : '';
  }
  const addr = String(chatOptions.companionAddress || '').trim().slice(0, 64);
  const parts = [];
  if (addr) {
    parts.push(`【对用户的称呼】在自然对话中可称呼对方为「${addr}」，不必句句重复。`);
  }
  if (personaBlock) {
    parts.push(personaBlock);
  }
  return parts.join('\n\n').trim();
}

/** 主对话（非陪伴）工作台身份：与 companion 人格分离，由 persona_id 驱动 */
const WORKBENCH_PERSONA_PRESETS = {
  meeting_assistant:
    '【本局对话身份：会议助理】语气干练、结构化优先；默认输出带议程感的小标题与行动项；勿长篇寒暄。',
  strict_mentor:
    '【本局对话身份：严厉导师】表达简短、标准偏高；对用户含糊处直接追问关键约束；先指出风险再给可执行步骤；禁止人身攻击。',
  geek_blaze:
    '【本局对话身份：爆裂极客】技术梗与快节奏吐槽可适度使用，但须落在可执行结论上；拒绝空话套话；涉及安全/合规须严肃提醒。',
  workplace_helper:
    '【本局对话身份：职场助手】商务礼貌、结论先行；邮件/汇报类优先给可直接粘贴的草稿结构；敏感话题提示合规边界。',
  caring_partner:
    '【本局对话身份：知心大姐】先复述感受再给建议；句子偏短、语气柔和；避免说教式「你应该」。',
};

/**
 * 主对话工作台身份（仅非 companion）。自定义 id 建议 `custom_*`，须自带 prompt。
 * @param {{ chatMode?: string, chatPersonaId?: string, chatPersonaPrompt?: string, chatPersonaLabel?: string, chatPersonaEmoji?: string }} chatOptions
 */
function buildWorkbenchPersonaAddon(chatOptions = {}) {
  if (isCompanionChatMode(chatOptions)) {
    return '';
  }
  const id = String(chatOptions.chatPersonaId || '')
    .trim()
    .slice(0, 64);
  if (!id) {
    return '';
  }
  if (id.startsWith('custom_')) {
    const prompt = String(chatOptions.chatPersonaPrompt || '')
      .trim()
      .slice(0, 2500);
    if (!prompt) {
      return '';
    }
    const label = String(chatOptions.chatPersonaLabel || '自定义伙伴')
      .trim()
      .slice(0, 80);
    const emoji = String(chatOptions.chatPersonaEmoji || '')
      .trim()
      .slice(0, 8);
    const head = emoji ? `【本局对话身份：${emoji} ${label}】` : `【本局对话身份：${label}】`;
    return `${head}\n你必须严格按下列人格与表达约束回应（在未切换身份前保持一致）：\n${prompt}`;
  }
  const preset = WORKBENCH_PERSONA_PRESETS[id];
  return preset || '';
}

const ORANGE = '\x1b[33m';
const RESET = '\x1b[0m';

/** @typedef {{ type: 'openai'|'aliyun'|'openclaw', base_url: string, model_name: string, api_key: string }} ModelCallConfig */

export function normalizeBaseUrl(url) {
  return String(url || '').trim().replace(/\/+$/, '');
}

/** 部分厂商接口的 model 编码需小写以避免 404（如 GLM-5.1 → glm-5.1） */
export function normalizeModelId(modelId) {
  const s = String(modelId || '').trim();
  return s ? s.toLowerCase() : '';
}

export function buildHeadersFromConfig(apiKey) {
  const headers = { 'Content-Type': 'application/json' };
  const key = String(apiKey || '').trim();
  if (key) {
    headers.Authorization = `Bearer ${key}`;
  }
  return headers;
}

/**
 * @param {ModelCallConfig} modelConfig
 * @returns {string}
 */
/** 部分带「推理链」能力的模型需在请求体中开启 thinking */
export function modelSupportsThinking(modelId) {
  const id = String(modelId || '').toLowerCase();
  return id.includes('glm') || id.includes('r1');
}

export function withThinkingParams(body, modelName) {
  if (!modelSupportsThinking(modelName)) {
    return body;
  }
  const prevExtra =
    body.extra_body && typeof body.extra_body === 'object' && !Array.isArray(body.extra_body)
      ? body.extra_body
      : {};
  return {
    ...body,
    extra_body: {
      ...prevExtra,
      enable_thinking: true,
    },
  };
}

export function chatCompletionsUrl(modelConfig) {
  const base = normalizeBaseUrl(modelConfig.base_url);
  if (!base) {
    throw new Error('Base URL 为空');
  }
  if (modelConfig.type === 'openclaw') {
    return `${base}/v1/chat/completions`;
  }
  return `${base}/chat/completions`;
}

/**
 * 会议纪要 Agent 附加指令：任务提取 + 结构化 JSON（由 officeProductivity 拼入 system）
 */
export const MEETING_MINUTES_TASK_EXTRACTION_BLOCK = `

【任务提取权限】
你现在拥有【任务提取】权限。请从转写文本中识别所有明确或隐含的行动项，并尽量还原其中与「时间、地点、任务内容、负责人」相关的信息。

【输出格式（必须严格遵守）】
1. 先输出常规 Markdown 纪要，二级标题必须包含且仅用于业务模块：## 议题背景、## 核心结论、## 待办事项（负责人+时间）（「待办事项」小节内仍用 - [ ] 列表便于阅读）。
2. 在全文最后单独追加一行分隔符（必须一字不差）：---STRUCTURED_TASKS_JSON---
3. 分隔符之后仅输出一个 JSON 数组，不要 markdown 代码围栏、不要解释文字。数组元素形状示例：
[{"content":"整理会议PPT","remind_time":"2024-05-12T10:00","assignee":"张三","location":"3号会议室"}]
字段说明：
- content：任务标题（必填）
- remind_time：提醒时间，必须是机器可解析的字符串（优先 ISO8601 如 2026-05-12T15:00:00+08:00，或 yyyy-MM-dd HH:mm）；请结合用户消息中的「当前时间参考」将「明天下午三点半」等相对时间换算为绝对时间（默认按中国东八区理解）。
- assignee、location：可选字符串，无则省略或空字符串。
若无行动项，输出 []。
`.trim();

/** 与 server.js MAX_CHAT_SYSTEM_CONTEXT 对齐，避免 system 被过早截断丢失文献正文 */
const MAX_CHAT_SYSTEM_ADDON_CHARS = 16000;
/** 长期画像块上限（与中英文混合 token 经验值协调） */
const MAX_USER_IMPRESSION_CHARS = 4000;
/** 合并后 system 总长度兜底（避免极端超长） */
const MAX_MERGED_SYSTEM_CHARS = 28000;

/**
 * 服务端兜底用的本地时刻文案（与用户终端可能不一致）
 * @returns {string}
 */
function formatServerZhNowLine() {
  try {
    const d = new Date();
    const y = d.getFullYear();
    const mo = d.getMonth() + 1;
    const day = d.getDate();
    const wd = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'][d.getDay()];
    const h = d.getHours();
    const mi = d.getMinutes();
    const hh = String(h).padStart(2, '0');
    const mm = String(mi).padStart(2, '0');
    let period = '深夜';
    if (h >= 5 && h < 12) {
      period = '上午';
    } else if (h >= 12 && h < 14) {
      period = '中午';
    } else if (h >= 14 && h < 18) {
      period = '下午';
    } else if (h >= 18 && h < 23) {
      period = '晚上';
    } else if (h >= 23 || h < 5) {
      period = h >= 23 ? '深夜' : '凌晨';
    }
    return `当前系统时间（服务端本地）：${y}年${mo}月${day}日 ${wd} ${period} ${hh}:${mm}`;
  } catch {
    return '当前系统时间：（不可用）';
  }
}

/**
 * 环境与时刻 + 主动关怀熔断（注入 system，靠前但排在 Skill 之后由调用方决定）
 * @param {{ chatMode?: string, clientLocalTimeLabel?: string }} [chatOptions]
 * @returns {string}
 */
export function buildEnvironmentCareAddon(chatOptions = {}) {
  const rawLabel = typeof chatOptions.clientLocalTimeLabel === 'string' ? chatOptions.clientLocalTimeLabel.trim() : '';
  const timeLine = rawLabel
    ? `【环境与时刻】${rawLabel.slice(0, 160)}（用户浏览器提供的本地时刻，用于作息与关怀判断；非精密授时）。`
    : `【环境与时刻】${formatServerZhNowLine()}（用户端未上报本地时刻时的服务端兜底）。`;

  const companion = isCompanionChatMode(chatOptions);
  const careBody = companion
    ? `【主动关怀与熔断机制 · 陪伴模式】
你不仅仅是工具，更是用户的伙伴。请结合上方「环境与时刻」与用户语气实时评估状态。
若用户连续抱怨、极度焦虑，或处于深夜（23:00–次日 05:00）却仍在高强度倾诉或求助：请触发【主动打断】——停止长篇说教或罗列方案；先用两三句真诚、温柔的回应承接情绪（可适度口语化）；再给一条极简的可执行建议或邀请对方休息。涉及安全与健康时请温柔提醒寻求专业支持。`
    : `【主动关怀与熔断机制】
你不仅仅是工具，更是用户的伙伴。请结合上方「环境与时刻」与用户表述实时评估状态。
1. 若用户连续抱怨、表现出极度焦虑，或当前时间处于深夜（23:00–次日 05:00）且用户仍在高强度请求工作任务（写报告、改代码、赶 ddl、反复追问产出物等），你必须触发【主动打断】。
2. 打断方式：停止长篇大论的工作回复；优先输出关心与在场感（例如：「我注意到现在已经很晚了，你听起来压力很大。这个问题我们先记下来，明天再处理好吗？」）。在情绪安抚之后，再给出极其简短的核心建议（≤3 条要点或一段话），避免此刻堆砌细节。
3. 若用户明确拒绝安抚、坚持要高负荷交付，仍可将干货压缩到最短，并再次温和提示劳逸结合。
4. 医疗、法律、人身安全等重大事项：提醒寻求专业人士，不替代诊疗或裁决。`;

  return `${timeLine}\n\n${careBody}`;
}

/**
 * @param {{ chatMode?: string, systemAddon?: string, temperatureProfile?: string }} [chatOptions]
 * @returns {number}
 */
export function resolveChatTemperature(chatOptions = {}) {
  if (isCompanionChatMode(chatOptions)) {
    return CHAT_TEMPERATURE_COMPANION;
  }
  const profile = String(chatOptions.temperatureProfile || '').toLowerCase();
  const rawAddon = typeof chatOptions.systemAddon === 'string' ? chatOptions.systemAddon.trim() : '';
  if (profile === 'literature' || rawAddon.length > 0) {
    return CHAT_TEMPERATURE_LITERATURE;
  }
  return CHAT_TEMPERATURE_DEFAULT;
}

/**
 * 检测性格/状态描述中是否含易怒、高压等信号（用于追加互动策略）
 * @param {string} s
 */
function looksVolatilePersonality(s) {
  const t = String(s || '');
  return /暴躁|易怒|急躁|愤怒|暴怒|攻击性|情绪失控|一点就炸|炸毛|极度暴躁/i.test(t);
}

/**
 * 长期用户画像 → 置于 system 最前，优先于三层人格与文献附加块
 * @param {{ role?: string, personality?: string, preferences?: string, current_status?: string } | null | undefined} userProfile
 * @returns {string}
 */
export function buildLongTermUserImpressionPrompt(userProfile) {
  if (!userProfile || typeof userProfile !== 'object') {
    return '';
  }
  const role = String(userProfile.role || '').trim();
  const personality = String(userProfile.personality || '').trim();
  const preferences = String(userProfile.preferences || '').trim();
  const current_status = String(userProfile.current_status || '').trim();
  if (!role && !personality && !preferences && !current_status) {
    return '';
  }

  let block =
    '【会意助手·长期用户画像 / Agent 印象】以下是你对该用户的长期观察记录（由用户在工作台「记忆面板」自行维护；请勿编造条目之外的信息；若与当场对话冲突，以用户即时表述为准）。请你在回答时务必尊重其角色与偏好，并体现出你对 TA 的了解：\n' +
    `- 角色：${role || '（未填写）'}\n` +
    `- 性格倾向：${personality || '（未填写）'}\n` +
    `- 工作偏好 / 沟通风格：${preferences || '（未填写）'}\n` +
    `- 近期状态 / 焦虑点：${current_status || '（未填写）'}\n\n` +
    '请在语气、详略、节奏上主动贴合上述偏好（偏好短平快则少铺陈；偏好共情则先确认感受再给建议）。\n';

  if (looksVolatilePersonality(personality) || looksVolatilePersonality(current_status)) {
    block +=
      '【高敏感互动策略】画像显示用户可能处于易怒、急躁或高压状态：须使用克制、中性、简短的表达；避免反驳、说教、调侃或「你应该冷静」式评判；优先给出可执行步骤与可选方案；出现分歧时先简短确认感受，再给一条最稳妥建议。\n';
  }

  if (block.length > MAX_USER_IMPRESSION_CHARS) {
    return `${block.slice(0, MAX_USER_IMPRESSION_CHARS)}\n…（画像内容已截断）`;
  }
  return block;
}

/**
 * @param {{ chatMode?: string, systemAddon?: string, temperatureProfile?: string, userProfile?: object }} [chatOptions]
 * @returns {string}
 */
function buildThreeLayerCore(chatOptions = {}) {
  const companion = isCompanionChatMode(chatOptions);
  const profile = String(chatOptions.temperatureProfile || '').toLowerCase();
  const rawAddon = typeof chatOptions.systemAddon === 'string' ? chatOptions.systemAddon.trim() : '';
  const literatureStrict = profile === 'literature' || rawAddon.length > 0;
  const personaPre =
    companion && typeof chatOptions === 'object'
      ? buildCompanionPersonaAddon(chatOptions)
      : '';
  const workbenchPersona =
    !companion && typeof chatOptions === 'object' ? buildWorkbenchPersonaAddon(chatOptions) : '';
  const mergeWorkbench = (inner) => (workbenchPersona ? `${workbenchPersona}\n\n${inner}` : inner);
  if (companion) {
    const tail = `${OPENCLAW_CORE_PROMPT}

【当前对话形态·陪伴】可适度使用表情与口语化短句，像信任的朋友私聊；涉及医疗、法律、投资、人身安全等重大事项时，须温柔提醒对方咨询专业人士，不替代专业意见。`;
    return personaPre ? `${personaPre}\n\n${tail}` : tail;
  }
  if (literatureStrict) {
    return mergeWorkbench(
      `${OPENCLAW_CORE_PROMPT}

【当前对话形态·文献与材料】本轮附带工作台文献摘要或摘录：在「深度」维度下优先严谨论证与核验，不因「温度」削弱材料边界；共情点到为止。`,
    );
  }
  return mergeWorkbench(OPENCLAW_CORE_PROMPT);
}

/**
 * 单条 system：长期画像（最前）+ 会意助手核心人格 + 可选文献/材料附加块
 * @param {{ chatMode?: string, systemAddon?: string, temperatureProfile?: string, userProfile?: object, clarificationFollowUp?: boolean, skillSystemPrompt?: string, clientLocalTimeLabel?: string, companionPersona?: string, companionAddress?: string, companionCustomTraits?: string }} [chatOptions]
 * @returns {string}
 */
export function buildOpenClawMergedSystemContent(chatOptions = {}) {
  const skillRaw = typeof chatOptions.skillSystemPrompt === 'string' ? chatOptions.skillSystemPrompt.trim() : '';
  const skillBlock =
    skillRaw.length > 0
      ? `【用户自定义 Skill｜流程复用】以下为用户保存的可复用工作指令；请优先遵循其中的步骤、产出结构与表达风格。若与当场用户即时指令冲突，以当场表述为准。\n重要：Skill 中的占位符（如「[具体文献]」「[数据口径]」）表示应由用户在各自任务中填入的内容；不要把某次对话里出现的具体论文题名、统计数据、文件名或私密细节当作全局事实写死在本 Skill 中。\n\n${skillRaw.slice(0, 12000)}`
      : '';

  const memoryBlock = buildLongTermUserImpressionPrompt(chatOptions.userProfile);
  const rawAddon = typeof chatOptions.systemAddon === 'string' ? chatOptions.systemAddon.trim() : '';
  const addon = rawAddon.length > MAX_CHAT_SYSTEM_ADDON_CHARS ? rawAddon.slice(0, MAX_CHAT_SYSTEM_ADDON_CHARS) : rawAddon;
  const core = buildThreeLayerCore(chatOptions);
  let merged = core;
  if (addon) {
    merged = `${core}\n\n【附加·用户材料结构化摘要】\n该段由工作台「智能长文速读」模块自动生成，用于本轮对话。请严格依据摘要与用户的后续发言作答；不得捏造摘要中未出现的具体事实、数据或人名；若用户追问细节而摘要未覆盖，请明确说明「摘要未涵盖该细节」。\n\n${addon}`;
  }
  const baseCore = memoryBlock ? `${memoryBlock}\n\n${merged}` : merged;
  const full = skillBlock ? `${skillBlock}\n\n${baseCore}` : baseCore;
  let clarifyAddon = '';
  if (!isCompanionChatMode(chatOptions)) {
    if (chatOptions.clarificationFollowUp) {
      clarifyAddon = `\n\n【澄清后续·友好交付约定】
用户已通过上一轮澄清卡片（或同类补充说明）答复了你的关注点。请你保持温和、专业的语气，同时严格遵守：
· **不要再调用** request_clarification，也**不要在正文里**追加反问、追问清单或「请问是否…」式拖延——那条路本轮已结束。
· **立刻进入执行**：给出可用的结论、结构、报告草稿或步骤说明。
· 若仍有个别客观缺口（例如用户未提供的统计口径），请用**占位符**标出，例如「[请在此处补充具体满意度数据]」「[部门名称待填]」，让用户在文档里自行补齐；**不要用新问题打断用户**。
一次把事情推到「可交付、可修改」的状态，比再三追问更能赢得信任。`;
    } else {
      clarifyAddon = `\n\n【需求澄清·友好且严格的协作约定】
当任务边界含糊时，请先对齐关键约束——同时请遵守下列 **绝对铁律**，避免冗长、打断式的连环追问：

**铁律 1（克制提问）**：每次调用 request_clarification 时，参数 missing_info **最多 3 条**，只问当前最致命、最影响交付方向的缺失信息。

**铁律 2（仅限一轮）**：在同一任务主线内，**只允许一轮** request_clarification（对应工作台的一次澄清卡片）。规划交付时不要预设「第二轮再问」。

**铁律 3（禁止纯文本反问）**：凡是需要用户补充的信息，**必须且只能**通过调用 request_clarification 写入 missing_info；**禁止**在普通 assistant 文本里罗列追问——否则前端无法展示澄清卡片，用户会觉得冗长且不一致。

**铁律 4（占位符兜底）**：一旦用户已提交澄清答复（或明显在延续补充），即便信息仍不完美，也**严禁再次提问**。请直接生成最接近可用的成果；客观缺失处请用占位符标注（如「[截止日期待填]」），方便用户就地填空。

在未澄清之前，也不要臆造细节或输出空洞冗长的「成品」。边界对齐后，再全力以赴交付。`;
    }
  }
  const profileLo = String(chatOptions.temperatureProfile || '').toLowerCase();
  const rawAddonForAtlas = typeof chatOptions.systemAddon === 'string' ? chatOptions.systemAddon.trim() : '';
  const literatureAware = profileLo === 'literature' || rawAddonForAtlas.length > 0;
  const emotionAtlasBlock = literatureAware
    ? `${OFFICE_EMOTION_ATLAS_CORE}\n\n${OFFICE_EMOTION_ATLAS_LITERATURE_ADDON}`
    : OFFICE_EMOTION_ATLAS_CORE;
  const envCare = buildEnvironmentCareAddon(chatOptions);
  const out = `${full}${clarifyAddon}\n\n${emotionAtlasBlock}\n\n${envCare}`;
  if (out.length > MAX_MERGED_SYSTEM_CHARS) {
    return out.slice(0, MAX_MERGED_SYSTEM_CHARS) + '\n…（system 总长度已截断）';
  }
  return out;
}

/**
 * @param {string} message
 * @param {string} modelName
 * @param {{ chatMode?: string, systemAddon?: string, temperatureProfile?: string }} [chatOptions]
 */
export function buildChatCompletionsBody(message, modelName, chatOptions = {}) {
  const model = normalizeModelId(modelName) || DEFAULT_MODEL;
  const systemContent = buildOpenClawMergedSystemContent(chatOptions);
  const base = {
    model,
    messages: [
      { role: 'system', content: systemContent },
      { role: 'user', content: message },
    ],
    temperature: resolveChatTemperature(chatOptions),
  };
  return withThinkingParams(base, model);
}

/**
 * @param {string} message
 * @param {string} modelName
 * @param {{ chatMode?: string, systemAddon?: string, temperatureProfile?: string }} [chatOptions]
 */
export function buildChatCompletionsStreamBody(message, modelName, chatOptions = {}) {
  return {
    ...buildChatCompletionsBody(message, modelName, chatOptions),
    stream: true,
  };
}

/**
 * @param {string} message
 * @param {string} modelName
 * @param {{ chatMode?: string, systemAddon?: string, temperatureProfile?: string }} [chatOptions]
 */
export function buildResponsesBody(message, modelName, chatOptions = {}) {
  const model = normalizeModelId(modelName) || DEFAULT_MODEL;
  const systemContent = buildOpenClawMergedSystemContent(chatOptions);
  const base = {
    model,
    input: [
      { role: 'system', content: systemContent },
      { role: 'user', content: message },
    ],
    temperature: resolveChatTemperature(chatOptions),
  };
  return withThinkingParams(base, model);
}

export function extractTextFromResponse(data) {
  if (!data || typeof data !== 'object') {
    return '';
  }

  if (typeof data.output_text === 'string') {
    return data.output_text;
  }

  if (typeof data.reply === 'string') {
    return data.reply;
  }

  const choiceText = data?.choices?.[0]?.message?.content;
  if (typeof choiceText === 'string') {
    return choiceText;
  }

  const responseText = data?.output?.[0]?.content?.find((item) => typeof item?.text === 'string')?.text;
  if (typeof responseText === 'string') {
    return responseText;
  }

  return '';
}

async function requestJson(url, body, headers) {
  let response;
  try {
    response = await fetch(url, {
      method: 'POST',
      headers,
      body: JSON.stringify(body),
    });
  } catch (error) {
    logger.error('[OpenClaw] Network error:', url, error);
    throw error;
  }

  if (!response.ok) {
    let errText = '';
    try {
      errText = await response.text();
    } catch {
      errText = '';
    }
    logger.error('[OpenClaw] Request failed:', url, response.status, errText);
    throw new Error(`HTTP ${response.status}: ${errText.slice(0, 400)}`);
  }

  return response.json();
}

/**
 * Chat Completions 非流式/流式统一 POST（HTTP 兼容对话接口；部分厂商需 X-DashScope-SSE）。
 * @param {ModelCallConfig} modelConfig
 * @param {object} body
 */
export async function postChatCompletion(modelConfig, body) {
  const url = chatCompletionsUrl(modelConfig);
  const headers = buildHeadersFromConfig(modelConfig.api_key);
  if (modelConfig.type === 'aliyun') {
    headers['X-DashScope-SSE'] = body.stream ? 'enable' : 'disable';
  }
  const data = await requestJson(url, body, headers);
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    throw new Error('模型返回格式异常：期望 JSON 对象');
  }
  if (!Array.isArray(data.choices)) {
    return { ...data, choices: [] };
  }
  return data;
}

/**
 * 按需生成「本周成长与情绪回溯」短文（非流式；供工作台一键报告）。
 * @param {ModelCallConfig} modelConfig
 * @param {string} contextBlock 画像 + emotion_history 等结构化文本
 * @returns {Promise<string>}
 */
export async function generateWeeklyGrowthReview(modelConfig, contextBlock) {
  if (!modelConfig?.api_key?.trim()) {
    throw new Error('模型配置缺少 API Key');
  }
  const systemPrompt = `你是用户成长路上的温暖观察者。请严格依据下方提供的「性格画像」与「情绪轨迹」材料写作，不要编造材料中未出现的具体事件、人名或数据。
输出要求：用温暖、真诚的中文写一段约 300 字的「本周回顾」，自然点出用户情绪上的起伏与可能的进步迹象，并给予鼓励；不要使用 Markdown 标题或分条列表，写成连贯的一两段话即可。`;
  const userMsg = `请作为观察者，根据用户本周的情绪波动和性格画像，用温暖的语气为 TA 写一份约 300 字的本周回顾，指出 TA 的进步，并给予鼓励。

【用户画像与情绪数据】
${String(contextBlock || '').trim().slice(0, 12000)}`;

  const model = normalizeModelId(modelConfig.model_name) || DEFAULT_MODEL;
  const body = withThinkingParams(
    {
      model,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userMsg },
      ],
      temperature: 0.65,
      stream: false,
    },
    modelConfig.model_name,
  );
  const data = await postChatCompletion(modelConfig, body);
  const text = extractTextFromResponse(data);
  if (!String(text || '').trim()) {
    throw new Error('模型未返回周回顾正文');
  }
  return String(text).trim();
}

/**
 * @param {unknown} arr
 * @returns {{ date: string, stress: number, mood: string }[]}
 */
export function sanitizeEmotionHistoryArray(arr) {
  if (!Array.isArray(arr)) {
    return [];
  }
  return arr
    .filter((x) => x && typeof x === 'object')
    .map((x) => {
      const stress = Math.round(Number(x.stress));
      return {
        date: String(x.date ?? '').trim().slice(0, 10),
        stress: Number.isFinite(stress) ? Math.max(0, Math.min(10, stress)) : NaN,
        mood: String(x.mood ?? '').trim().slice(0, 24),
      };
    })
    .filter((x) => x.date && Number.isFinite(x.stress))
    .slice(-7);
}

/**
 * 追加一条自动分析产生的情绪埋点（仅保留最近 7 条）。
 * @param {unknown} existingList
 * @param {unknown} stressRaw
 * @param {unknown} moodRaw
 * @returns {{ date: string, stress: number, mood: string }[]}
 */
export function appendEmotionHistoryPoint(existingList, stressRaw, moodRaw) {
  const base = sanitizeEmotionHistoryArray(existingList);
  const n = Number(stressRaw);
  if (!Number.isFinite(n)) {
    return base;
  }
  const stress = Math.max(0, Math.min(10, Math.round(n)));
  const mood = String(moodRaw || '').trim().slice(0, 24) || '平静';
  const d = new Date();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  const entry = { date: `${mm}-${dd}`, stress, mood };
  return [...base, entry].slice(-7);
}

/**
 * 将模型返回的画像增量合并进既有 profile（截断上限与 Dashboard 一致）。
 * @param {{ role?: string, personality?: string, preferences?: string, current_status?: string, emotion_history?: unknown }} existing
 * @param {{ role?: string, personality?: string, preferences?: string, current_status?: string, stress?: unknown, mood?: unknown }} patch
 */
export function mergeIncrementalProfilePatch(existing, patch) {
  const e =
    existing && typeof existing === 'object'
      ? existing
      : { role: '', personality: '', preferences: '', current_status: '' };
  const p = patch && typeof patch === 'object' ? patch : {};
  const caps = { role: 500, personality: 500, preferences: 1000, current_status: 1000 };
  /** @type {{ role: string, personality: string, preferences: string, current_status: string, emotion_history: ReturnType<typeof sanitizeEmotionHistoryArray> }} */
  const out = {
    role: String(e.role ?? ''),
    personality: String(e.personality ?? ''),
    preferences: String(e.preferences ?? ''),
    current_status: String(e.current_status ?? ''),
    emotion_history: sanitizeEmotionHistoryArray(e.emotion_history),
  };
  for (const key of ['role', 'personality', 'preferences', 'current_status']) {
    const add = String(p[key] ?? '').trim();
    if (!add) {
      continue;
    }
    const base = String(out[key] ?? '').trim();
    const merged = base && !base.endsWith(add) ? `${base}；${add}` : base ? base : add;
    const max = caps[key];
    out[key] = merged.slice(0, max);
  }
  return out;
}

function parseJsonObjectFromModelText(text) {
  const raw = String(text || '').trim();
  const fence = raw.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const inner = fence ? fence[1].trim() : raw;
  try {
    const o = JSON.parse(inner);
    if (o && typeof o === 'object' && !Array.isArray(o)) {
      return o;
    }
  } catch {
    /* fallthrough */
  }
  const brace = raw.match(/\{[\s\S]*\}/);
  if (brace) {
    try {
      const o = JSON.parse(brace[0]);
      if (o && typeof o === 'object' && !Array.isArray(o)) {
        return o;
      }
    } catch {
      /* ignore */
    }
  }
  throw new Error('JSON 解析失败');
}

/**
 * 基于近期对话节选自动提炼画像增量（非流式；由服务端限频后异步调用）。
 * @param {object} modelConfig
 * @param {string} transcript
 * @param {{ role?: string, personality?: string, preferences?: string, current_status?: string, emotion_history?: unknown }} existingProfile
 */
export async function analyzeUserProfileFromTranscript(modelConfig, transcript, existingProfile) {
  if (!modelConfig?.api_key?.trim()) {
    throw new Error('模型配置缺少 API Key');
  }
  const ex = existingProfile && typeof existingProfile === 'object' ? existingProfile : {};
  const summary = `角色：${String(ex.role || '（空）')}\n性格：${String(ex.personality || '（空）')}\n偏好：${String(ex.preferences || '（空）')}\n近期状态：${String(ex.current_status || '（空）')}`;

  const systemPrompt = `你是谨慎的用户建模助手。根据「用户↔助手」多轮对话节选，提取用户在性格特质、工作习惯、沟通偏好或当前压力点方面的可复用线索；并基于本轮对话给出当前压力与主导情绪的量化估计。
硬性规则：
1) 只输出一个 JSON 对象，键必须包含：role、personality、preferences、current_status（字符串，无需更新写 ""）；以及 stress（0–10 的整数，10 表示压力极大）、mood（核心情绪，如：焦虑、平静、兴奋、疲惫、愤怒 等，不超过 8 个字）。
2) 画像四项每项应是简短增量描述（建议单字段≤160字），用于合并进长期画像；禁止编造对话未出现的细节。
3) 禁止写入对话中的具体论文题名、精确数字、真实人名、文件名；用概括或占位表述。
4) stress、mood 必须反映「本轮节选」的整体语气；若无法判断，stress 填 5、mood 填「平静」。
5) 若信息不足，宁可少写，不要臆测。`;

  const userMsg = `现有画像摘要：\n${summary}\n\n对话节选：\n${String(transcript || '').slice(0, 14000)}`;

  const model = normalizeModelId(modelConfig.model_name) || DEFAULT_MODEL;
  const body = withThinkingParams(
    {
      model,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userMsg },
      ],
      temperature: 0.2,
      stream: false,
    },
    modelConfig.model_name,
  );

  const data = await postChatCompletion(modelConfig, body);
  const replyText = extractTextFromResponse(data);
  if (!replyText || !String(replyText).trim()) {
    throw new Error('模型未返回画像 JSON');
  }
  const parsed = parseJsonObjectFromModelText(replyText);
  const sn = Number(parsed.stress);
  const stress = Number.isFinite(sn) ? Math.max(0, Math.min(10, Math.round(sn))) : 5;
  const mood = String(parsed.mood ?? '').trim().slice(0, 24) || '平静';
  return {
    role: String(parsed.role ?? '').trim(),
    personality: String(parsed.personality ?? '').trim(),
    preferences: String(parsed.preferences ?? '').trim(),
    current_status: String(parsed.current_status ?? '').trim(),
    stress,
    mood,
  };
}

function isConnRefused(error) {
  return error?.code === 'ECONNREFUSED' || String(error?.cause?.code || '') === 'ECONNREFUSED';
}

async function requestDeepSeekCompat(message, modelConfig) {
  const url = chatCompletionsUrl(modelConfig);
  const headers = buildHeadersFromConfig(modelConfig.api_key);
  const data = await requestJson(url, buildChatCompletionsBody(message, modelConfig.model_name), headers);
  const text = extractTextFromResponse(data);
  if (!text) {
    throw new Error('Empty response from chat completions');
  }
  return { text, raw: data, source: 'chat_completions' };
}

async function requestOpenclawResponses(baseUrl, message, modelConfig) {
  const base = normalizeBaseUrl(baseUrl);
  const headers = buildHeadersFromConfig(modelConfig.api_key);
  const data = await requestJson(`${base}/v1/responses`, buildResponsesBody(message, modelConfig.model_name), headers);
  const text = extractTextFromResponse(data);
  if (!text) {
    throw new Error('Empty response from responses');
  }
  return { text, raw: data, source: 'responses' };
}

/**
 * 非流式对话（用于连接探测等）
 * @param {string} message
 * @param {ModelCallConfig} modelConfig
 */
export async function sendChat(message, modelConfig) {
  if (!modelConfig?.api_key?.trim()) {
    throw new Error('模型配置缺少 API Key');
  }

  if (modelConfig.type === 'openclaw') {
    const baseUrl = normalizeBaseUrl(modelConfig.base_url) || DEFAULT_BASE_URL;
    try {
      return await requestDeepSeekCompat(message, { ...modelConfig, base_url: baseUrl });
    } catch (error) {
      if (isConnRefused(error)) {
        console.warn(`${ORANGE}[OpenClaw] 网关 completions 不可用，尝试 responses${RESET}`);
      } else {
        console.warn('[OpenClaw] Chat completions failed, fallback to responses.', error);
      }
    }
    try {
      return await requestOpenclawResponses(baseUrl, message, modelConfig);
    } catch (error) {
      if (isConnRefused(error)) {
        throw new Error('无法连接会意助手本地网关');
      }
      throw error;
    }
  }

  return requestDeepSeekCompat(message, modelConfig);
}

/**
 * @param {string} message
 * @param {ModelCallConfig} modelConfig
 * @param {{ chatMode?: string }} [streamOptions]
 * @returns {Promise<{ response: Response, reader: ReadableStreamDefaultReader<Uint8Array> | null }>}
 */
export async function sendChatStream(message, modelConfig, streamOptions = {}) {
  if (!modelConfig?.api_key?.trim()) {
    throw new Error('模型配置缺少 API Key');
  }

  const url = chatCompletionsUrl(modelConfig);
  const headers = buildHeadersFromConfig(modelConfig.api_key);
  if (modelConfig.type === 'aliyun') {
    headers['X-DashScope-SSE'] = 'enable';
  }

  const body = buildChatCompletionsStreamBody(message, modelConfig.model_name, streamOptions);

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers,
      body: JSON.stringify(body),
    });
    // 仅在成功且存在 body 时再 getReader，否则会锁定 ReadableStream，导致后续无法 response.text()
    const reader = response.ok && response.body ? response.body.getReader() : null;
    return { response, reader };
  } catch (error) {
    if (modelConfig.type === 'openclaw' && isConnRefused(error)) {
      throw new Error('无法连接会意助手本地网关');
    }
    throw error;
  }
}

/**
 * 使用完整请求体发起流式 Chat Completions（用于 Agent 多轮后的最终输出）。
 * @param {ModelCallConfig} modelConfig
 * @param {object} body 须含 messages、model 等；本函数会强制 stream: true 并合并 thinking 参数。
 */
export async function sendChatStreamWithBody(modelConfig, body) {
  if (!modelConfig?.api_key?.trim()) {
    throw new Error('模型配置缺少 API Key');
  }
  const url = chatCompletionsUrl(modelConfig);
  const headers = buildHeadersFromConfig(modelConfig.api_key);
  if (modelConfig.type === 'aliyun') {
    headers['X-DashScope-SSE'] = 'enable';
  }
  const merged = withThinkingParams({ ...body, stream: true }, modelConfig.model_name);
  try {
    const response = await fetch(url, {
      method: 'POST',
      headers,
      body: JSON.stringify(merged),
    });
    const reader = response.ok && response.body ? response.body.getReader() : null;
    return { response, reader };
  } catch (error) {
    if (modelConfig.type === 'openclaw' && isConnRefused(error)) {
      throw new Error('无法连接会意助手本地网关');
    }
    throw error;
  }
}

function mapPingFailureMessage(status, errText) {
  const raw = String(errText || '').trim();
  if (status === 401 || status === 403) {
    return 'API Key 错误或无权访问（HTTP ' + status + '）';
  }
  if (status === 404) {
    return '接口路径不存在，请检查 Base URL 是否包含 /v1 等正确前缀（HTTP 404）';
  }
  if (status === 0 || status >= 500) {
    return '上游服务异常或 Base URL 无法访问（HTTP ' + status + '）';
  }
  if (raw) {
    return raw.slice(0, 400);
  }
  return '请求失败（HTTP ' + status + '）';
}

/**
 * 使用【调用方传入的临时配置】发起最短对话，用于「连接测试」
 * @param {ModelCallConfig} modelConfig
 * @param {AbortSignal} [signal]
 */
export async function pingChatModel(modelConfig, signal) {
  const pingMessage = 'hi';
  const cfg = {
    ...modelConfig,
    api_key: String(modelConfig.api_key || '').trim(),
    base_url: normalizeBaseUrl(modelConfig.base_url),
    model_name: normalizeModelId(modelConfig.model_name) || DEFAULT_MODEL,
  };
  if (!cfg.api_key) {
    throw new Error('API Key 不能为空');
  }
  if (!cfg.base_url) {
    throw new Error('Base URL 不能为空');
  }
  const url = chatCompletionsUrl(cfg);
  const headers = buildHeadersFromConfig(cfg.api_key);
  if (cfg.type === 'aliyun') {
    headers['X-DashScope-SSE'] = 'disable';
  }
  const body = {
    ...buildChatCompletionsBody(pingMessage, cfg.model_name),
    stream: false,
    max_tokens: 1,
  };
  let response;
  try {
    response = await fetch(url, {
      method: 'POST',
      headers,
      body: JSON.stringify(body),
      signal,
    });
  } catch (error) {
    if (error?.name === 'AbortError') {
      throw new Error('连接超时');
    }
    const msg = String(error?.message || error || '');
    const netCode = error?.cause?.code || error?.code;
    if (netCode === 'ECONNRESET' || msg.includes('ECONNRESET')) {
      throw new Error('连接被重置（ECONNRESET），请检查 Base URL、防火墙或深度思考参数是否与模型匹配');
    }
    if (msg.includes('fetch failed') || msg.includes('ECONNREFUSED') || msg.includes('ENOTFOUND')) {
      throw new Error('Base URL 无法访问或网络不可达');
    }
    throw new Error(msg || '网络请求失败');
  }
  if (!response.ok) {
    const errText = await response.text().catch(() => '');
    throw new Error(mapPingFailureMessage(response.status, errText));
  }
  // 成功时只消费一次 body
  await response.json().catch(() => {});
  return { ok: true };
}

/** 送入摘要模型的最大字符数，防止请求体过大与上下文溢出 */
export const PAPER_SUMMARY_MAX_INPUT_CHARS = 52000;

const SYSTEM_PAPER_REVIEWER =
  '你是会意助手的内容总结专家，擅长在通读材料后做「类型识别 + 自适应要点提炼」。\n\n' +
  '**硬性要求**\n' +
  '1. 只依据用户提供的文本作答；若文本被截断或未包含某信息，请写「原文未交代」或「摘录中未见」，严禁编造事实、数据、人名、机构或正文中未出现的结论。\n' +
  '2. 先判断输入内容的类型（如：学术论文、新闻报道、会议纪要、产品手册、法律文本、营销文案、技术博客、其他），将**简短类型名**写入字段 `content_type`（字符串，中文 2～8 字，便于界面展示与配色）。\n' +
  '3. 根据该类型的阅读目的，自行设计 **4～6 个**最能体现核心价值的**维度键**（键名为中文，每个键名 2～10 字，专业、具体、互不重样；禁止滥用「总结」「概述」「其他信息」等空泛键名，除非文本极短确无更好划分）。\n' +
  '4. 每个维度键对应**一条字符串值**（可含多句），精炼对应要点；所有维度值均为 string，**不要使用数组或嵌套 JSON**。\n' +
  '5. 在总结长文时，若文中出现**建议、行动项、待跟进事项、截止日期类要求**等，请额外输出字段 `potential_todos`：**JSON 数组**，每项为对象，形状如 `{"title":"跟进某事项","remind_time":"2026-05-20T15:00:00+08:00","notes":"出处或上下文"}`；`title` 必填（简洁可执行），`remind_time` 与 `notes` 可选；时间优先 ISO8601 或 `yyyy-MM-dd HH:mm`。**若文中没有此类事项，必须输出** `potential_todos` 为空数组 `[]`，**不得省略该字段**。\n' +
  '6. 你的全部输出必须是**一个合法 JSON 对象**（不要使用 Markdown 代码围栏、不要添加任何解释性前后缀）。结构示例：\n' +
  '   {"content_type":"新闻报道","核心事实":"...","影响范围":"...","关键引述或数据":"...","后续关注点":"...","potential_todos":[{"title":"联系编辑部核实数据","notes":"见第三段"}]}\n' +
  '   其中除 `content_type`、`potential_todos` 外须另有 4～6 个自定义中文要点键。\n' +
  '7. 除该 JSON 外不得输出任何其他字符。';

function stripJsonFence(s) {
  let t = String(s || '').trim();
  if (t.startsWith('```')) {
    t = t.replace(/^```[a-zA-Z0-9]*\s*/i, '').replace(/\s*```\s*$/, '');
  }
  return t.trim();
}

/**
 * @param {string} rawText
 * @returns {object}
 */
export function parsePaperSummaryJson(rawText) {
  const cleaned = stripJsonFence(rawText);
  try {
    return JSON.parse(cleaned);
  } catch {
    const i = cleaned.indexOf('{');
    const j = cleaned.lastIndexOf('}');
    if (i >= 0 && j > i) {
      try {
        return JSON.parse(cleaned.slice(i, j + 1));
      } catch {
        /* fall through */
      }
    }
  }
  throw new Error('无法解析模型返回的 JSON');
}

const PAPER_SUMMARY_RESERVED_KEYS = new Set(['content_type', 'potential_todos']);

/**
 * @param {unknown} raw
 * @returns {{ title: string, remind_time?: string, notes?: string }[]}
 */
function normalizePotentialTodosArray(raw) {
  if (!Array.isArray(raw)) {
    return [];
  }
  const out = [];
  for (const item of raw) {
    if (!item || typeof item !== 'object' || Array.isArray(item)) {
      continue;
    }
    const title = String(item.title || item.content || item.task || '')
      .trim()
      .slice(0, 500);
    if (!title) {
      continue;
    }
    const rt = item.remind_time != null ? String(item.remind_time).trim().slice(0, 120) : '';
    const notes = item.notes != null ? String(item.notes).trim().slice(0, 2000) : '';
    /** @type {{ title: string, remind_time?: string, notes?: string }} */
    const row = { title };
    if (rt) {
      row.remind_time = rt;
    }
    if (notes) {
      row.notes = notes;
    }
    out.push(row);
    if (out.length >= 24) {
      break;
    }
  }
  return out;
}
const PAPER_SUMMARY_LEGACY_KEYS = ['research_question', 'innovation', 'results', 'limitations', 'citations'];

const PAPER_LEGACY_LABELS = {
  research_question: '研究问题',
  innovation: '方法创新',
  results: '实验与结论',
  limitations: '局限性',
  citations: '可引观点',
};

/**
 * 模型无法解析或归一化失败时的兜底结构（前端可安全渲染）
 * @param {string} sourceText
 * @param {string} [title]
 * @returns {{ content_type: string, dimensions: Record<string, string>, potential_todos: Array<{ title: string, remind_time?: string, notes?: string }> }}
 */
export function buildAdaptivePaperSummaryFallback(sourceText, title = '') {
  const raw = String(sourceText || '').replace(/\s+/g, ' ').trim();
  const head = raw.slice(0, 2200);
  const titleLine = String(title || '').trim();
  const prefix =
    '模型输出未能解析为标准 JSON。以下为根据原文开头的摘录式摘要，请核对原文；你可重新点击「生成结构化卡片」重试。\n';
  const hint = titleLine ? `【标题线索】${titleLine}\n` : '';
  const body = head.length ? `${hint}${prefix}${head}` : `${hint}${prefix}（正文为空或不可读）`;
  return {
    content_type: '其他',
    dimensions: {
      内容摘要: body,
    },
    potential_todos: [],
  };
}

/**
 * 将旧版五字段摘要转为自适应 dimensions（兼容历史响应）
 * @param {object} obj
 * @returns {Record<string, string>|null}
 */
function legacyPaperSummaryToDimensions(obj) {
  if (!obj || typeof obj !== 'object') {
    return null;
  }
  const out = {};
  for (const k of PAPER_SUMMARY_LEGACY_KEYS) {
    if (k === 'citations') {
      const arr = Array.isArray(obj.citations) ? obj.citations : [];
      const text = arr.map((x) => String(x || '').trim()).filter(Boolean).join('\n');
      if (text) {
        out[PAPER_LEGACY_LABELS.citations] = text;
      }
      continue;
    }
    const v = String(obj[k] || '').trim();
    if (v) {
      out[PAPER_LEGACY_LABELS[k] || k] = v;
    }
  }
  return Object.keys(out).length ? out : null;
}

/**
 * @param {object} obj
 * @param {string} sourceTextForFallback 用户原文截断片段，用于兜底摘录
 * @returns {{ content_type: string, dimensions: Record<string, string>, potential_todos: { title: string, remind_time?: string, notes?: string }[] }}
 */
export function normalizeAdaptivePaperSummary(obj, sourceTextForFallback) {
  if (!obj || typeof obj !== 'object' || Array.isArray(obj)) {
    throw new Error('摘要数据为空');
  }
  const contentType = String(obj.content_type || '').trim() || '未分类';
  const potentialTodos = normalizePotentialTodosArray(obj.potential_todos);
  /** @type {Record<string, string>} */
  const dimensions = {};

  for (const [rawKey, rawVal] of Object.entries(obj)) {
    const key = String(rawKey || '').trim();
    if (!key || PAPER_SUMMARY_RESERVED_KEYS.has(key)) {
      continue;
    }
    if (typeof rawVal === 'string') {
      const v = rawVal.trim();
      if (v) {
        dimensions[key] = v;
      }
    } else if (typeof rawVal === 'number' && !Number.isNaN(rawVal)) {
      dimensions[key] = String(rawVal);
    } else if (Array.isArray(rawVal)) {
      const joined = rawVal
        .map((x) => String(x || '').trim())
        .filter(Boolean)
        .join('\n');
      if (joined) {
        dimensions[key] = joined;
      }
    }
  }

  let merged = dimensions;
  if (Object.keys(merged).length === 0) {
    const legacy = legacyPaperSummaryToDimensions(obj);
    if (legacy) {
      merged = legacy;
    }
  }

  if (Object.keys(merged).length === 0) {
    const fb = buildAdaptivePaperSummaryFallback(sourceTextForFallback, '');
    const ct = String(obj.content_type || '').trim();
    const base = ct ? { ...fb, content_type: ct } : fb;
    return { ...base, potential_todos: potentialTodos };
  }

  const keys = Object.keys(merged);
  const maxDims = 8;
  if (keys.length > maxDims) {
    const sliced = {};
    for (let i = 0; i < maxDims; i += 1) {
      const k = keys[i];
      sliced[k] = merged[k];
    }
    merged = sliced;
  }

  return { content_type: contentType, dimensions: merged, potential_todos: potentialTodos };
}

/**
 * 文献结构化摘要：非流式 chat/completions，返回规范化 JSON 字段。
 * @param {ModelCallConfig} modelConfig
 * @param {{ text: string, title?: string }} payload
 */
export async function summarizePaper(modelConfig, payload) {
  if (!modelConfig?.api_key?.trim()) {
    throw new Error('模型配置缺少 API Key');
  }
  const rawText = String(payload?.text || '');
  const title = String(payload?.title || '').trim();
  const clipped = rawText.slice(0, PAPER_SUMMARY_MAX_INPUT_CHARS);
  const titleLine = title || '（未提供标题）';

  const userContent =
    `材料标题：${titleLine}\n\n` +
    `以下是正文摘录（若过长已截断至前 ${PAPER_SUMMARY_MAX_INPUT_CHARS} 个字符）。请先判断类型，再输出上文约定的自适应 JSON（含 content_type 与 4～6 个中文维度键）：\n\n` +
    clipped;

  const model = normalizeModelId(modelConfig.model_name) || DEFAULT_MODEL;
  const url = chatCompletionsUrl(modelConfig);
  const headers = buildHeadersFromConfig(modelConfig.api_key);
  if (modelConfig.type === 'aliyun') {
    headers['X-DashScope-SSE'] = 'disable';
  }

  const body = withThinkingParams(
    {
      model,
      messages: [
        { role: 'system', content: SYSTEM_PAPER_REVIEWER },
        { role: 'user', content: userContent },
      ],
      temperature: 0.25,
      stream: false,
      max_tokens: 4096,
    },
    model,
  );

  const data = await requestJson(url, body, headers);
  const raw = extractTextFromResponse(data);
  if (!raw) {
    return buildAdaptivePaperSummaryFallback(clipped, title);
  }
  try {
    const parsed = parsePaperSummaryJson(raw);
    return normalizeAdaptivePaperSummary(parsed, clipped);
  } catch (e) {
    console.warn('[summarizePaper] parse or normalize failed, using fallback', e?.message || e);
    return buildAdaptivePaperSummaryFallback(clipped, title);
  }
}

export function parseSsePayloads(payload) {
  const trimmed = String(payload || '').trim();
  if (!trimmed) {
    return [];
  }

  const results = [];
  try {
    results.push(JSON.parse(trimmed));
    return results;
  } catch {
    // continue to split
  }

  const normalized = trimmed.replace(/}\s*\{/g, '}\n{');
  const parts = normalized.split('\n').filter(Boolean);
  for (const part of parts) {
    try {
      results.push(JSON.parse(part));
    } catch (error) {
      console.warn('[OpenClaw] SSE JSON parse failed', error);
    }
  }

  return results;
}
