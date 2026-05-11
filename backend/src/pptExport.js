import PptxGenJS from 'pptxgenjs';

/**
 * 将 Markdown 风格提纲转为 pptx buffer（简化解析：# / ## / - 列表）
 * @param {string} markdown
 * @param {string} [deckTitleFallback]
 * @returns {Promise<Buffer>}
 */
export async function markdownToPptxBuffer(markdown, deckTitleFallback = '演示文稿') {
  const md = String(markdown || '').trim();
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = '会意助手';

  const lines = md.split(/\r?\n/);
  let deckTitle = deckTitleFallback;
  const slides = [];

  /** @type {{ title: string, bullets: string[] } | null} */
  let current = null;

  function flush() {
    if (current && (current.title || current.bullets.length)) {
      slides.push(current);
    }
    current = null;
  }

  for (const line of lines) {
    const t = line.trim();
    if (!t) {
      continue;
    }
    if (t.startsWith('# ')) {
      flush();
      deckTitle = t.slice(2).trim() || deckTitleFallback;
      continue;
    }
    if (t.startsWith('## ')) {
      flush();
      current = { title: t.slice(3).trim() || '幻灯片', bullets: [] };
      continue;
    }
    if (t.startsWith('- ') || t.startsWith('* ')) {
      if (!current) {
        current = { title: '要点', bullets: [] };
      }
      current.bullets.push(t.replace(/^[-*]\s+/, '').trim());
      continue;
    }
    if (!current) {
      current = { title: '要点', bullets: [] };
    }
    current.bullets.push(t);
  }
  flush();

  const titleSlide = pptx.addSlide();
  titleSlide.addText(deckTitle, {
    x: 0.5,
    y: 2.2,
    w: 12.5,
    h: 1.5,
    fontSize: 36,
    bold: true,
    color: '1e293b',
    align: 'center',
  });
    titleSlide.addText('由会意助手自 Markdown 提纲生成', {
    x: 0.5,
    y: 4,
    w: 12.5,
    h: 0.5,
    fontSize: 14,
    color: '64748b',
    align: 'center',
  });

  for (const s of slides) {
    const slide = pptx.addSlide();
    slide.addText(s.title, {
      x: 0.6,
      y: 0.45,
      w: 12.3,
      h: 0.85,
      fontSize: 26,
      bold: true,
      color: '0f172a',
    });
    const body = s.bullets.length
      ? s.bullets.map((b) => ({ text: b, options: { bullet: true, indentLevel: 0 } }))
      : [{ text: '（无要点）', options: {} }];
    slide.addText(body, {
      x: 0.7,
      y: 1.45,
      w: 11.8,
      h: 5.2,
      fontSize: 16,
      color: '334155',
      valign: 'top',
    });
  }

  if (slides.length === 0) {
    const slide = pptx.addSlide();
    slide.addText('未解析到 ## 分页标题', {
      x: 0.6,
      y: 2,
      w: 12,
      h: 1,
      fontSize: 20,
      color: 'b91c1c',
    });
    slide.addText('请在提纲中使用「## 第 n 页：标题」分段，并用「- 」列出要点。', {
      x: 0.6,
      y: 2.9,
      w: 12,
      h: 2,
      fontSize: 14,
      color: '475569',
    });
  }

  const out = await pptx.write({ outputType: 'nodebuffer' });
  return Buffer.from(out);
}
