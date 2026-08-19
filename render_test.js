const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

async function convertMdToPdf(mdPath, outputPath) {
  const mdContent = fs.readFileSync(mdPath, 'utf8');
  
  // 提取大标题
  const titleMatch = mdContent.match(/^#\s+(.+)$/m);
  const title = titleMatch ? titleMatch[1] : '产品规划';
  
  // 提取引用元数据块
  const quoteLines = [];
  const lines = mdContent.split('\n');
  for (const line of lines) {
    if (line.trim().startsWith('>')) {
      quoteLines.push(line.trim().replace(/^>\s*/, ''));
    } else if (quoteLines.length > 0 && line.trim().startsWith('##')) {
      break;
    }
  }

  const kvPairs = [];
  for (const q of quoteLines) {
    const m = q.match(/^\*{0,2}([\u4e00-\u9fa5a-zA-Z0-9_（）\s]{2,12})\*{0,2}[：:]\s*(.+)$/);
    if (m) {
      kvPairs.push([m[1].trim(), m[2].trim()]);
    }
  }

  // 渲染 2 列元数据表格 HTML (同 Word 版)
  let tableHtml = '<table class="cover-table">';
  for (const [k, v] of kvPairs) {
    const formattedV = v.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
    tableHtml += `<tr><td class="cover-key">${k}</td><td class="cover-val">${formattedV}</td></tr>`;
  }
  tableHtml += '</table>';

  // 提取正文并去除头部 H1 与 blockquote 与 ---
  let bodyMd = mdContent.replace(/^#\s+.+$/m, '');
  for (const q of quoteLines) {
    bodyMd = bodyMd.replace(new RegExp('>\\s*' + q.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'm'), '');
  }
  bodyMd = bodyMd.replace(/^\s*-{3,}\s*$/m, '');

  const html = `<!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <script src="https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
    <style>
      @page {
        size: A4;
        margin: 20mm 20mm 20mm 20mm;
      }
      body {
        font-family: -apple-system, BlinkMacSystemFont, 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', sans-serif;
        color: #1F2937;
        font-size: 10.5pt;
        line-height: 1.6;
        margin: 0;
        padding: 0;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }
      
      /* 独立封面页 */
      .cover-page {
        page-break-after: always;
        break-after: page;
        padding-top: 60px;
        box-sizing: border-box;
      }
      .cover-title {
        text-align: center;
        font-size: 24pt;
        font-weight: bold;
        color: #0C1A32;
        margin-bottom: 50px;
        line-height: 1.3;
      }
      .cover-table {
        width: 100%;
        border-collapse: collapse;
        border: 1px solid #CBD5E1;
        margin-top: 20px;
      }
      .cover-key {
        width: 25%;
        background-color: #F1F5F9;
        color: #0C1A32;
        font-weight: bold;
        text-align: center;
        vertical-align: middle;
        padding: 12px 14px;
        border: 1px solid #CBD5E1;
        font-size: 10pt;
      }
      .cover-val {
        width: 75%;
        background-color: #FFFFFF;
        color: #1F2937;
        padding: 12px 16px;
        border: 1px solid #CBD5E1;
        font-size: 10pt;
        line-height: 1.5;
      }
      
      /* 正文页排版 */
      .content-body {
        padding-top: 0;
      }
      h2 {
        color: #0C1A32;
        font-size: 15pt;
        font-weight: bold;
        margin-top: 24px;
        margin-bottom: 8px;
        page-break-after: avoid;
        break-after: avoid;
      }
      h3 {
        color: #1E293B;
        font-size: 12.5pt;
        font-weight: bold;
        margin-top: 18px;
        margin-bottom: 6px;
        page-break-after: avoid;
        break-after: avoid;
      }
      p {
        margin: 6px 0;
      }
      ul, ol {
        margin: 6px 0;
        padding-left: 24px;
      }
      li {
        margin: 4px 0;
      }
      table:not(.cover-table) {
        width: 100%;
        border-collapse: collapse;
        margin: 16px 0;
        page-break-inside: avoid;
        break-inside: avoid;
      }
      table:not(.cover-table) th, table:not(.cover-table) td {
        border: 1px solid #CBD5E1;
        padding: 8px 12px;
        font-size: 9.5pt;
      }
      table:not(.cover-table) th {
        background-color: #0C1A32;
        color: #FFFFFF;
        font-weight: bold;
        text-align: left;
      }
      table:not(.cover-table) tr:nth-child(even) td {
        background-color: #F8FAFC;
      }
      .mermaid {
        display: flex;
        justify-content: center;
        align-items: center;
        margin: 20px auto;
        page-break-inside: avoid;
        break-inside: avoid;
      }
      .mermaid svg {
        max-height: 18cm !important;
        max-width: 100% !important;
      }
    </style>
  </head>
  <body>
    <div class="cover-page">
      <div class="cover-title">${title}</div>
      ${tableHtml}
    </div>
    <div class="content-body" id="content"></div>

    <script>
      const rawMd = ${JSON.stringify(bodyMd)};
      document.getElementById('content').innerHTML = marked.parse(rawMd);
      
      // 渲染 Mermaid 代码块
      document.querySelectorAll('pre code.language-mermaid').forEach((el) => {
        const pre = el.parentElement;
        const div = document.createElement('div');
        div.className = 'mermaid';
        div.textContent = el.textContent;
        pre.parentNode.replaceChild(div, pre);
      });
      
      mermaid.initialize({ startOnLoad: true, theme: 'neutral' });
    </script>
  </body>
  </html>`;

  const browser = await chromium.launch({
    executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    headless: true
  });
  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: 'networkidle' });
  // 等待 Mermaid 渲染完成
  await page.waitForSelector('.mermaid svg', { timeout: 5000 }).catch(() => {});
  await page.waitForTimeout(1000);
  
  await page.pdf({
    path: outputPath,
    format: 'A4',
    margin: { top: '20mm', bottom: '20mm', left: '20mm', right: '20mm' },
    printBackground: true
  });
  
  await browser.close();
  console.log('Successfully generated standalone PDF:', outputPath);
}

convertMdToPdf(
  '/Users/alanqin/Library/Mobile Documents/iCloud~md~obsidian/Documents/notes/Work/产品相关/产品规划-03-墨爻阁(创作者展厅与作品名片).md',
  '/Users/alanqin/.gemini/antigravity/brain/dd9e000b-f46c-4612-80c3-7170eb67a30e/scratch/standalone_03.pdf'
).catch(err => {
  console.error(err);
  process.exit(1);
});
