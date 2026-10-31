// 申請書 Word 檔共用排版工具：字型、段落、表格、圖表編號、文件組裝與輸出
// 標記語法：**粗體**、[[待填資料（黃色螢光）]]
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Footer, PageNumber, AlignmentType, WidthType, BorderStyle, ShadingType,
  VerticalAlign, TableLayoutType, HeadingLevel, LevelFormat, LineRuleType,
} = require('docx');

const EA = '標楷體';
const LAT = 'Times New Roman';
const FONT = { ascii: LAT, hAnsi: LAT, eastAsia: EA, cs: LAT };
const BODY = 24;          // 12pt
const TBL = 21;           // 10.5pt
const TEXT_W = 9638;      // A4 寬 11906 − 左右邊界各 1134（2 cm）
const FIG_DIR = path.join(__dirname, 'figs');

const LINE = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const BORDERS = { top: LINE, bottom: LINE, left: LINE, right: LINE, insideHorizontal: LINE, insideVertical: LINE };
const HEAD_FILL = 'DCE6F2';
const SUB_FILL = 'F2F2F2';
const CEN = AlignmentType.CENTER;
const RIGHT = AlignmentType.RIGHT;

// lineScale：固定行高之整體縮放比例
// tables／figures：圖表鍵值，依出現順序自動編號；prefix：編號前綴（如「6-」產生「表 6-1」）
function createKit({ lineScale = 1, tables = [], figures = [], prefix = '' } = {}) {
  // ---------- 圖表編號（內文引用與標題一致） ----------
  const num = (order, k) => {
    const i = order.indexOf(k);
    if (i < 0) throw new Error(`unknown key ${k}`);
    return `${prefix}${i + 1}`;
  };
  const T = (k) => num(tables, k);
  const F = (k) => num(figures, k);
  const seenTables = [];
  const seenFigs = [];
  function checkNumbering() {
    if (JSON.stringify(seenTables) !== JSON.stringify(tables)) {
      throw new Error(`table order mismatch:\n${seenTables.join(',')}\n${tables.join(',')}`);
    }
    if (JSON.stringify(seenFigs) !== JSON.stringify(figures)) {
      throw new Error(`figure order mismatch: ${seenFigs.join(',')}`);
    }
  }

  // ---------- 文字與段落 ----------
  function tr(t, o = {}) {
    return new TextRun({ text: t, font: FONT, size: o.size || BODY, bold: o.bold, highlight: o.highlight, color: o.color });
  }
  function runs(text, base = {}) {
    const out = [];
    const re = /(\*\*[^*]+\*\*|\[\[[^\]]+\]\])/g;
    let last = 0;
    let m;
    while ((m = re.exec(text))) {
      if (m.index > last) out.push(tr(text.slice(last, m.index), base));
      const tok = m[0];
      if (tok.startsWith('**')) out.push(tr(tok.slice(2, -2), { ...base, bold: true }));
      else out.push(tr(tok.slice(2, -2), { ...base, highlight: 'yellow' }));
      last = m.index + tok.length;
    }
    if (last < text.length) out.push(tr(text.slice(last), base));
    return out;
  }
  const exact = (pt) => ({ line: Math.round(pt * 20 * lineScale), lineRule: LineRuleType.EXACT });

  function H1(text) {
    return new Paragraph({ heading: HeadingLevel.HEADING_1, keepNext: true, children: runs(text, { size: 32, bold: true }) });
  }
  function H2(text) {
    return new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, children: runs(text, { size: 28, bold: true }) });
  }
  function H3(text, opts = {}) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_3, keepNext: true,
      indent: opts.hang ? { left: opts.hang, hanging: opts.hang } : undefined,
      children: runs(text, { size: opts.size || BODY, bold: true }),
    });
  }
  function P(text, opts = {}) {
    return new Paragraph({
      alignment: AlignmentType.BOTH,
      spacing: { ...exact(20), before: opts.before || 0, after: opts.after || 40 },
      indent: { firstLine: opts.noIndent ? 0 : 480, left: opts.left || 0 },
      keepNext: opts.keepNext,
      children: runs(text, { size: opts.size }),
    });
  }
  // 編號項目：label 如 "1."、"(1)"；以 tab 對齊懸掛縮排
  function N(label, text, opts = {}) {
    const left = opts.left || 840;
    const hang = opts.hang || 360;
    return new Paragraph({
      alignment: AlignmentType.BOTH,
      spacing: { ...exact(20), after: 40 },
      indent: { left, hanging: hang },
      children: [tr(label + '\t'), ...runs(text)],
    });
  }
  const items = (arr, opts) => arr.map((t, i) => N(`${i + 1}.`, t, opts));
  const pitems = (arr) => arr.map((t, i) => N(`(${i + 1})`, t, { left: 960, hang: 480 }));
  // 小節標示（➢）
  function ARROW(text) {
    return new Paragraph({
      numbering: { reference: 'arrow', level: 0 },
      keepNext: true,
      spacing: { ...exact(20), before: 60, after: 40 },
      children: runs(text, { bold: true }),
    });
  }
  function CAP(text, opts = {}) {
    return new Paragraph({
      alignment: AlignmentType.CENTER,
      keepNext: !!opts.keepNext,
      spacing: { before: opts.before ?? 100, after: opts.after ?? 60, ...exact(16) },
      children: runs(text, { size: 22, bold: true }),
    });
  }
  function TCAP(key, title, opts = {}) {
    seenTables.push(key);
    return CAP(`表 ${T(key)}　${title}`, { keepNext: true, ...opts });
  }
  function NOTE(text) {
    return new Paragraph({
      spacing: { before: 40, after: 100, ...exact(15) },
      children: runs(text, { size: 20 }),
    });
  }
  function FIGURE(file, key, title, widthPx = 640) {
    seenFigs.push(key);
    const caption = `圖 ${F(key)}　${title}`;
    const buf = fs.readFileSync(path.join(FIG_DIR, file));
    const w = buf.readUInt32BE(16);
    const h = buf.readUInt32BE(20);
    return [
      new Paragraph({
        alignment: AlignmentType.CENTER, keepNext: true,
        spacing: { before: 120, after: 40 },
        children: [new ImageRun({
          type: 'png', data: buf,
          transformation: { width: widthPx, height: Math.round(widthPx * h / w) },
          altText: { title: caption, description: caption, name: file },
        })],
      }),
      CAP(caption, { before: 0, after: 140 }),
    ];
  }

  // ---------- 表格 ----------
  function cellParas(content, o) {
    const arr = Array.isArray(content) ? content : [content];
    const size = o.size || TBL;
    return arr.map((t) => {
      const s = String(t);
      const m = /^(\d+\.)\s(.*)$/.exec(s);   // 儲存格內「1. 」開頭者採懸掛縮排
      const hang = m && m[1].length > 2 ? 340 : 260;
      return new Paragraph({
        alignment: o.align || AlignmentType.LEFT,
        spacing: { ...exact(size === 18 ? 12 : size === 20 ? 14 : 15.5), before: 0, after: 0 },
        indent: m ? { left: hang, hanging: hang } : undefined,
        keepNext: o.keepNext,
        children: m ? [tr(m[1] + '\t', { size }), ...runs(m[2], { size, bold: o.bold })] : runs(s, { size, bold: o.bold }),
      });
    });
  }
  // spec: 字串／字串陣列，或 { t, span, rowSpan, fill, align, bold, vAlign }
  function C(spec, width, o = {}) {
    const s = (typeof spec === 'object' && !Array.isArray(spec)) ? spec : { t: spec };
    return new TableCell({
      width: { size: width, type: WidthType.DXA },
      columnSpan: s.span,
      rowSpan: s.rowSpan,
      verticalAlign: s.vAlign || o.vAlign || VerticalAlign.CENTER,
      shading: (s.fill || o.fill) ? { fill: s.fill || o.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
      margins: { top: o.padY ?? 50, bottom: o.padY ?? 50, left: o.padX ?? 85, right: o.padX ?? 85 },
      children: cellParas(s.t, { align: s.align || o.align, bold: s.bold ?? o.bold, size: o.size, keepNext: o.keepNext }),
    });
  }
  // cols: [{ w, h(表頭), align }]；rows: 每列為 cell spec 陣列（合併欄位時省略被合併者）
  // o.keep：整張表不跨頁（除最後一列外皆設 keepNext）
  function TABLE(cols, rows, o = {}) {
    const widths = cols.map((c) => c.w);
    const sum = widths.reduce((a, b) => a + b, 0);
    if (sum !== TEXT_W) throw new Error(`table width ${sum} != ${TEXT_W}`);
    const header = new TableRow({
      tableHeader: true, cantSplit: true,
      children: cols.map((c) => C(c.h, c.w, { fill: HEAD_FILL, bold: true, align: AlignmentType.CENTER, size: o.size, padX: o.padX, keepNext: true })),
    });
    const body = rows.map((r, ri) => {
      let ci = 0;
      const cells = r.map((spec) => {
        const s = (typeof spec === 'object' && !Array.isArray(spec)) ? spec : { t: spec };
        const span = s.span || 1;
        const w = widths.slice(ci, ci + span).reduce((a, b) => a + b, 0);
        const align = s.align || cols[ci].align;
        ci += span;
        return C(s, w, { align, size: o.size, padX: o.padX, keepNext: o.keep && ri < rows.length - 1 });
      });
      return new TableRow({ cantSplit: o.cantSplit !== false, children: cells });
    });
    return new Table({
      width: { size: TEXT_W, type: WidthType.DXA },
      columnWidths: widths,
      layout: TableLayoutType.FIXED,
      borders: BORDERS,
      rows: o.noHeader ? body : [header, ...body],
    });
  }

  return {
    T, F, checkNumbering,
    tr, runs, exact, H1, H2, H3, P, N, items, pitems, ARROW, CAP, TCAP, NOTE, FIGURE,
    C, TABLE,
  };
}

// ---------- 文件組裝與輸出 ----------
function makeDocument({ title, description, children }) {
  return new Document({
    creator: '國立高雄師範大學軟體工程與管理學系',
    title,
    description,
    styles: {
      default: { document: { run: { font: FONT, size: BODY } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { font: FONT, size: 32, bold: true, color: '000000' },
          paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 0, keepNext: true } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { font: FONT, size: 28, bold: true, color: '000000' },
          paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 1, keepNext: true } },
        { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { font: FONT, size: 24, bold: true, color: '000000' },
          paragraph: { spacing: { before: 140, after: 60 }, outlineLevel: 2, keepNext: true } },
      ],
    },
    numbering: {
      config: [{
        reference: 'arrow',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 480, hanging: 360 } }, run: { font: 'Wingdings' } },
        }],
      }],
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1134, bottom: 1134, left: 1134, right: 1134, header: 567, footer: 567 },
        },
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20 })] })],
        }),
      },
      children,
    }],
  });
}

// docx 套件會在螢光標示外另寫入 <w:highlightCs>，不符 OOXML 綱要，輸出前移除
// （jszip 為 docx 之相依套件，自 docx 所在位置解析）
function writeDocx(doc, out) {
  const JSZip = require(require.resolve('jszip', { paths: [path.dirname(require.resolve('docx'))] }));
  return Packer.toBuffer(doc)
    .then((buf) => JSZip.loadAsync(buf))
    .then(async (zip) => {
      const xml = await zip.file('word/document.xml').async('string');
      zip.file('word/document.xml', xml.replace(/<w:highlightCs\b[^>]*\/>/g, ''));
      return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    })
    .then((buf) => {
      fs.writeFileSync(out, buf);
      console.log('wrote', out, buf.length, 'bytes');
    });
}

module.exports = {
  createKit, makeDocument, writeDocx,
  FONT, BODY, TBL, TEXT_W, BORDERS, HEAD_FILL, SUB_FILL, CEN, RIGHT,
};
