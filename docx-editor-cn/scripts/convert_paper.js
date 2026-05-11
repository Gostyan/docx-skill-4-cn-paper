/**
 * convert_paper.js — 基于 Q-Learning 算法的走迷宫智能体复现与分析
 *
 * All formatting follows Chinese academic standards:
 *   - A4 page, 2.5 cm margins on all sides
 *   - SimSun 12pt body, SimHei headings (Cambria Math for English/numbers)
 *   - Manual multi-level heading numbering
 *   - Three-line table (三线表)
 *   - LaTeX formula via temml + Word native math
 *   - Citation superscript handling [n] format
 *
 * Usage:
 *   node scripts/convert_paper.js
 *
 * Dependencies:
 *   npm install docx temml fast-xml-parser
 */

'use strict';

const fs   = require('fs');
const path = require('path');

const {
  Document, Packer,
  Paragraph, TextRun, Math, MathRun,
  Table, TableRow, TableCell,
  Header, Footer,
  PageNumber, AlignmentType, LineRuleType, HeadingLevel,
  LevelFormat, LevelSuffix, BorderStyle, WidthType, ShadingType, VerticalAlign,
  TableOfContents, PageBreak, ImageRun, SequentialIdentifier,
} = require('docx');

const { mathmlToDocxChildren } = require('./mathml-to-docx');
const temml = require('temml');

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

const OUTPUT_PATH   = path.resolve(__dirname, '..', 'output.docx');
const IMG_BASE      = path.resolve(__dirname, '..', '..', 'example', 'markdown论文', 'images');

// Page / margin (DXA)
const PAGE_W        = 11906;   // A4
const PAGE_H        = 16838;
const MARGIN        = 1418;    // 2.5 cm
const CONTENT_W     = PAGE_W - 2 * MARGIN;  // 9070 DXA

// Three-line table border presets
const THICK = { style: BorderStyle.SINGLE, size: 12, color: '000000' }; // 1.5 pt
const THIN  = { style: BorderStyle.SINGLE, size: 6,  color: '000000' }; // 0.75 pt
const NONE  = { style: BorderStyle.NONE,   size: 0,  color: 'FFFFFF' };

let _chapter = 0;

// ─────────────────────────────────────────────────────────────────────────────
// Unicode Math Symbols to LaTeX Mapping
// ─────────────────────────────────────────────────────────────────────────────

const UNICODE_TO_LATEX = {
  'α': '\\alpha', 'β': '\\beta', 'γ': '\\gamma', 'δ': '\\delta', 'ε': '\\varepsilon',
  'ζ': '\\zeta', 'η': '\\eta', 'θ': '\\theta', 'ι': '\\iota', 'κ': '\\kappa',
  'λ': '\\lambda', 'μ': '\\mu', 'ν': '\\nu', 'ξ': '\\xi', 'π': '\\pi',
  'ρ': '\\rho', 'σ': '\\sigma', 'τ': '\\tau', 'υ': '\\upsilon', 'φ': '\\phi',
  'χ': '\\chi', 'ψ': '\\psi', 'ω': '\\omega',
  'Γ': '\\Gamma', 'Δ': '\\Delta', 'Θ': '\\Theta', 'Λ': '\\Lambda', 'Ξ': '\\Xi',
  'Π': '\\Pi', 'Σ': '\\Sigma', 'Φ': '\\Phi', 'Ψ': '\\Psi', 'Ω': '\\Omega',
  '₀': '_0', '₁': '_1', '₂': '_2', '₃': '_3', '₄': '_4',
  '₅': '_5', '₆': '_6', '₇': '_7', '₈': '_8', '₉': '_9',
  'ₐ': '_a', 'ₑ': '_e', 'ₕ': '_h', 'ᵢ': '_i', 'ₖ': '_k',
  'ₗ': '_l', 'ₘ': '_m', 'ₙ': '_n', 'ₒ': '_o', 'ₚ': '_p',
  'ᵣ': '_r', 'ₛ': '_s', 'ₜ': '_t', 'ᵤ': '_u', 'ᵥ': '_v',
  'ₓ': '_x', 'ᵧ': '_y',
  '⁰': '^0', '¹': '^1', '²': '^2', '³': '^3', '⁴': '^4',
  '⁵': '^5', '⁶': '^6', '⁷': '^7', '⁸': '^8', '⁹': '^9',
  'ⁿ': '^n', 'ⁱ': '^i',
  '∞': '\\infty', '∑': '\\sum', '∏': '\\prod', '∫': '\\int',
  '≤': '\\leq', '≥': '\\geq', '≠': '\\neq', '≈': '\\approx',
  '→': '\\to', '←': '\\leftarrow', '↔': '\\leftrightarrow',
  '∈': '\\in', '∉': '\\notin', '⊂': '\\subset', '⊃': '\\supset',
  '∀': '\\forall', '∃': '\\exists', '∧': '\\land', '∨': '\\lor',
  '×': '\\times', '÷': '\\div', '±': '\\pm', '∓': '\\mp',
  '·': '\\cdot', '…': '\\ldots', '⋯': '\\cdots',
  '′': "'", '″': "''",
  '⟨': '\\langle ', '⟩': ' \\rangle',
  '*': '^*',
};

const GREEK_CHARS   = 'αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ';
const SUB_SUP_CHARS = '₀₁₂₃₄₅₆₇₈₉ₐₑₕᵢₖₗₘₙₒₚᵣₛₜᵤᵥₓᵧ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ';
const MATH_OPS      = '∞∑∏∫≤≥≠≈→←↔∈∉⊂⊃∀∃∧∨×÷±∓·…⋯′″⟨⟩';
const GREEK_CLASS   = new RegExp(`[${GREEK_CHARS}]`);
const SUB_SUP_CLASS = new RegExp(`[${SUB_SUP_CHARS}]`);
const MATH_OPS_CLASS = new RegExp(`[${MATH_OPS}]`);

function unicodeToLatex(text) {
  let result = text.replace(/([A-Za-z])̄/g, '\\bar{$1}');
  result = result.replace(/(\^|_)(?!\{)([A-Za-z0-9]+)/g, '$1{$2}');
  for (const [unicode, latex] of Object.entries(UNICODE_TO_LATEX)) {
    result = result.split(unicode).join(latex);
  }
  return result;
}

function findBracedExtent(text, pos) {
  let start = pos - 1;
  while (start >= 0 && /[A-Za-z0-9]/.test(text[start])) start--;
  start++;
  const braceOpen = pos + 1;
  if (braceOpen >= text.length || text[braceOpen] !== '{') return null;
  let depth = 1;
  let i = braceOpen + 1;
  while (i < text.length && depth > 0) {
    if (text[i] === '{') depth++;
    if (text[i] === '}') depth--;
    i++;
  }
  if (depth !== 0) return null;
  return { start, end: i };
}

function containsMath(text) {
  if (GREEK_CLASS.test(text)) return true;
  if (SUB_SUP_CLASS.test(text)) return true;
  if (MATH_OPS_CLASS.test(text)) return true;
  if (/[A-Z]\*/.test(text)) return true;
  if (/\$[^$]+\$/.test(text)) return true;
  if (/[_^]\{[^}]+\}/.test(text)) return true;
  if (/[A-Za-z]\d*[_^][a-zA-Z0-9]/.test(text)) return true;
  if (/\\[a-zA-Z]/.test(text)) return true;
  if (/[A-Za-z]̄/.test(text)) return true;
  return false;
}

function containsCitation(text) {
  return /\[\d+\]/.test(text);
}

function parseInlineContentWithCitations(text) {
  const combinedPattern = new RegExp(
    `(\\[\\d+\\])` +
    `|\\$([^$]+)\\$` +
    `|(\\\\[a-zA-Z]+\\{[^}]*\\}[ \\t]*\\{[^}]*\\})` +
    `|(\\\\[a-zA-Z]+\\{[^}]*\\})` +
    `|(\\\\[a-zA-Z]+)` +
    `|([A-Za-z${GREEK_CHARS}]+[_^]\\{[^}]+\\})` +
    `|([A-Za-z]+\\d*[_^][a-zA-Z0-9]+\\s*\\([^)]+\\))` +
    `|([A-Za-z]+\\d*[_^][a-zA-Z0-9]+)` +
    `|([A-Z][${SUB_SUP_CHARS}]+\\*?\\s*\\([^)]+\\))` +
    `|([A-Z]\\*\\s*\\([^)]+\\))` +
    `|([A-Z]\\s*\\([^)]*[${GREEK_CHARS}${SUB_SUP_CHARS}][^)]*\\))` +
    `|([${GREEK_CHARS}][${SUB_SUP_CHARS}]*\\*?)` +
    `|([A-Za-z]+[${SUB_SUP_CHARS}]+\\*?)` +
    `|([A-Z]\\*)` +
    `|([${MATH_OPS}])`,
    'g');

  // Find nested-brace expressions (_{...{...}...} and ^{...{...}...})
  const nested = [];
  const trigger = /[_^]\{/g;
  let t;
  while ((t = trigger.exec(text)) !== null) {
    const ext = findBracedExtent(text, t.index);
    if (ext && ext.start < t.index && ext.end - ext.start > t[0].length + 1) {
      nested.push({ start: ext.start, end: ext.end, content: text.slice(ext.start, ext.end) });
    }
  }

  const children = [];
  let lastIndex = 0;
  let m;
  while ((m = combinedPattern.exec(text)) !== null) {
    if (nested.some(n => m.index + m[0].length > n.start && m.index < n.end)) continue;
    if (m.index > lastIndex) {
      for (const n of nested) {
        if (n.start >= lastIndex && n.start < m.index) {
          if (n.start > lastIndex) children.push(new TextRun(text.slice(lastIndex, n.start)));
          const latex = unicodeToLatex(n.content);
          try {
            const ml = temml.renderToString(latex, { displayMode: false, throwOnError: false });
            const kids = mathmlToDocxChildren(ml);
            children.push(new Math({ children: kids && kids.length ? kids : [new MathRun(n.content)] }));
          } catch (e) { children.push(new Math({ children: [new MathRun(n.content)] })); }
          lastIndex = n.end;
        }
      }
      if (m.index > lastIndex) children.push(new TextRun(text.slice(lastIndex, m.index)));
    }
    if (m[1]) {
      children.push(new TextRun({ text: m[1], superScript: true }));
    } else {
      const mc = m[2] || m[3] || m[4] || m[5] || m[6] || m[7] || m[8] || m[9] || m[10] || m[11] || m[12] || m[13] || m[14] || m[15];
      if (mc) {
        const latex = unicodeToLatex(mc);
        try {
          const ml = temml.renderToString(latex, { displayMode: false, throwOnError: false });
          const kids = mathmlToDocxChildren(ml);
          children.push(new Math({ children: kids && kids.length ? kids : [new MathRun(mc)] }));
        } catch (e) { children.push(new Math({ children: [new MathRun(mc)] })); }
      }
    }
    lastIndex = m.index + m[0].length;
  }
  if (lastIndex < text.length) {
    for (const n of nested) {
      if (n.start >= lastIndex) {
        if (n.start > lastIndex) children.push(new TextRun(text.slice(lastIndex, n.start)));
        const latex = unicodeToLatex(n.content);
        try {
          const ml = temml.renderToString(latex, { displayMode: false, throwOnError: false });
          const kids = mathmlToDocxChildren(ml);
          children.push(new Math({ children: kids && kids.length ? kids : [new MathRun(n.content)] }));
        } catch (e) { children.push(new Math({ children: [new MathRun(n.content)] })); }
        lastIndex = n.end;
      }
    }
    if (lastIndex < text.length) children.push(new TextRun(text.slice(lastIndex)));
  }
  if (children.length === 0) children.push(new TextRun(text));
  return children;
}

// ─────────────────────────────────────────────────────────────────────────────
// Helper Functions
// ─────────────────────────────────────────────────────────────────────────────

function body(text) {
  if (containsMath(text) || containsCitation(text)) {
    return new Paragraph({ children: parseInlineContentWithCitations(text) });
  }
  return new Paragraph({ children: [new TextRun(text)] });
}

function bodyMulti(runs) {
  return new Paragraph({ children: runs });
}

function h1Chinese(text) {
  _chapter++;
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    numbering: { reference: `sections_c${_chapter}`, level: 0 },
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    numbering: { reference: `sections_c${_chapter}`, level: 1 },
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}

function figCaption(text) {
  return new Paragraph({
    style: 'FigureCaption',
    children: [
      new TextRun(`图 ${_chapter}-`),
      new SequentialIdentifier(`figure_c${_chapter}`),
      new TextRun(` ${text}`),
    ],
  });
}

function tableCaption(text) {
  return new Paragraph({
    style: 'TableCaption',
    children: [
      new TextRun(`表 ${_chapter}-`),
      new SequentialIdentifier(`table_c${_chapter}`),
      new TextRun(` ${text}`),
    ],
  });
}

function threeLineTable(headers, rows, colWidths) {
  const n = headers.length;
  if (!colWidths) {
    const w = Math.floor(CONTENT_W / n);
    colWidths = Array(n).fill(w);
    colWidths[n - 1] = CONTENT_W - w * (n - 1);
  }
  if (colWidths.length !== n) throw new Error('colWidths length must match headers length');

  const cellOf = (text, w, borders, bold = false) => {
    let cellChildren;
    if (containsMath(text)) {
      cellChildren = parseInlineContentWithCitations(text);
      if (bold) {
        cellChildren = cellChildren.map(child => {
          if (child instanceof TextRun) {
            return new TextRun({ text: child.options?.text || '', bold: true });
          }
          return child;
        });
      }
    } else {
      cellChildren = [new TextRun({ text, bold })];
    }
    return new TableCell({
      width:   { size: w, type: WidthType.DXA },
      borders,
      shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        indent:    { firstLine: 0 },
        children:  cellChildren,
      })],
    });
  };

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) =>
      cellOf(h, colWidths[i], { top: THICK, bottom: THIN, left: NONE, right: NONE }, true)
    ),
  });

  const bodyRows = rows.map((row, ri) => {
    const isLast = ri === rows.length - 1;
    return new TableRow({
      children: row.map((cell, i) =>
        cellOf(String(cell), colWidths[i], {
          top:    NONE,
          bottom: isLast ? THICK : NONE,
          left:   NONE,
          right:  NONE,
        })
      ),
    });
  });

  return new Table({
    width:        { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: colWidths,
    rows:         [headerRow, ...bodyRows],
  });
}

function formula(latex, number) {
  const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE };
  let mathObj;
  try {
    const mathml = temml.renderToString(latex, { displayMode: true, throwOnError: false });
    const children = mathmlToDocxChildren(mathml);
    if (children && children.length) {
      mathObj = new Math({ children });
    } else {
      mathObj = new Math({ children: [new MathRun(latex)] });
    }
  } catch (e) {
    console.warn(`[formula] LaTeX parse error: ${latex}`, e.message);
    mathObj = new Math({ children: [new MathRun(latex)] });
  }

  const leftCell = new TableCell({
    width: { size: 567, type: WidthType.DXA },
    borders: noBorders,
    shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ indent: { firstLine: 0 }, children: [] })],
  });

  const formulaCell = new TableCell({
    width: { size: 7936, type: WidthType.DXA },
    borders: noBorders,
    shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      indent: { firstLine: 0 },
      children: [mathObj],
    })],
  });

  const numberCell = new TableCell({
    width: { size: 567, type: WidthType.DXA },
    borders: noBorders,
    shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      indent: { firstLine: 0 },
      children: [new TextRun(`(${number})`)],
    })],
  });

  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [567, 7936, 567],
    borders: {
      top: NONE, bottom: NONE, left: NONE, right: NONE,
      insideHorizontal: NONE, insideVertical: NONE,
    },
    rows: [new TableRow({ children: [leftCell, formulaCell, numberCell] })],
  });
}

function ref(text) {
  return new Paragraph({
    style: 'Reference',
    numbering: { reference: 'references', level: 0 },
    children: [new TextRun(text)],
  });
}

function blank() {
  return new Paragraph({ children: [] });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function image(filename, w, h) {
  const imgPath = path.join(IMG_BASE, filename);
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    indent: { firstLine: 0 },
    children: [new ImageRun({
      type: 'jpg',
      data: fs.readFileSync(imgPath),
      transformation: { width: w || 400, height: h || 300 },
      altText: { title: filename, description: filename, name: filename },
    })],
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Styles & Numbering
// ─────────────────────────────────────────────────────────────────────────────

const STYLES = {
  default: {
    document: {
      run: {
        font: { ascii: 'Cambria Math', hAnsi: 'Cambria Math', eastAsia: 'SimSun' },
        size: 24,
      },
      paragraph: {
        spacing: { line: 240, lineRule: LineRuleType.AUTO },
        indent:  { firstLine: 480 },
      },
    },
  },
  paragraphStyles: [
    {
      id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' }, size: 32, bold: true },
      paragraph: {
        alignment:    AlignmentType.CENTER,
        indent:       { firstLine: 0 },
        spacing:      { line: 288, lineRule: LineRuleType.AUTO },
        outlineLevel: 0,
      },
    },
    {
      id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' }, size: 28, bold: true },
      paragraph: {
        alignment:    AlignmentType.LEFT,
        indent:       { firstLine: 0 },
        spacing:      { line: 360, lineRule: LineRuleType.AUTO },
        outlineLevel: 1,
      },
    },
    {
      id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' }, size: 24, bold: true },
      paragraph: {
        alignment:    AlignmentType.LEFT,
        indent:       { firstLine: 0 },
        spacing:      { line: 264, lineRule: LineRuleType.AUTO },
        outlineLevel: 2,
      },
    },
    {
      id: 'FigureCaption', name: 'Figure Caption', basedOn: 'Normal',
      run: { font: { ascii: 'Cambria Math', eastAsia: 'SimSun', hAnsi: 'Cambria Math' }, size: 22, bold: true },
      paragraph: {
        alignment: AlignmentType.CENTER,
        indent:    { firstLine: 0 },
        spacing:   { before: 120, after: 60, line: 240, lineRule: LineRuleType.AUTO },
      },
    },
    {
      id: 'TableCaption', name: 'Table Caption', basedOn: 'Normal',
      run: { font: { ascii: 'Cambria Math', eastAsia: 'SimSun', hAnsi: 'Cambria Math' }, size: 22, bold: true },
      paragraph: {
        alignment: AlignmentType.CENTER,
        indent:    { firstLine: 0 },
        spacing:   { before: 120, after: 60, line: 240, lineRule: LineRuleType.AUTO },
      },
    },
    {
      id: 'Reference', name: 'Reference', basedOn: 'Normal',
      run: { font: { ascii: 'Cambria Math', hAnsi: 'Cambria Math', eastAsia: 'SimSun' }, size: 24 },
      paragraph: {
        spacing: { line: 240, lineRule: LineRuleType.AUTO },
        indent:  { left: 480, hanging: 480, firstLine: 0 },
      },
    },
  ],
};

function buildNumberingConfig(chapterCount) {
  const configs = [
    { reference: 'references', levels: [
      { level: 0, format: LevelFormat.DECIMAL, text: '[%1]',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 480, hanging: 480 } } } }] },
    { reference: 'bullets', levels: [
      { level: 0, format: LevelFormat.BULLET, text: '•',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: 'numbers', levels: [
      { level: 0, format: LevelFormat.DECIMAL, text: '%1.',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ];
  for (let c = 1; c <= chapterCount; c++) {
    configs.push({
      reference: `sections_c${c}`,
      levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: `${c}.%1`,
          suffix: LevelSuffix.SPACE, alignment: AlignmentType.LEFT },
        { level: 1, format: LevelFormat.DECIMAL, text: `${c}.%1.%2`,
          suffix: LevelSuffix.SPACE, alignment: AlignmentType.LEFT },
      ],
    });
  }
  return { config: configs };
}

const NUMBERING = buildNumberingConfig(3);

// ─────────────────────────────────────────────────────────────────────────────
// CONTENT
// ─────────────────────────────────────────────────────────────────────────────

const CONTENT = [

  // ── Title ─────────────────────────────────────────────────────────────────
  blank(), blank(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing:   { before: 0, after: 120 },
    indent:    { firstLine: 0 },
    children:  [new TextRun({ text: '基于 Q-Learning 算法的走迷宫智能体复现与分析', bold: true, size: 36,
                              font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' } })],
  }),
  blank(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    indent:    { firstLine: 0 },
    children:  [new TextRun({ text: 'xxx(name)', size: 28,
                              font: { ascii: 'Cambria Math', eastAsia: 'SimSun', hAnsi: 'Cambria Math' } })],
  }),
  blank(), blank(),

  // ── Abstract ──────────────────────────────────────────────────────────────
  new Paragraph({
    alignment: AlignmentType.CENTER,
    indent:    { firstLine: 0 },
    children:  [new TextRun({ text: '摘要', bold: true })],
  }),
  blank(),
  body('强化学习作为机器学习的重要范式, 通过 Agent 与环境的交互, 依据奖励信号学习最优行为策略, 在机器人控制、自动驾驶、大模型训练等领域取得显著成功。Q-learning 是由 Watkins 于 1989 年提出的无模型强化学习算法, 其核心在于直接估计状态-动作价值函数, 而无需显式建模环境动态。1992 年 Watkins 与 Dayan 共同给出 Q-learning 在离散状态-动作空间下的严格收敛性证明, 本文旨在复现其推导并设计对应实代码实验, 验证其在满足学习率衰减条件与无限探索的条件下, Q 值以概率 1 收敛到最优值。'),
  blank(),
  body('本文以走迷宫问题为应用载体，完整复现 Q-learning 算法，并基于马尔可夫决策过程框架构建实验环境。迷宫采用 6 × 6 网格,包含起点、终点、墙壁及陷阱等元素, Agent 可执行上、下、左、右四个离散动作。环境奖励函数设计为:到达终点+10，撞墙或踩陷阱-10，每步移动-1，以鼓励智能体寻找最短路径。智能体采用 ε-greedy 策略平衡探索与利用, 初始探索率设为 1.0 , 并随回合数以 0.995 的衰减率逐渐降低至 0.01 , 以保证在训练后期充分利用已学知识。Q 值表初始化为零, 更新时采用标准贝尔曼方程, 学习率 α = 0.1 ,折扣因子 γ = 0.9 。训练共进行 500 回合,每回合最大步数限制为 100 步。'),
  blank(),
  body('实验结果表明, 算法在约 80 回合后成功率即达到 100%, 平均每回合奖励从初期的-600提升至接近 0 ，平均步数从 52.52 步下降至 11.33 步，最终学习到一条从起点 (0,0) 到终点 (5,5) 的 10 步最优路径。这一结果直观验证了 Q-learning 在有限 MDP 中的收敛性。此外，本文对关键超参数进行敏感性分析:探索率衰减过快(0.99)可快速收敛，但可能陷入次优；衰减过慢(0.999 )则导致过度探索，收敛缓慢；折扣因子 γ = 0.5 时 Agent 偏向 "短视"，学习效率略低于 γ = 0.9 或 0.99，但三种 γ 值最终均能收敛，显示算法对折扣因子具有一定鲁棒性。这些实验与 Watkins-Dayan 收敛定理的条件(无限探索、学习率条件)相呼应，并凸显了离策略特性行为:策略(ε-greedy)负责探索，目标策略负责学习最优值, 二者解耦使算法兼具稳定性和灵活性。'),
  blank(),
  body('本研究不仅复现了经典 Q-learning 算法, 验证了其理论收敛性, 也为理解强化学习核心机制提供了实验支撑, 对后续研究具有参考价值。'),
  blank(),
  new Paragraph({
    indent: { firstLine: 0 },
    children: [
      new TextRun({ text: '关键词：', bold: true }),
      new TextRun('强化学习；Q-Learning；马尔可夫决策过程；走迷宫；ε-greedy 策略'),
    ],
  }),

  // Page break after abstract
  pageBreak(),

  // ── Table of Contents ────────────────────────────────────────────────────
  new Paragraph({
    alignment: AlignmentType.CENTER,
    indent:    { firstLine: 0 },
    spacing:   { after: 240 },
    children:  [new TextRun({ text: '目录', bold: true, size: 32,
                              font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' } })],
  }),
  new TableOfContents('目录', { hyperlink: true, headingStyleRange: '1-3' }),
  pageBreak(),

  // ═══════════════════════════════════════════════════════════════════════════
  // 一、研究背景与问题定义
  // ═══════════════════════════════════════════════════════════════════════════
  h1Chinese('一、研究背景与问题定义'),
  blank(),

  // 1.1
  h2('问题背景与研究动机'),
  blank(),
  body('强化学习 (Reinforcement Learning) 是机器学习的三大范式之一, 与监督学习和无监督学习并列。其核心思想是让 Agent 通过与 Environment 的持续交互, 基于环境反馈的奖励信号 (Reward) 学习最优的行为策略。这种学习范式模拟了人类和动物通过试错学习的过程, 在游戏博弈、机器人控制、自动驾驶等领域取得了显著成功。'),
  blank(),
  body('Q-Learning 算法由 Christopher Watkins 于 1989 年在其博士论文中首次提出, 并于 1992 年与 Peter Dayan 合作发表了收敛性证明的经典论文 "Q-learning"。该算法属于无模型 (Model-free) 的时序差分 (Temporal Difference) 学习方法, 不需要预知环境的状态转移概率, 仅通过与环境交互获得的经验即可学习最优策略。[1][2]'),
  blank(),
  body('走迷宫问题是强化学习领域的经典基准问题, 它将复杂的决策问题简化为离散的状态空间和动作空间，便于理解和验证算法的有效性。智能体需要在迷宫中从起点导航到终点, 同时避开障碍物, 学习出一条最优路径。这一问题可以自然地建模为马尔可夫决策过程, 是验证 Q-Learning 算法的理想测试平台。'),
  blank(),
  body('本文的研究目标是:(1)基于原始论文完整复现 Q-Learning 算法;(2)构建并训练走迷宫 Agent;(3)通过实验验证算法的收敛性；(4)分析关键超参数对算法性能的影响。'),
  blank(),

  // 1.2
  h2('相关研究调研与文献方法介绍'),
  blank(),
  body('Q-Learning 算法的理论基础源于动态规划中的贝尔曼最方程。Richard Bellman 于 1957 年提出的动态规划方法为序贯决策问题提供了数学框架, 但其要求完全已知环境的状态转移概率, 这在实际应用中往往难以满足。[5]'),
  blank(),
  body('为解决这一问题, 强化学习领域发展出了两类主要方法: 基于模型的方法 (Model-based) 和无模型方法 (Model-free)。Q-Learning 属于后者, 它通过采样的方式直接从与环境的交互经验中学习，无需建立环境的显式模型。[5]'),
  blank(),
  body('Watkins 和 Dayan (1992) 在论文 "Q-learning" 中严格证明了 Q-Learning 的收敛性。其核心定理指出: 在满足以下条件时, Q 值将以概率 1 收敛到最优值 Q*:'),
  blank(),
  body('(1)所有状态-动作对被无限次访问；'),
  body('(2)学习率满足随机逼近条件；'),
  body('(3)奖励有界。'),
  blank(),
  body('该证明采用了 Action-Replay Process 的构造方法,通过建立 Q 值更新与一个人工构造的马尔可夫过程之间的对应关系，利用随机逼近理论完成了收敛性的严格证明。[1]'),
  blank(),
  body('Q-Learning 的一个重要特性是其 Off-policy 性质: 算法更新 Q 值时使用的是贪婪策略 (选择最大 Q 值),而实际执行时可以使用其他探索策略 (如 ε-greedy)。这种解耦使得智能体可以在保证充分探索的同时, 仍能学习到最优策略。'),
  blank(),

  // 1.3 本文符号说明
  h2('本文符号说明'),
  blank(),
  body('为便于后续讨论, 本文采用以下符号体系, 与原论文保持一致:'),
  blank(),
  tableCaption('基本符号说明'),
  threeLineTable(
    ['符号', '说明'],
    [
      ['S', '状态空间，表示智能体所有可能状态的集合。在走迷宫问题中，状态即为智能体在网格中的位置坐标(x, y)'],
      ['A', '动作空间，表示智能体可执行的所有动作的集合。本文设定 A = {上，下，左，右}，共 4 个离散动作。'],
      ['s, x', '当前状态，表示智能体当前所处的位置。'],
      ['a', '当前动作，表示智能体选择执行的动作。'],
      ["s', y", '下一状态，表示执行动作后转移到的新状态。'],
      ['r', '即时奖励，表示执行动作后环境给予的反馈信号。'],
      ['Rₓ(a)', '状态 x 下执行动作 a 的期望即时奖励，Rₓ(a) = E[r | x, a]。'],
      ['R', '奖励上界常数,满足 rₙ ≤ R 对所有 n 成立'],
      ['Q(s, a)', '状态-动作价值函数,表示在状态 s 下采取动作 a 并遵循最优策略所能获得的期望累积回报。'],
      ['V(s)', '状态价值函数,定义为 V(s) = max Q(s, a),表示状态 s 的最优价值。'],
      ['α', '学习率,控制新信息覆盖旧信息的程度,取值范围 (0,1]。'],
      ['γ', '折扣因子,控制对未来奖励的重视程度,取值范围 [0,1)。'],
      ['ε', '探索率,在 ε-greedy 策略中控制随机探索的概率。'],
      ['π', '策略函数,表示从状态到动作的映射, π(s) 表示在状态 s 下选择的动作。'],
      ['π*', '最优策略，使得累积回报期望最大化的策略。'],
      ['Q*', '最优状态-动作价值函数,对应最优策略下的 Q 值。'],
    ],
    [1800, 7270]
  ),
  blank(),

  // ═══════════════════════════════════════════════════════════════════════════
  // 二、Q-Learning 核心原理
  // ═══════════════════════════════════════════════════════════════════════════
  h1Chinese('二、Q-Learning 核心原理'),
  blank(),
  body('Q-Learning 是一种 Model-Free 的强化学习算法, 旨在让 Agent 在马尔可夫决策过程中通过学习找到最优策略。下面基于 Watkins 和 Dayan (1992) 的经典论文, 提取并整理了 Q-Learning 算法收敛性的完整理论推断过程。'),
  blank(),
  body('目标: 证明在满足特定条件下, Q-Learning 算法估计的动作价值函数 Qₙ(x, a) 会以概率 1 收敛到最优动作价值函数 Q*(x, a) 。收敛性证明是强化学习理论的基石。它保证了只要我们给算法足够的时间和探索机会, 它最终一定能学会最好的策略。'),
  blank(),

  // 2.1 问题定义与算法描述
  h2('问题定义与算法描述'),
  blank(),
  body('环境模型:智能体在一个离散、有限的世界中移动，这是一个受控的马尔可夫过程。'),
  blank(),
  body('状态 (State): xₙ ∈ X'),
  body('动作 (Action): aₙ ∈ A'),
  body('奖励 (Reward): rₙ ,其期望值 r̄ₓ(a) 仅取决于状态和动作。'),
  body('状态转移: 状态以概率 P_{xy}[a] 从 x 变为 y 。'),
  body('折扣因子: γ (0 < γ < 1) ,表示未来奖励的现值折扣。'),
  blank(),
  body('价值函数: 对于策略 π ,状态 x 的价值 V^π(x) 定义为总折扣期望奖励。最优价值函数 V*(x) 满足 Bellman 最优方程:'),
  blank(),
  formula('V^{*}(x) = \\max_{a} \\left( R_{x}(a) + \\gamma \\sum_{y} P_{xy}[a] V^{*}(y) \\right)', 1),
  blank(),
  body('Q 值: 定义为在状态 x 执行动作 a ,随后遵循策略 π 的期望折扣奖励:'),
  blank(),
  formula('Q^{\\pi}(x, a) = R_{x}(a) + \\gamma \\sum_{y} P_{xy}[a] V^{\\pi}(y)', 2),
  blank(),
  body('最优 Q 值记为 Q*(x, a) 。显然有 V*(x) = maxₐ Q*(x, a) 。'),
  blank(),
  body('Q-Learning 更新规则: 在第 n 次经历 (episode) 中,智能体观察状态 xₙ ,选择动作 aₙ ,观察新状态 yₙ 并获得奖励 rₙ 。Q 值更新公式如下:'),
  blank(),
  formula('Q_{n}(x, a) = \\begin{cases} (1 - \\alpha_{n}) Q_{n-1}(x, a) + \\alpha_{n} [r_{n} + \\gamma V_{n-1}(y_{n})] & \\text{if } x = x_{n} \\text{ and } a = a_{n} \\\\ Q_{n-1}(x, a) & \\text{otherwise} \\end{cases}', 3),
  blank(),
  body('其中:'),
  body('αₙ 是学习率 (0 < αₙ < 1) 。'),
  body('V_{n-1}(y) = max_b Q_{n-1}(y, b) 是智能体当前认为的状态 y 的最大价值。'),
  blank(),
  body('这个公式的核心思想是 "时序差分"。我们用当前的即时奖励 rₙ (短期收益)加上对未来价值的估计 γ V_{n-1}(yₙ) (长期收益)来更新旧的估计值。'),
  blank(),

  // 2.2 收敛性定理
  h2('收敛性定理'),
  blank(),
  body('定理条件:'),
  blank(),
  body('(1)奖励有界: |rₙ| ≤ R 。'),
  blank(),
  body('(2)学习率条件:'),
  blank(),
  formula('\\sum_{i=1}^{\\infty} \\alpha_{n_{i}(x, a)} = \\infty, \\quad \\sum_{i=1}^{\\infty} \\alpha_{n_{i}(x, a)}^{2} < \\infty, \\quad \\forall x, a', 4),
  blank(),
  body('其中 n_i(x, a) 表示动作 a 在状态 x 第 i 次被尝试的时间步索引。学习率之和为无穷大保证能充分学习，平方和有限保证噪声逐渐减小，最终稳定。'),
  blank(),
  body('(3)充分探索:所有状态-动作对(x，a)都被无限次地采样。'),
  blank(),
  body('定理结论:'),
  blank(),
  body('在上述条件下,当 n → ∞ 时,以概率 1 收敛到最优解:'),
  blank(),
  formula('Q_{n}(x, a) \\xrightarrow{1} Q^{*}(x, a), \\quad \\forall x, a', 5),
  blank(),

  // 2.3 证明核心思想
  h2('证明核心思想: The Action-Replay Process'),
  blank(),
  body('证明的关键在于构造一个人工的控制马尔可夫过程, 称为动作回放过程 (Action-Replay Process, 简称 ARP)。'),
  blank(),
  body('ARP 定义: 基于智能体在真实环境中经历的一系列 Episodes 序列以及对应的学习率序列 αₙ 。原论文使用 "Card Game" 进行类比: 每一段经历 (x_t, a_t, y_t, r_t, α_t) 被写在一张卡片上，所有卡片按时间顺序堆叠成无限高的牌堆。'),
  blank(),
  body('ARP 的状态:由真实过程的状态 x 和阶段 n 组成，记为 (x, n) 。ARP 的动作: 与真实过程相同。'),
  blank(),
  body('状态转移机制:'),
  blank(),
  body('(1)若在 ARP 状态 (x, n) 执行动作 a 。'),
  body('(2)从牌堆顶部向下查找，找到最近的一张记录了对 (x, a) 操作的卡片(假设在第 t 层)。'),
  body('(3)以概率 α_t 回放该卡片。获得奖励 r_t ，转移到状态 (y_t, t - 1) 。以概率 1 - α_t 反面朝上: 丢弃该卡片，继续向下查找更早的卡片。若查到底部卡片(第 0 层)，过程进入吸收态,奖励为 Q₀(x, a) 。'),
  blank(),
  body('构造这个过程是因为 Q-Learning 的更新规则非常像在这个牌堆上进行动态规划。证明的核心逻辑是: Q-Learning 更新的 Qₙ 值恰好等于 ARP 过程在层级 n 的最优价值。当 n 很大时, ARP 过程统计上的转移概率和奖励会收敛到真实环境,因此 Q-Learning 的 Qₙ 会收敛到真实环境的最优 Q* 。'),
  blank(),

  // 2.4 关键引理推导
  h2('关键引理推导'),
  blank(),

  // 2.4.1 引理 A
  h3('引理 A: Qₙ(x, a) 是 ARP 的最优动作值'),
  blank(),
  body('描述: Qₙ(x, a) 恰好是 ARP 状态 (x, n) 下动作 a 的最优动作价值。'),
  blank(),
  formula('Q_{n}(x, a) = Q_{ARP}^{*}(\\langle x, n\\rangle, a)', 6),
  blank(),
  body('证明思路 (归纳法):'),
  blank(),
  body('(1) n = 0 时， Q₀(x, a) 是初始值，也是 ARP 底层的唯一可能价值。'),
  blank(),
  body('(2)归纳假设:假设对于某个 n - 1 (n ≥ 1) ，有: Q_{n-1}(x, a) = Q*_{ARP}(⟨x, n-1⟩, a), ∀x, a.'),
  blank(),
  body('(3)归纳步骤:'),
  blank(),
  body('若 x ≠ xₙ 或 a ≠ aₙ : 此时在 ARP 中执行动作 a 与在状态 ⟨x, n-1⟩ 中执行相同，因为第 n 层没有新的 episode 匹配该状态-动作对。因此:'),
  blank(),
  formula('Q_{ARP}^{*}(\\langle x, n\\rangle, a) = Q_{ARP}^{*}(\\langle x, n-1\\rangle, a) = Q_{n-1}(x, a) = Q_{n}(x, a)', 7),
  blank(),
  body('若 x = xₙ 且 a = aₙ : 在 ARP 中,从状态 ⟨xₙ, n⟩ 执行动作 aₙ 时,根据构造,有两种可能: 以概率 1 - αₙ 等同于从 ⟨xₙ, n-1⟩ 执行动作 aₙ ; 或以概率 αₙ 重放 episode n ,得到即时奖励 rₙ 并转移到状态 ⟨yₙ, n-1⟩。因此,该动作的最优值满足 Bellman 方程:'),
  blank(),
  formula('Q_{ARP}^{*}(\\langle x_{n}, n\\rangle, a_{n}) = (1 - \\alpha_{n}) Q_{ARP}^{*}(\\langle x_{n}, n-1\\rangle, a_{n}) + \\alpha_{n} [r_{n} + \\gamma V_{ARP}^{*}(\\langle y_{n}, n-1\\rangle)]', 8),
  blank(),
  body('利用归纳假设, Q*_{ARP}(⟨xₙ, n-1⟩, aₙ) = Q_{n-1}(xₙ, aₙ) ,且 V*_{ARP}(⟨yₙ, n-1⟩) = V_{n-1}(yₙ) 。代入得:'),
  blank(),
  formula('Q_{ARP}^{*}(\\langle x_{n}, n\\rangle, a_{n}) = (1 - \\alpha_{n}) Q_{n-1}(x_{n}, a_{n}) + \\alpha_{n} [r_{n} + \\gamma V_{n-1}(y_{n})]', 9),
  blank(),
  body('这正是 Q-learning 更新公式所定义的 Qₙ(xₙ, aₙ) 。因此归纳完成。'),
  blank(),

  // 2.4.2 引理 B
  h3('引理 B: ARP 收敛到真实过程'),
  blank(),
  body('这一组引理证明当层级 n → ∞ 时, ARP 的统计特性趋近于真实环境。也就是说采样次数足够多, 对期望的估计会越来越准。'),
  blank(),

  body('引理 B.1: 折扣序列的有限性'),
  blank(),
  body('描述: 一个折扣、有界奖励、有限马尔可夫过程。从任意起始状态 x 出发,执行任意 s 个动作所获得的值,与执行这 s 个动作后再执行任意后续策略所获得的值之差,当 s → ∞ 时趋于 0 。'),
  blank(),
  body('证明: 设执行 s 个动作 a₁, …, a_s 后,后续策略为 π 。则完整无限序列的值为:'),
  blank(),
  formula('Q(x, a_{1}, \\ldots, a_{s}, \\pi) = Q(x, a_{1}, \\ldots, a_{s}) + \\gamma^{s} \\sum_{y_{s+1}} P_{y_{s} y_{s+1}}[a_{s}] V^{\\pi}(y_{s+1})', 10),
  blank(),
  body('其中 y_s 是第 s 步后的状态。忽略后续部分带来的误差为:'),
  blank(),
  formula('\\delta = \\gamma^{s} \\sum_{y_{s+1}} P_{y_{s} y_{s+1}}[a_{s}] V^{\\pi}(y_{s+1})', 11),
  blank(),
  body('由于所有奖励有界 |rₙ| ≤ R ,状态值也有界:'),
  blank(),
  formula('|V^{\\pi}(y)| \\leq \\frac{R}{1 - \\gamma}', 12),
  blank(),
  body('因此,'),
  blank(),
  formula('|\\delta| \\leq \\gamma^{s} \\frac{R}{1 - \\gamma} \\to 0 \\quad \\text{当 } s \\to \\infty', 13),
  blank(),

  body('引理 B.2: 执行 s 步后跌至低层的概率可任意小'),
  blank(),
  body('描述: 对于任意层级 l ,存在一个更高的层级 h ,使得从 h 开始执行 s 步动作后, 层级掉到 l 以下的概率可以任意小。'),
  blank(),
  body('证明: 固定状态-动作对 (x, a) 。定义 i_h 为满足 n^i(x, a) ≤ n 的最大 i, i_l 为满足 n^i(x, a) ≥ l 的最小 i 。从层 n > l 执行一次动作 a 后,跌至低于 l 的概率等于在搜索匹配卡片过程中最终重放的卡片索引小于 i_l 的概率。根据 ARP 的构造,该概率为:'),
  blank(),
  formula('[\\prod_{i=i_{l}}^{i_{h}} (1 - \\alpha_{n^{i}})] \\sum_{j=0}^{i_{l}-1} \\{ \\alpha_{n^{j}} \\prod_{k=j+1}^{i_{l}-1} (1 - \\alpha_{n^{k}}) \\}', 14),
  blank(),
  body('上式左边因子表示所有从 i_l 到 i_h 的卡片都被跳过 (即未重放),右边求和表示最终重放的卡片来自低于 i_l 的某张。该概率不大于:'),
  blank(),
  formula('\\prod_{i=i_{l}}^{i_{h}} (1 - \\alpha_{n^{i}}) \\leq \\exp(-\\sum_{i=i_{l}}^{i_{h}} \\alpha_{n^{i}})', 15),
  blank(),
  body('由于条件 (3) 保证 Σ_{i=i_l}^∞ α_n^i = ∞ ,当 n 趋于无穷时,该指数趋于 0 。因此,对任意 η > 0 ,存在足够大的 n 使得上述概率小于 η。因为状态和动作空间有限,我们可以统一选择 n 使得对所有 (x, a) 都满足。对于 s 步的情况,我们需确保每一步后都不跌至低于 l 。通过递推,可以选取足够高的起始层,使得每一步跌落的概率都足够小,从而总概率可控。'),
  blank(),

  body('引理 B.3: 奖励和转移概率以概率 1 收敛'),
  blank(),
  body('描述: ARP 中的期望即时奖励 Rₓ^(n)(a) 和转移概率 P_{xy}^(n)[a] 当层数 n → ∞ 时收敛以概率 1 到真实过程的 Rₓ(a) 和 P_{xy}[a]。'),
  blank(),
  body('证明: 采用随机逼近理论中的经典结论 (Kushner & Clark, 1978, 定理 2.3.1): 若随机变量序列 Xₙ 满足更新规则:'),
  blank(),
  formula('X_{n+1} = X_{n} + \\beta_{n} (\\xi_{n} - X_{n})', 16),
  blank(),
  body('其中 0 ≤ βₙ < 1, Σ_{i=1}^∞ βₙ = ∞, Σ_{i=1}^∞ βₙ² < ∞ ,且 ξₙ 是有界随机变量,期望为 Ξ , 则 Xₙ →¹ Ξ 。'),
  blank(),
  body('奖励的收敛: 对于固定的 (x, a) ,令 n^i 表示第 i 次在状态 x 执行动作 a 的时刻。在 ARP 中,期望奖励 R_{⟨x, n^{i+1}⟩}(a) 的更新为:'),
  blank(),
  formula('R_{\\langle x, n^{i+1}\\rangle}(a) = R_{\\langle x, n^{i}\\rangle}(a) + \\alpha_{n^{i+1}} (r_{n^{i+1}} - R_{\\langle x, n^{i}\\rangle}(a))', 17),
  blank(),
  body('因此,由上述定理 R_{⟨x, n^i⟩}(a) →¹ Rₓ(a) 。由于状态和动作有限,收敛是一致的。'),
  blank(),
  body('转移概率的收敛: 先定义指示函数:'),
  blank(),
  formula('\\chi_{n}(y) = \\begin{cases} 1 & \\text{if } y_{n} = y \\\\ 0 & \\text{otherwise} \\end{cases}', 18),
  blank(),
  body('其期望为 P_{xy}[a] 。ARP 中转移概率 P_{xy}^{(n^{i+1})}[a] 的更新为:'),
  blank(),
  formula('P_{xy}^{(n^{i+1})}[a] = P_{xy}^{(n^{i})}[a] + \\alpha_{n^{i+1}} (\\chi_{n^{i+1}} - P_{xy}^{(n^{i})}[a])', 19),
  blank(),
  body('同样满足定理条件,故 P_{xy}^{(n)}[a] →¹ P_{xy}[a] 。此外，结合 B.2，若我们条件于执行 s 步后仍高于某个固定层 k ，则上述收敛性仍然成立,因为条件事件概率趋于 1,不影响极限。'),
  blank(),

  body('引理 B.4: 接近的奖励和转移概率导致接近的动作值'),
  blank(),
  body('描述: 设有 s 个马尔可夫链,其转移矩阵分别为 P^i_{xy}[a] ,奖励函数分别为 R^i_x(a) (i = 1, …, s)。考虑由这些链依次串联而成的 s 步链: 从状态 x₁ 出发,按 P¹[a₁] 转移到 x₂ ,获得奖励 R^i_x(a₁) ,然后按 P²[a₂] 转移,以此类推。给定 η > 0 ,若对任意 a, x, y 有:'),
  blank(),
  body('|P^i_{xy}[a] - P_{xy}[a]| < η/R, |R^i_x(a) - R_x(a)| < η'),
  blank(),
  body('则串联链中执行动作 a₁, …, a_s 所获得的值与真实过程中执行相同动作的值之差不超过 s(s+1)η/2 。'),
  blank(),
  body('证明: 首先考虑 s = 2 的情况。定义真实过程的两步值:'),
  blank(),
  formula('Q(x, a_{1}, a_{2}) = R_{x}(a_{1}) + \\gamma \\sum_{y} P_{xy}[a_{1}] R_{y}(a_{2})', 21),
  blank(),
  body('串联链的两步值:'),
  blank(),
  formula("Q'(x, a_{1}, a_{2}) = R_{x}^{1}(a_{1}) + \\gamma \\sum_{y} P_{xy}^{1}[a_{1}] R_{y}^{2}(a_{2})", 22),
  blank(),
  body('可以通过逐层递推证明总误差不超过 Σ_{k=1}^s kη = s(s+1)η/2 。'),
  blank(),

  // 2.5 收敛定理的证明
  h2('收敛定理的证明'),
  blank(),
  body('定理描述:'),
  blank(),
  body('在上述给定的条件下, 有'),
  blank(),
  formula('Q_{n}(x, a) \\xrightarrow{1} Q^{*}(x, a) \\quad n \\to \\infty', 26),
  blank(),
  body('证明:'),
  blank(),
  body('采用引理 A 和 B 的结果。固定状态 x 和动作 a ,给定任意 ε > 0 。'),
  blank(),
  body('选择截断步数 s : 由 B.1,存在 s 使得对于任何策略,忽略 s 步后的奖励所带来的误差小于 ε/6 。具体地,取 s 满足:'),
  blank(),
  formula('\\gamma^{s} \\frac{R}{1 - \\gamma} < \\frac{\\epsilon}{6}', 27),
  blank(),
  body('选择足够高的层数 l : 由 B.3,以概率 1,存在 l 使得对 ∀n > l 和 ∀x, a, y ,有:'),
  blank(),
  body('|P_{xy}^{(n)}[a] - P_{xy}[a]| < ε / (3s(s+1)R), |R_x^{(n)}(a) - R_x(a)| < ε / (3s(s+1))'),
  blank(),
  body('选择更高的起始层 h : 由 B.2, ∃h > l 使得从任何高于 h 的层开始,在 ARP 中执行 s 步后跌至低于 l 的概率小于 δ := min{ε(1-γ)/(6sR), ε/(3s(s+1)R)} 。'),
  blank(),
  body('这意味着,对于 n > h ,考虑条件概率 P_{xy}^{(n)}[a] 和 R_x^{(n)}(a) ,条件于执行 s 步后仍高于 l ,它们与真实值的偏差可进一步控制。'),
  blank(),
  body('估计 ARP 中 s 步动作值与真实值的差: 考虑在 ARP 中从状态 (⟨x, n⟩) 执行动作 a₁, …, a_s 的 s 步值 Q̄_ARP(⟨x, n⟩, a₁, …, a_s) ,与真实过程中相同动作序列的值 Q̄(x, a₁, …, a_s) 之差。将其分解为两部分: 一部分是由于可能跌至低于 l 造成的风险; 另一部分是在高于 l 的条件下,转移概率和奖励的偏差。'),
  blank(),
  body('首先,跌至低于 l 的概率不超过 δ ,而一旦跌至低于 l ,可能导致的奖励偏差最大为 2sR/(1-γ) 。因此,风险项的贡献不超过 δ · 2sR/(1-γ) ≤ ε/3 。'),
  blank(),
  body('其次,在高于 l 的条件下,由步骤 2 和 B.3 的收敛性,条件部分的误差不超过 s(s+1)/2 · 2ε/(3s(s+1)) = ε/3 。'),
  blank(),
  body('因此,总误差小于 2ε/3 。'),
  blank(),
  body('考虑无限步与有限步的差异: 由 B.1,真实过程中,用 s 步近似代替无限步的误差小于 ε/6 ; 同样,在 ARP 中,由于折扣因子相同,截断误差也小于 ε/6 。因此,对于任意动作序列, 特别是对于最优动作, 有:'),
  blank(),
  body('|Q*_ARP(⟨x, n⟩, a) - Q*(x, a)| < 2ε/3 + ε/6 + ε/6 = ε'),
  blank(),
  body('结合引理 A: 我们有 Qₙ(x, a) = Q*_ARP(⟨x, n⟩, a) 。因此, |Qₙ(x, a) - Q*(x, a)| < ε, ∀n > h . 由于 ε 任意,且上述论证对几乎所有样本路径成立 (以概率 1 ),故:'),
  blank(),
  body('Qₙ(x, a) →¹ Q*(x, a) 当 n → ∞'),
  blank(),

  // ═══════════════════════════════════════════════════════════════════════════
  // 三、基于 Q-Learning 的走迷宫智能体构建
  // ═══════════════════════════════════════════════════════════════════════════
  h1Chinese('三、基于 Q-Learning 的走迷宫智能体构建'),
  blank(),
  body('Q-Learning 算法的核心是 Q 值更新公式, 其数学表达如下:'),
  blank(),
  formula('Q_{n}(x, a) = (1 - \\alpha_{n}) Q_{n-1}(x, a) + \\alpha_{n} [r_{n} + \\gamma \\max_{b} Q_{n-1}(y_{n}, b)]', 36),
  blank(),
  body('原始论文的收敛性证明依赖于 "所有状态-动作对被无限次访问" 的存在性条件，并未指定具体探索方式。在实际实现中,我们采用 ε-greedy 策略来保证这一条件:'),
  blank(),
  formula('\\pi(s) = \\begin{cases} \\operatorname{argmax}_{a} Q(s, a), & \\text{以概率 } 1 - \\varepsilon \\text{ 利用} \\\\ \\text{均匀概率随机选择动作}, & \\text{以概率 } \\varepsilon \\text{ 探索} \\end{cases}', 37),
  blank(),
  body('该策略通过 ε 参数平衡探索与利用: 大部分时间选择当前认为最优的动作 (利用), 小部分时间随机尝试其他动作 (探索), 以确保充分探索状态空间。[3][4]'),
  blank(),

  // 3.1 实验环境设计
  h2('实验环境设计'),
  blank(),
  body('本文构建了一个 6×6 的网格迷宫环境，共包含 36 个状态。迷宫中的状态编码如下: 0 表示空地 (可通行的白色格子), 1 表示墙壁 (不可通行的黑色区域, 共 10 个), 2 表示起点 (位于坐标 (0,0) 的绿色格子),3 表示终点 (位于坐标 (5,5) 的红色格子),4 表示陷阱 (位于坐标 (4,5) 的橙色格子)。动作空间包含 4 个离散动作 (上、下、左、右)。'),
  blank(),
  body('奖励函数设计遵循以下原则:到达终点给予 +10 的正奖励以激励智能体完成任务； 撞墙和踩陷阱给予-10的负奖励以惩罚错误行为；每走一步给予-1的小额负奖励以鼓励 agent 寻找最短路径。'),
  blank(),
  image('1.jpg', 400, 400),
  figCaption('迷宫环境示意图'),
  blank(),

  // 3.2 代码设计
  h2('代码设计'),
  blank(),
  body('本文的代码实现采用模块化设计，主要包含以下几个核心模块:'),
  blank(),
  body('(1)MazeEnv 类:环境模拟模块。实现了迷宫环境的核心功能，包括 reset (重置环境到初始状态)、step(执行动作并返回下一状态、奖励和终止标志)、is_valid_position (判断位置是否可通行)等方法。'),
  blank(),
  body('(2)QLearningAgent 类:Q 表管理模块。负责维护 Q 表 (大小为[36, 4]的二维矩阵)、基于 ε-greedy 策略选择动作(以概率 ε 随机探索，否则选择 Q 值最大的动作)、以及使用贝尔曼更新公式进行 Q 值更新。'),
  blank(),
  body('(3)train 函数:训练循环模块。实现了完整的训练流程，包括回合迭代、状态转移、Q 值更新和探索率衰减,同时记录每回合的统计数据 (总奖励、步数、是否成功)。'),
  blank(),
  body('(4)可视化模块:多个可视化函数，用于生成训练曲线(奖励、步数、成功率、探索率衰减)、Q表热力图(状态价值、最优策略、Q表)、最优路径展示以及参数敏感性分析图表。'),
  blank(),

  // 3.3 Q-Learning 算法伪代码
  h2('Q-Learning 算法伪代码'),
  blank(),
  body('下面给出本文实现的 Q-Learning 迷宫探索算法的伪代码描述:'),
  blank(),

  // Algorithm pseudocode as formatted paragraphs
  new Paragraph({
    indent: { firstLine: 0 },
    spacing: { line: 240, lineRule: LineRuleType.AUTO },
    children: [new TextRun({ text: 'Algorithm: Q-Learning for Maze Navigation', bold: true, font: { ascii: 'Cambria Math', hAnsi: 'Cambria Math' } })],
  }),
  blank(),
  new Paragraph({ indent: { firstLine: 0 }, children: [new TextRun('Input: 学习率 α = 0.1, 折扣因子 γ = 0.9, 初始探索率 ε = 1.0')] }),
  blank(),
  new Paragraph({ indent: { firstLine: 0 }, children: [new TextRun('Initialize: Q(s, a) = 0 for all s ∈ S, a ∈ A')] }),
  blank(),
  new Paragraph({ indent: { firstLine: 0 }, children: [new TextRun('For episode = 1 To 500:')] }),
  new Paragraph({ indent: { left: 480 }, children: [new TextRun('s ← reset environment (start position)')] }),
  new Paragraph({ indent: { left: 480 }, children: [new TextRun('While s is not terminal and steps < 100:')] }),
  new Paragraph({ indent: { left: 960 }, children: [new TextRun('With probability ε: a ← random action')] }),
  new Paragraph({ indent: { left: 960 }, children: [new TextRun('Otherwise: a ← argmax_a Q(s, a)')] }),
  new Paragraph({ indent: { left: 960 }, children: [new TextRun('Execute a, observe reward r and next state s\'')] }),
  new Paragraph({ indent: { left: 960 }, children: [new TextRun('Q(s, a) ← Q(s, a) + α[r + γ · max_{a\'} Q(s\', a\') - Q(s, a)]')] }),
  new Paragraph({ indent: { left: 960 }, children: [new TextRun('s ← s\'')] }),
  new Paragraph({ indent: { left: 480 }, children: [new TextRun('End While')] }),
  new Paragraph({ indent: { left: 480 }, children: [new TextRun('ε ← max(0.01, ε × 0.995)')] }),
  new Paragraph({ indent: { firstLine: 0 }, children: [new TextRun('End For')] }),
  blank(),
  body('该算法的核心是贝尔曼更新公式, 其中学习率 α 控制新信息的接受程度, 折扣因子 γ 决定未来奖励的重要性。ε-greedy 策略在探索与利用之间取得平衡,随着训练进行逐渐从探索转向利用。每回合最多 100 步的限制防止智能体在早期无限循环。'),
  blank(),

  // 3.4 实验结果与分析
  h2('实验结果与分析'),
  blank(),
  body('经过 500 回合训练, Q-Learning 算法在本文构建的 6 × 6 迷宫环境中表现出良好的收敛性和有效性。下面从迷宫环境结构、训练过程、Q 表学习结果、最优路径等方面进行详细分析。'),
  blank(),

  image('2.jpg', 500, 400),
  figCaption('训练曲线'),
  blank(),
  body('图 3-2 展示了训练过程中各项指标的变化趋势。左上子图为每回合总奖励曲线, 可见奖励从初期的约 -600 逐步提升，在约 300 回合后稳定在接近 0 附近；右上子图为每回合步数曲线，步数从初期的 100 步(达到上限)逐渐下降，在约 200 回合后稳定在约 11-13 步; 左下子图为成功率曲线, 显示智能体在约 80 回合后即达到 100% 成功率并保持稳定; 右下子图为探索率 ε 的衰减曲线,从 1.0 平滑衰减至约 0.08 。整体来看,算法在约 200-300 回合实现全面收敛。'),
  blank(),

  image('3.jpg', 500, 400),
  figCaption('Q 表可视化'),
  blank(),
  body('图 3-3 包含三个子图, 从不同角度展示了 Q 表的学习结果。左图为状态价值函数 V(s)热力图,显示了每个状态的最大 Q 值,值域从 -2.5(深红色，远离终点的死胡同区域)到10.0(深绿色，终点位置)，形成了清晰的价值梯度，引导智能体向终点移动。中图为最优策略 π*(s) 可视化，用蓝色箭头表示每个状态的最优动作方向，灰色方块表示障碍物,可见箭头均指向终点方向并巧妙绕过障碍。右图为完整 Q 表热力图,横轴为 36 个状态，纵轴为 4 个动作，明亮的黄色区域集中在终点附近状态，表明这些状态-动作对具有较高的期望回报。'),
  blank(),

  image('4.jpg', 500, 400),
  figCaption('最优路径'),
  blank(),
  body('图 3-4 展示了训练完成后智能体学会的最优路径。该路径共计 10 步, 具体走法为: 从起点 (0,0) 出发，经过 (1,0) → (1,1) → (2,1) → (2,2) → (2,3) → (3,3) → (4,3) → (5,3) → (5,4) 最终到达终点 (5,5) 。蓝色连线和圆形标记清晰地显示了路径轨迹,智能体成功绕过了所有黑色墙壁障碍和橙色陷阱，找到了一条安全且高效的路径。该结果证实了 Q-Learning 算法在离散状态空间中学习最优策略的能力。'),
  blank(),

  // 3.5 参数敏感性分析
  h2('参数敏感性分析'),
  blank(),

  image('5.jpg', 500, 400),
  figCaption('探索率 ε 敏感性分析'),
  blank(),
  body('图 3-5 对比了三种不同 ε 衰减率(0.99、0.995、0.999)对算法性能的影响。左图显示奖励曲线: ε 衰减率为 0.99 时收敛最快，奖励快速接近最优值；0.995 时收敛稍慢但最终也能达到相似水平; 0.999 时因探索过多导致学习明显较慢, 500 回合内未能充分收敛。中图显示步数曲线: 0.99 和 0.995 均能快速降至最优步数 (约 10 步), 而 0.999 保持在约 30 步且波动较大。右图显示 ε 衰减过程: 0.99 在 400 回合前已接近 0,0.995 在 500 回合时降至约 0.08，0.999 在 500 回合时仍约为 0.6 。实验表明，适当的探索率衰减速度对算法性能至关重要, 衰减过慢会导致过度探索, 阻碍策略收敛。'),
  blank(),

  image('6.jpg', 500, 400),
  figCaption('折扣因子 γ 敏感性分析'),
  blank(),
  body('图 3-6 对比了三种不同折扣因子 γ (0.5、0.9、0.99)对算法性能的影响。左图为平均奖励曲线，三种配置均从约 -450 逐步提升并收敛，其中 γ = 0.5 (蓝色)波动稍大且收敛稍慢， γ = 0.9 (橙色)和 γ = 0.99 (绿色)表现相近且更为平滑。右图为平均步数曲线，均从约 90 步降至约 10-15 步, γ = 0.9 和 0.99 收敛更快。实验表明,算法对折扣因子的选择具有一定的鲁棒性,三种 γ 值最终均能收敛。但较高的 γ 值(0.9 和 0.99)有利于智能体考虑长期回报，学习速度更快且收敛更平滑，而 γ = 0.5 时智能体偏向"短视"，只关注即时奖励，导致学习效率略低。'),
  blank(),

  body('综合以上实验结果, 本文验证了 Q-Learning 算法的有效性和收敛性。算法成功学会了从起点到终点的 10 步最优路径, 最终成功率达到 100%, 平均步数从初期的 52.52 步降至最终的 11.33 步。与原论文的理论分析一致, 在满足所有状态-动作对被充分访问的条件下, Q 值收敛到最优值。探索率 ε 的衰减策略和折扣因子 γ 的选择对算法性能有显著影响, 合理的参数设置是算法成功的关键。实验同时证实了 Q-Learning 的离策略特性 ——智能体使用 ε-greedy 策略进行探索，但 Q 值更新使用贪婪策略，两者解耦使得算法能够在探索的同时学习最优策略。'),
  blank(),

  // ═══════════════════════════════════════════════════════════════════════════
  // 参考文献
  // ═══════════════════════════════════════════════════════════════════════════
  pageBreak(),

  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    indent:  { firstLine: 0 },
    children: [new TextRun('参考文献')],
  }),
  blank(),
  ref('Watkins C J C H, Dayan P. Q-learning[J]. Machine Learning, 1992, 8(3-4): 279-292.'),
  ref('Watkins C J C H. Learning from delayed rewards[D]. Cambridge: King\'s College, University of Cambridge, 1989.'),
  ref('周志华. 机器学习[M]. 北京: 清华大学出版社, 2016.'),
  ref('Mohri, M., Rostamizadeh, A., & Talwalkar, A.. Foundations of machine learning (2nd ed.)[M]. The MIT Press, 2018.'),
  ref('Sutton R S, Barto A G. Reinforcement Learning: An Introduction[M]. 2nd ed. Cambridge: MIT Press, 2018.'),
];

// ─────────────────────────────────────────────────────────────────────────────
// Build & write document
// ─────────────────────────────────────────────────────────────────────────────

const doc = new Document({
  styles:    STYLES,
  numbering: NUMBERING,
  sections: [{
    properties: {
      page: {
        size:   { width: PAGE_W, height: PAGE_H },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
      },
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            indent:    { firstLine: 0 },
            children:  [new TextRun({ children: [PageNumber.CURRENT] })],
          }),
        ],
      }),
    },
    children: CONTENT,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUTPUT_PATH, buf);
  console.log(`✓  Written: ${OUTPUT_PATH}`);
}).catch(err => {
  console.error('Error building document:', err.message);
  process.exit(1);
});
