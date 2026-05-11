/**
 * new_doc.js — Chinese Academic Paper Template (课程论文模板)
 *
 * All formatting is pre-configured to GB/T 7714 + Chinese university standards:
 *   - A4 page, 2.5 cm margins on all sides
 *   - SimSun 12pt body, SimHei headings (Cambria Math for English/numbers)
 *   - Word auto-numbering for H2/H3 headings (per-chapter reset via sections_cN references)
 *   - Word auto-numbering for H2/H3 headings (per-chapter reset via sections_cN references)
 *   - Figure/Table captions via Word SEQ fields (per-chapter: 图 1-1, 表 2-3)
 *   - Three-line table helper with proper border handling
 *   - LaTeX formula support (block and inline) via temml + Word native math
 *   - Citation superscript handling [n] format
 *   - Reference list [1][2][3] numbering (Word auto-numbered)
 *   - Footer: centered page number
 *
 * Usage:
 *   1. Edit the CONTENT SECTION below
 *   2. node scripts/new_doc.js
 *   3. Outputs output.docx (or set OUTPUT_PATH)
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

// MathML to docx Math converter
const { mathmlToDocxChildren } = require('./mathml-to-docx');
const temml = require('temml');

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

const OUTPUT_PATH   = 'output.docx';

// Page / margin (DXA: 1440 = 1 inch, 567 ≈ 1 cm)
const PAGE_W        = 11906;   // A4
const PAGE_H        = 16838;
const MARGIN        = 1418;    // 2.5 cm
const CONTENT_W     = PAGE_W - 2 * MARGIN;  // 9070 DXA

// Three-line table border presets
// ISSUE 1 FIX: Use NONE with proper color for invisible borders
const THICK = { style: BorderStyle.SINGLE, size: 12, color: '000000' }; // 1.5 pt
const THIN  = { style: BorderStyle.SINGLE, size: 6,  color: '000000' }; // 0.75 pt
const NONE  = { style: BorderStyle.NONE,   size: 0,  color: 'FFFFFF' };

// Chapter counter — tracks current chapter for H2/H3/caption auto-numbering
// Incremented by h1Chinese() at each chapter boundary
let _chapter = 0;

// ─────────────────────────────────────────────────────────────────────────────
// ISSUE 5/7 FIX: Inline Math Detection and Conversion
// ─────────────────────────────────────────────────────────────────────────────

// Unicode math symbols to LaTeX mapping
const UNICODE_TO_LATEX = {
  // Greek letters
  'α': '\\alpha', 'β': '\\beta', 'γ': '\\gamma', 'δ': '\\delta', 'ε': '\\varepsilon',
  'ζ': '\\zeta', 'η': '\\eta', 'θ': '\\theta', 'ι': '\\iota', 'κ': '\\kappa',
  'λ': '\\lambda', 'μ': '\\mu', 'ν': '\\nu', 'ξ': '\\xi', 'π': '\\pi',
  'ρ': '\\rho', 'σ': '\\sigma', 'τ': '\\tau', 'υ': '\\upsilon', 'φ': '\\phi',
  'χ': '\\chi', 'ψ': '\\psi', 'ω': '\\omega',
  'Γ': '\\Gamma', 'Δ': '\\Delta', 'Θ': '\\Theta', 'Λ': '\\Lambda', 'Ξ': '\\Xi',
  'Π': '\\Pi', 'Σ': '\\Sigma', 'Φ': '\\Phi', 'Ψ': '\\Psi', 'Ω': '\\Omega',
  // Subscript digits
  '₀': '_0', '₁': '_1', '₂': '_2', '₃': '_3', '₄': '_4',
  '₅': '_5', '₆': '_6', '₇': '_7', '₈': '_8', '₉': '_9',
  // Subscript Latin letters (full set)
  'ₐ': '_a', 'ₑ': '_e', 'ₕ': '_h', 'ᵢ': '_i', 'ⱼ': '_j', 'ₖ': '_k',
  'ₗ': '_l', 'ₘ': '_m', 'ₙ': '_n', 'ₒ': '_o', 'ₚ': '_p',
  'ᵣ': '_r', 'ₛ': '_s', 'ₜ': '_t', 'ᵤ': '_u', 'ᵥ': '_v',
  'ₓ': '_x', 'ᵧ': '_y', 'ᵦ': '\\beta', 'ᵧ': '\\gamma',
  // Superscript
  '⁰': '^0', '¹': '^1', '²': '^2', '³': '^3', '⁴': '^4',
  '⁵': '^5', '⁶': '^6', '⁷': '^7', '⁸': '^8', '⁹': '^9',
  'ⁿ': '^n', 'ⁱ': '^i',
  // Special symbols
  '∞': '\\infty', '∑': '\\sum', '∏': '\\prod', '∫': '\\int',
  '≤': '\\leq', '≥': '\\geq', '≠': '\\neq', '≈': '\\approx',
  '→': '\\to', '←': '\\leftarrow', '↔': '\\leftrightarrow',
  '∈': '\\in', '∉': '\\notin', '⊂': '\\subset', '⊃': '\\supset',
  '∀': '\\forall', '∃': '\\exists', '∧': '\\land', '∨': '\\lor',
  '×': '\\times', '÷': '\\div', '±': '\\pm', '∓': '\\mp',
  '·': '\\cdot', '…': '\\ldots', '⋯': '\\cdots',
  '′': "'", '″': "''",
  '⟨': '\\langle ', '⟩': ' \\rangle',
  // Superscript letter (for π*, Q*, etc.)
  '*': '^*',
};

// Full regex character classes for detection — must be kept in sync with UNICODE_TO_LATEX keys
const GREEK_CHARS   = 'αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ';
const SUB_SUP_CHARS = '₀₁₂₃₄₅₆₇₈₉ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓᵧᵦ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ';
const MATH_OPS      = '∞∑∏∫≤≥≠≈→←↔∈∉⊂⊃∀∃∧∨×÷±∓·…⋯′″⟨⟩';
const GREEK_CLASS   = new RegExp(`[${GREEK_CHARS}]`);
const SUB_SUP_CLASS = new RegExp(`[${SUB_SUP_CHARS}]`);
const MATH_OPS_CLASS = new RegExp(`[${MATH_OPS}]`);

/**
 * Convert text with Unicode math symbols to LaTeX format.
 * Also handles combining macron ̄ (e.g. r̄ → \bar{r}).
 * @param {string} text - Text with Unicode math symbols
 * @returns {string} LaTeX formatted text
 */
function unicodeToLatex(text) {
  // Normalize combining marks: (letter)(̄) → \bar{letter}
  let result = text.replace(/([A-Za-z])̄/g, '\\bar{$1}');
  // Normalize bare ^ and _ to LaTeX braced form: n^i → n^{i}, x_n → x_{n}
  // Use negative lookahead to skip already-braced forms like ^{xy}
  result = result.replace(/(\^|_)(?!\{)([A-Za-z0-9]+)/g, '$1{$2}');
  for (const [unicode, latex] of Object.entries(UNICODE_TO_LATEX)) {
    result = result.split(unicode).join(latex);
  }
  return result;
}

/**
 * Find full extent of a [_^]{...} expression with nested brace counting.
 * @param {string} text - Full text
 * @param {number} pos - Position of the ^ or _ character
 * @returns {{start: number, end: number}}|null
 */
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

/**
 * Detect if text contains math content (needs formula editor rendering)
 * ISSUE 7 FIX: Strict detection - don't match plain numbers or English words
 * @param {string} text - Text to detect
 * @returns {boolean}
 */
function containsMath(text) {
  if (GREEK_CLASS.test(text)) return true;
  if (SUB_SUP_CLASS.test(text)) return true;
  if (MATH_OPS_CLASS.test(text)) return true;
  // Starred symbols like π*, Q*
  if (/[A-Z]\*/.test(text)) return true;
  // $...$ LaTeX delimiters
  if (/\$[^$]+\$/.test(text)) return true;
  // LaTeX-style subscripts/superscripts: _{xy}, ^{2}, _{n+1}
  if (/[_^]\{[^}]+\}/.test(text)) return true;
  // Bare _ and ^ subscripts/superscripts: n^i, x_n, i_h, P_xy
  if (/[A-Za-z]\d*[_^][a-zA-Z0-9]/.test(text)) return true;
  // Backslash LaTeX commands: \leq, \alpha, \bar{x}, \frac{a}{b} etc.
  if (/\\[a-zA-Z]/.test(text)) return true;
  // Combining macron above letter
  if (/[A-Za-z]̄/.test(text)) return true;
  return false;
}

/**
 * ISSUE 8 FIX: Detect if text contains citation [n] format
 */
function containsCitation(text) {
  return /\[\d+\]/.test(text);
}

/**
 * ISSUE 7 FIX: Parse text into TextRun and Math mixed array (for inline formulas)
 * Strict regex - only matches actual math content, not plain numbers or words
 * @param {string} text - Input text
 * @returns {Array} Array of TextRun and Math objects
 */
function parseInlineContent(text) {
  // Regex for non-braced patterns (_{...} with simple content handled by regex, nested handled below)
  const mathPattern = new RegExp(
    `\\$([^$]+)\\$` +
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

  // Find nested-brace expressions first (_{...{...}...} and ^{...{...}...})
  const nested = [];
  const trigger = /[_^]\{/g;
  let t;
  while ((t = trigger.exec(text)) !== null) {
    const ext = findBracedExtent(text, t.index);
    if (ext && ext.start < t.index && ext.end - ext.start > t[0].length + 1) {
      // Is actually nested (extent longer than simple _{x})
      nested.push({ start: ext.start, end: ext.end, content: text.slice(ext.start, ext.end) });
    }
  }

  const children = [];
  let lastIndex = 0;
  let m;
  while ((m = mathPattern.exec(text)) !== null) {
    // Skip matches inside nested-brace regions
    if (nested.some(n => m.index + m[0].length > n.start && m.index < n.end)) continue;
    if (m.index > lastIndex) {
      // Insert any nested matches that fall in the gap
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
    const mc = m[1] || m[2] || m[3] || m[4] || m[5] || m[6] || m[7] || m[8] || m[9] || m[10] || m[11] || m[12] || m[13] || m[14];
    if (mc) {
      const latex = unicodeToLatex(mc);
      try {
        const ml = temml.renderToString(latex, { displayMode: false, throwOnError: false });
        const kids = mathmlToDocxChildren(ml);
        children.push(new Math({ children: kids && kids.length ? kids : [new MathRun(mc)] }));
      } catch (e) { children.push(new Math({ children: [new MathRun(mc)] })); }
    }
    lastIndex = m.index + m[0].length;
  }
  // Handle remaining text and nested matches after the last regex match
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

/**
 * ISSUE 8 FIX: Parse text with math content and citations
 * Citations [n] are converted to superscript format
 * @param {string} text - Input text
 * @returns {Array} Array of TextRun and Math objects
 */
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

  const children = [];
  let lastIndex = 0;
  let m;
  while ((m = combinedPattern.exec(text)) !== null) {
    if (m.index > lastIndex) {
      children.push(new TextRun(text.slice(lastIndex, m.index)));
    }
    if (m[1]) {
      children.push(new TextRun({ text: m[1], superScript: true }));
    } else {
      const mc = m[2] || m[3] || m[4] || m[5] || m[6] || m[7] || m[8] || m[9] || m[10] || m[11] || m[12] || m[13] || m[14] || m[15];
      if (mc) {
        const latex = unicodeToLatex(mc);
        try {
          const mathml = temml.renderToString(latex, { displayMode: false, throwOnError: false });
          const mathChildren = mathmlToDocxChildren(mathml);
          if (mathChildren && mathChildren.length) {
            children.push(new Math({ children: mathChildren }));
          } else {
            children.push(new Math({ children: [new MathRun(mc)] }));
          }
        } catch (e) {
          children.push(new Math({ children: [new MathRun(mc)] }));
        }
      }
    }
    lastIndex = m.index + m[0].length;
  }
  if (lastIndex < text.length) children.push(new TextRun(text.slice(lastIndex)));
  if (children.length === 0) children.push(new TextRun(text));
  return children;
}

// ─────────────────────────────────────────────────────────────────────────────
// Helper: body paragraph (首行缩进 2 字符, 单倍行距) - supports inline formulas
// ─────────────────────────────────────────────────────────────────────────────

function body(text) {
  // Detect if contains math content or citations
  if (containsMath(text) || containsCitation(text)) {
    return new Paragraph({
      children: parseInlineContentWithCitations(text),
    });
  }
  // Plain text paragraph
  return new Paragraph({
    children: [new TextRun(text)],
  });
}

// Paragraph with multiple TextRuns
function bodyMulti(runs) {
  return new Paragraph({
    children: runs,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Heading numbering — H1 uses Chinese numerals in text, H2/H3 use Word auto-numbering
// Per-chapter numbering references (sections_c1, sections_c2, ...) ensure H2/H3
// reset at each chapter boundary. Adding/removing sections within a chapter
// auto-updates all numbers.
// ─────────────────────────────────────────────────────────────────────────────

/** H1 — Chapter heading with Chinese numeral in text (一、二、三...) */
function h1Chinese(text) {
  _chapter++;
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}

/** H2 — Section heading with Word auto-numbering (1, 2, 3 per chapter) */
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    numbering: { reference: `sections_c${_chapter}`, level: 0 },
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}

/** H3 — Subsection heading with Word auto-numbering (1.1, 1.2 per chapter) */
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    numbering: { reference: `sections_c${_chapter}`, level: 1 },
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Helper: captions — using Word SEQ fields for per-chapter auto-numbering
// Figure captions render as "图 章-序" (e.g., 图 1-1, 图 3-2)
// Table captions render as "表 章-序" (e.g., 表 1-1, 表 2-3)
// ─────────────────────────────────────────────────────────────────────────────

/** Figure caption with per-chapter SEQ auto-numbering */
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

/** Table caption with per-chapter SEQ auto-numbering */
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

// ─────────────────────────────────────────────────────────────────────────────
// ISSUE 1 FIX: 三线表 (three-line table) with proper border handling
// Body row borders must be NONE (not just thin), only:
// - Header top: THICK
// - Header bottom: THIN
// - Last row bottom: THICK
// ─────────────────────────────────────────────────────────────────────────────

function threeLineTable(headers, rows, colWidths) {
  const n = headers.length;

  // Default: equal column widths
  if (!colWidths) {
    const w = Math.floor(CONTENT_W / n);
    colWidths = Array(n).fill(w);
    colWidths[n - 1] = CONTENT_W - w * (n - 1);
  }

  if (colWidths.length !== n) throw new Error('colWidths length must match headers length');

  // Cell helper function - supports math content in cells
  const cellOf = (text, w, borders, bold = false) => {
    // Detect if contains math content
    let cellChildren;
    if (containsMath(text)) {
      cellChildren = parseInlineContent(text);
      // If bold needed, add bold property to TextRuns
      if (bold) {
        cellChildren = cellChildren.map(child => {
          if (child instanceof TextRun) {
            return new TextRun({ text: child.text || '', bold: true });
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

  // Header row: thick top, thin bottom, no sides
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) =>
      cellOf(h, colWidths[i], { top: THICK, bottom: THIN, left: NONE, right: NONE }, true)
    ),
  });

  // Body rows: NO borders except last gets thick bottom
  // ISSUE 1 FIX: All body row borders are NONE, only last row bottom is THICK
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

// ─────────────────────────────────────────────────────────────────────────────
// Helper: reference entry (GB/T 7714-2015)
// ─────────────────────────────────────────────────────────────────────────────

function ref(text) {
  return new Paragraph({
    style: 'Reference',
    numbering: { reference: 'references', level: 0 },
    children: [new TextRun(text)],
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Helper: empty paragraph (spacing)
// ─────────────────────────────────────────────────────────────────────────────

function blank() {
  return new Paragraph({ children: [] });
}

// ─────────────────────────────────────────────────────────────────────────────
// Markdown detection helpers — strip manual numbering from input text
// ─────────────────────────────────────────────────────────────────────────────

function stripH1Number(text) {
  return text.replace(/^(?:[一二三四五六七八九十]+)[、．.]\s*/, '');
}

function stripH2Number(text) {
  return text.replace(/^\d+\.\d+\s+/, '');
}

function stripCaptionNumber(text) {
  return text.replace(/^(?:图|表)\s*\d+[-–—]\d+\s*/, '');
}

// ─────────────────────────────────────────────────────────────────────────────
// ISSUE 10 FIX: Page break helper
// Insert after abstract/keywords, and before references section
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Create a page break paragraph
 * Use after keywords section and before references section
 * @returns {Paragraph} Paragraph containing a page break
 */
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}


// ─────────────────────────────────────────────────────────────────────────────
// ISSUE 5 FIX: Block formula using temml + mathmlToDocxChildren (Word native math)
// ISSUE 2 FIX: Formula table borders all set to NONE including insideHorizontal/insideVertical
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Convert LaTeX formula to docx Math component
 * @param {string} latex - LaTeX formula string
 * @returns {Math} docx Math object
 */
function latexToMath(latex) {
  try {
    const mathml = temml.renderToString(latex, { displayMode: true, throwOnError: false });
    const children = mathmlToDocxChildren(mathml);
    if (children && children.length) {
      return new Math({ children });
    }
  } catch (e) {
    console.warn(`[formula] LaTeX parse error: ${latex}`, e.message);
  }
  // Fallback: return plain text
  return new Math({ children: [new MathRun(latex)] });
}

/**
 * Block formula layout using 3-column borderless table
 * ISSUE 2 FIX: All borders including insideHorizontal/insideVertical set to NONE
 * @param {string} latex - LaTeX formula string
 * @param {number|string} number - Equation number
 * @returns {Table} Formula table
 */
function formula(latex, number) {
  // Use 3-column borderless table layout: left margin | centered formula | right-aligned number
  const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE };
  
  const leftCell = new TableCell({
    width: { size: 567, type: WidthType.DXA },
    borders: noBorders,
    shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,  // ISSUE 11 FIX: Center content vertically
    children: [new Paragraph({ indent: { firstLine: 0 }, children: [] })],
  });
  
  // Use temml + mathmlToDocxChildren to create Word native formula
  const mathObj = latexToMath(latex);
  const formulaCell = new TableCell({
    width: { size: 7936, type: WidthType.DXA },
    borders: noBorders,
    shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,  // ISSUE 11 FIX: Center content vertically
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
    verticalAlign: VerticalAlign.CENTER,  // ISSUE 11 FIX: Center content vertically
    children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      indent: { firstLine: 0 },
      children: [new TextRun(`(${number})`)],
    })],
  });
  
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [567, 7936, 567],
    // ISSUE 2 FIX: Include insideHorizontal and insideVertical as NONE
    borders: {
      top: NONE,
      bottom: NONE,
      left: NONE,
      right: NONE,
      insideHorizontal: NONE,
      insideVertical: NONE,
    },
    rows: [new TableRow({ 
      children: [leftCell, formulaCell, numberCell],
    })],
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Document structure: styles + numbering
// ISSUE 4/6 FIX: Mixed fonts for headings/captions (Cambria Math for English/numbers)
// ─────────────────────────────────────────────────────────────────────────────

const STYLES = {
  default: {
    document: {
      // English/math: Cambria Math. Substitute "Times New Roman" if preferred.
      run: {
        font: { ascii: 'Cambria Math', hAnsi: 'Cambria Math', eastAsia: 'SimSun' },
        size: 24,  // 12pt (half-points)
      },
      // Line spacing: single. Alternatives: fixed 20pt → line:400,EXACT | 1.5x → line:360,AUTO
      paragraph: {
        spacing: { line: 240, lineRule: LineRuleType.AUTO },
        indent:  { firstLine: 480 },  // 2-character indent
      },
    },
  },
  paragraphStyles: [
    {
      // ISSUE 4/6 FIX: Heading fonts use Cambria Math for ascii/hAnsi, SimHei for eastAsia
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
      // ISSUE 4 FIX: Figure/Table captions use Cambria Math for English/numbers
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

/**
 * Build numbering config dynamically based on chapter count.
 * Each chapter gets its own numbering reference (sections_c1, sections_c2, ...)
 * so H2/H3 auto-numbering resets per chapter.
 * @param {number} chapterCount - Total number of chapters in the document
 * @returns {object} Numbering config for Document constructor
 */
function buildNumberingConfig(chapterCount) {
  const configs = [
    {
      reference: 'references',
      levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: '[%1]',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 480, hanging: 480 } } } },
      ],
    },
    {
      reference: 'bullets',
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: '•',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
      ],
    },
    {
      reference: 'numbers',
      levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: '%1.',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
      ],
    },
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
// ██  CONTENT SECTION — Edit below this line  ████████████████████████████████
// ─────────────────────────────────────────────────────────────────────────────

const CONTENT = [
  // ── Example: Title (manual, no heading style) ──────────────────────────────
  // ISSUE 6 FIX: Title uses Cambria Math for English/numbers
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing:   { before: 0, after: 240 },
    indent:    { firstLine: 0 },
    children:  [new TextRun({ text: '论文标题', bold: true, size: 36,
                              font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' } })],
  }),

  // ── Abstract and Keywords ─────────────────────────────────────────────────
  new Paragraph({
    alignment: AlignmentType.CENTER,
    indent:    { firstLine: 0 },
    children:  [new TextRun({ text: '摘要', bold: true })],
  }),
  body('本文研究了强化学习中的Q-learning算法，分析了其收敛性和应用场景。'),
  new Paragraph({
    indent:    { firstLine: 0 },
    children:  [
      new TextRun({ text: '关键词：', bold: true }),
      new TextRun('强化学习；Q-learning；马尔可夫决策过程'),
    ],
  }),
  // ISSUE 10 FIX: Page break after keywords (abstract ends here)
  pageBreak(),

  // ── Table of Contents (optional) ───────────────────────────────────────────
  // new TableOfContents('目录', { hyperlink: true, headingStyleRange: '1-3' }),
  // pageBreak(),

  // ── Section 1 ──────────────────────────────────────────────────────────────
  h1Chinese('一、引言'),
  blank(),
  body('强化学习（Reinforcement Learning）是机器学习的一个重要分支。'),

  h2('研究背景'),
  blank(),
  body('近年来，深度强化学习取得了显著进展。[1][2]'),  // ISSUE 8: Citations become superscript

  h3('研究现状'),
  blank(),
  body('目前已有多种经典算法被提出并验证。'),

  // ── Formula example ────────────────────────────────────────────────────────
  blank(),
  body('Q-Learning 更新规则如下:'),
  blank(),
  formula('Q_n(x, a) = (1 - \\alpha_n) Q_{n-1}(x, a) + \\alpha_n [r_n + \\gamma V_{n-1}(y_n)]', 1),
  blank(),

  // ── Table example ──────────────────────────────────────────────────────────
  blank(),
  tableCaption('符号说明'),
  threeLineTable(
    ['符号', '说明'],
    [
      ['S',   '状态空间，表示所有可能状态的集合'],
      ['A',   '动作空间，表示所有可执行动作的集合'],
      ['Qₙ(x,a)', 'Q值函数，第n次迭代时状态x下动作a的价值'],  // Math in table cell
    ],
    [1800, 7270]  // must sum to CONTENT_W = 9070
  ),
  blank(),

  // ── Section 2 ──────────────────────────────────────────────────────────────
  h1Chinese('二、方法'),
  blank(),
  body('本文提出一种改进的 Q-learning 算法。'),


  // ISSUE 10 FIX: Page break before references (new page for references)
  pageBreak(),

  // ── References ─────────────────────────────────────────────────────────
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    indent:  { firstLine: 0 },
    children: [new TextRun('参考文献')],
    // No numbering on References heading — write it manually
  }),
  blank(),
  ref('Watkins C J C H, Dayan P. Q-learning[J]. Machine learning, 1992, 8(3): 279-292.'),
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
  const out = path.resolve(OUTPUT_PATH);
  fs.writeFileSync(out, buf);
  console.log(`✓  Written: ${out}`);
}).catch(err => {
  console.error('Error building document:', err.message);
  process.exit(1);
});
