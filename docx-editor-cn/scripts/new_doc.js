/**
 * new_doc.js — Chinese Academic Paper Template (课程论文模板)
 *
 * All formatting is pre-configured to GB/T 7714 + Chinese university standards:
 *   - A4 page, 2.5 cm margins on all sides
 *   - SimSun 12pt body, SimHei headings (Cambria Math for English/numbers)
 *   - Manual multi-level heading numbering (synchronized counters)
 *   - Three-line table helper with proper border handling
 *   - LaTeX formula support (block and inline) via temml + Word native math
 *   - Citation superscript handling [n] format
 *   - Reference list [1][2][3] numbering
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
  LevelFormat, BorderStyle, WidthType, ShadingType, VerticalAlign,
  TableOfContents, PageBreak, ImageRun,
} = require('docx');

// MathML to docx Math converter
const { mathmlToDocxChildren } = require('./mathml-to-docx');
const temml = require('temml');

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

const INPUT_MARKDOWN = process.argv[2] ? path.resolve(process.argv[2]) : null;
const OUTPUT_PATH = process.argv[3]
  ? path.resolve(process.argv[3])
  : (INPUT_MARKDOWN
      ? path.join(path.dirname(INPUT_MARKDOWN), `${path.parse(INPUT_MARKDOWN).name}.docx`)
      : 'output.docx');

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

// ISSUE 3 FIX: Manual heading counters for synchronized numbering
let currentChapter = 0;
let currentSection = 0;
let currentSubsection = 0;

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
  'ₙ': '_n', 'ₓ': '_x', 'ᵢ': '_i', 'ₜ': '_t', 'ₛ': '_s',
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
  '⟨': '\\langle', '⟩': '\\rangle',
  // Superscript letter (for π*, Q*, etc.)
  '*': '^*',
};

/**
 * Convert text with Unicode math symbols to LaTeX format
 * @param {string} text - Text with Unicode math symbols
 * @returns {string} LaTeX formatted text
 */
function unicodeToLatex(text) {
  let result = text;
  for (const [unicode, latex] of Object.entries(UNICODE_TO_LATEX)) {
    result = result.split(unicode).join(latex);
  }
  return result;
}

/**
 * Detect if text contains math content (needs formula editor rendering)
 * ISSUE 7 FIX: Strict detection - don't match plain numbers or English words
 * @param {string} text - Text to detect
 * @returns {boolean}
 */
function containsMath(text) {
  // Detect Greek letters
  if (/[αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ]/.test(text)) return true;
  // Detect Unicode subscript/superscript characters
  if (/[₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]/.test(text)) return true;
  // Detect math operators and special symbols
  if (/[∞∑∏∫≤≥≠≈→←↔∈∉⊂⊃∀∃∧∨×÷±∓·…⋯′″⟨⟩]/.test(text)) return true;
  // Detect starred symbols like π*, Q* (but not plain words like Agent)
  if (/[A-Z]\*/.test(text)) return true;
  // Detect $...$ LaTeX delimiters
  if (/\$[^$]+\$/.test(text)) return true;
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
  const children = [];
  
  // Regex matching math content blocks (in priority order):
  // 1. $...$  explicit LaTeX (highest priority)
  // 2. Function form Q(s,a) V(s) Rₓ(a) etc. (only with subscripts or specific single letters)
  // 3. Greek letters (alone or with subscripts)
  // 4. Variables with subscripts like Qₙ xₙ αₙ etc.
  // 5. Starred symbols like π* Q* (single letter only)
  // 
  // NOTE: Does NOT match plain English words like Agent, Watkins, Dayan
  // Does NOT match plain numbers like 1992, 500
  // Does NOT match plain parentheses expressions like (1992)
  
  const mathPattern = /\$([^$]+)\$|([A-Z][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]+\*?\s*\([^)]+\))|([A-Z]\s*\([^)]*[αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ][^)]*\))|([αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]*\*?)|([A-Za-z][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]+\*?)|([A-Z]\*)/g;
  
  let lastIndex = 0;
  let match;
  
  while ((match = mathPattern.exec(text)) !== null) {
    // Add plain text before match
    if (match.index > lastIndex) {
      const plainText = text.slice(lastIndex, match.index);
      if (plainText) {
        children.push(new TextRun(plainText));
      }
    }
    
    // Get matched math content
    const mathContent = match[1] || match[2] || match[3] || match[4] || match[5] || match[6];
    if (mathContent) {
      // Convert Unicode to LaTeX and create Math object
      const latex = unicodeToLatex(mathContent);
      try {
        const mathml = temml.renderToString(latex, { displayMode: false, throwOnError: false });
        const mathChildren = mathmlToDocxChildren(mathml);
        if (mathChildren && mathChildren.length) {
          children.push(new Math({ children: mathChildren }));
        } else {
          // fallback
          children.push(new Math({ children: [new MathRun(mathContent)] }));
        }
      } catch (e) {
        // Parse failed, use MathRun to display original text
        children.push(new Math({ children: [new MathRun(mathContent)] }));
      }
    }
    
    lastIndex = match.index + match[0].length;
  }
  
  // Add remaining plain text
  if (lastIndex < text.length) {
    children.push(new TextRun(text.slice(lastIndex)));
  }
  
  // If no math content matched, return plain text
  if (children.length === 0) {
    children.push(new TextRun(text));
  }
  
  return children;
}

/**
 * ISSUE 8 FIX: Parse text with math content and citations
 * Citations [n] are converted to superscript format
 * @param {string} text - Input text
 * @returns {Array} Array of TextRun and Math objects
 */
function parseInlineContentWithCitations(text) {
  const children = [];
  
  // Combined regex: match math content or citations
  // Citations [n] become superscript
  // Math content becomes Math objects
  const combinedPattern = /(\[\d+\])|\$([^$]+)\$|([A-Z][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]+\*?\s*\([^)]+\))|([A-Z]\s*\([^)]*[αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ][^)]*\))|([αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]*\*?)|([A-Za-z][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]+\*?)|([A-Z]\*)/g;
  
  let lastIndex = 0;
  let match;
  
  while ((match = combinedPattern.exec(text)) !== null) {
    // Add plain text before match
    if (match.index > lastIndex) {
      const plainText = text.slice(lastIndex, match.index);
      if (plainText) {
        children.push(new TextRun(plainText));
      }
    }
    
    if (match[1]) {
      // Citation [n] - convert to superscript
      children.push(new TextRun({
        text: match[1],
        superScript: true,
      }));
    } else {
      // Math content
      const mathContent = match[2] || match[3] || match[4] || match[5] || match[6] || match[7];
      if (mathContent) {
        const latex = unicodeToLatex(mathContent);
        try {
          const mathml = temml.renderToString(latex, { displayMode: false, throwOnError: false });
          const mathChildren = mathmlToDocxChildren(mathml);
          if (mathChildren && mathChildren.length) {
            children.push(new Math({ children: mathChildren }));
          } else {
            children.push(new Math({ children: [new MathRun(mathContent)] }));
          }
        } catch (e) {
          children.push(new Math({ children: [new MathRun(mathContent)] }));
        }
      }
    }
    
    lastIndex = match.index + match[0].length;
  }
  
  // Add remaining plain text
  if (lastIndex < text.length) {
    children.push(new TextRun(text.slice(lastIndex)));
  }
  
  if (children.length === 0) {
    children.push(new TextRun(text));
  }
  
  return children;
}

// ─────────────────────────────────────────────────────────────────────────────
// Helper: body paragraph (首行缩进 2 字符, 单倍行距) - supports inline formulas
// ─────────────────────────────────────────────────────────────────────────────

function makeBodyParagraph(text, overrides = {}) {
  const children = containsMath(text) || containsCitation(text)
    ? parseInlineContentWithCitations(text)
    : [new TextRun(text)];

  return new Paragraph({
    ...overrides,
    children,
  });
}

function body(text) {
  return makeBodyParagraph(text);
}

function bodyIndented(text, left = 0) {
  return makeBodyParagraph(text, {
    indent: { left, firstLine: 0 },
  });
}

function bodyNoIndent(text, overrides = {}) {
  return makeBodyParagraph(text, {
    indent: { firstLine: 0 },
    ...overrides,
  });
}

// Paragraph with multiple TextRuns
function bodyMulti(runs) {
  return new Paragraph({
    children: runs,
  });
}

const INLINE_LATEX_TO_UNICODE = {
  '\\alpha': 'α',
  '\\beta': 'β',
  '\\gamma': 'γ',
  '\\delta': 'δ',
  '\\epsilon': 'ε',
  '\\varepsilon': 'ε',
  '\\zeta': 'ζ',
  '\\eta': 'η',
  '\\theta': 'θ',
  '\\vartheta': 'ϑ',
  '\\iota': 'ι',
  '\\kappa': 'κ',
  '\\lambda': 'λ',
  '\\mu': 'μ',
  '\\nu': 'ν',
  '\\xi': 'ξ',
  '\\pi': 'π',
  '\\rho': 'ρ',
  '\\sigma': 'σ',
  '\\tau': 'τ',
  '\\upsilon': 'υ',
  '\\phi': 'φ',
  '\\varphi': 'φ',
  '\\chi': 'χ',
  '\\psi': 'ψ',
  '\\omega': 'ω',
  '\\Gamma': 'Γ',
  '\\Delta': 'Δ',
  '\\Theta': 'Θ',
  '\\Lambda': 'Λ',
  '\\Xi': 'Ξ',
  '\\Pi': 'Π',
  '\\Sigma': 'Σ',
  '\\Phi': 'Φ',
  '\\Psi': 'Ψ',
  '\\Omega': 'Ω',
  '\\langle': '⟨',
  '\\rangle': '⟩',
  '\\leq': '≤',
  '\\geq': '≥',
  '\\neq': '≠',
  '\\to': '→',
  '\\rightarrow': '→',
  '\\leftarrow': '←',
  '\\cdot': '·',
  '\\times': '×',
  '\\infty': '∞',
};

const SUBSCRIPT_CHARS = {
  '0': '₀',
  '1': '₁',
  '2': '₂',
  '3': '₃',
  '4': '₄',
  '5': '₅',
  '6': '₆',
  '7': '₇',
  '8': '₈',
  '9': '₉',
  '+': '₊',
  '-': '₋',
  '=': '₌',
  '(': '₍',
  ')': '₎',
  'a': 'ₐ',
  'e': 'ₑ',
  'h': 'ₕ',
  'i': 'ᵢ',
  'j': 'ⱼ',
  'k': 'ₖ',
  'l': 'ₗ',
  'm': 'ₘ',
  'n': 'ₙ',
  'o': 'ₒ',
  'p': 'ₚ',
  'r': 'ᵣ',
  's': 'ₛ',
  't': 'ₜ',
  'u': 'ᵤ',
  'v': 'ᵥ',
  'x': 'ₓ',
  'β': 'ᵦ',
  'γ': 'ᵧ',
  'ρ': 'ᵨ',
  'φ': 'ᵩ',
  'χ': 'ᵪ',
};

const SUPERSCRIPT_CHARS = {
  '0': '⁰',
  '1': '¹',
  '2': '²',
  '3': '³',
  '4': '⁴',
  '5': '⁵',
  '6': '⁶',
  '7': '⁷',
  '8': '⁸',
  '9': '⁹',
  '+': '⁺',
  '-': '⁻',
  '=': '⁼',
  '(': '⁽',
  ')': '⁾',
  'i': 'ⁱ',
  'n': 'ⁿ',
};

function convertCharsWithMap(text, mapping) {
  return [...text].map((char) => mapping[char] || char).join('');
}

function unwrapLatexTextCommands(text) {
  let result = text;
  let previous = '';

  while (result !== previous) {
    previous = result;
    result = result.replace(
      /\\(?:mathrm|mathbf|mathit|text|operatorname)\{([^{}]*)\}/g,
      '$1'
    );
  }

  return result;
}

function simplifyInlineLatex(latex) {
  let text = latex.trim();
  text = unwrapLatexTextCommands(text);

  for (const [command, replacement] of Object.entries(INLINE_LATEX_TO_UNICODE)) {
    text = text.split(command).join(replacement);
  }

  text = text
    .replace(/\\left/g, '')
    .replace(/\\right/g, '')
    .replace(/\\,/g, ' ')
    .replace(/\\;/g, ' ')
    .replace(/\\!/g, '')
    .replace(/\\ /g, ' ')
    .replace(/\\\{/g, '{')
    .replace(/\\\}/g, '}');

  text = text
    .replace(/_\{([^{}]+)\}/g, (_, group) => convertCharsWithMap(group, SUBSCRIPT_CHARS))
    .replace(/_([A-Za-z0-9+\-=()α-ωΑ-Ω])/g, (_, group) => convertCharsWithMap(group, SUBSCRIPT_CHARS))
    .replace(/\^\{([^{}]+)\}/g, (_, group) => group === '*' ? '*' : convertCharsWithMap(group, SUPERSCRIPT_CHARS))
    .replace(/\^([A-Za-z0-9+\-=()α-ωΑ-Ω*])/g, (_, group) => group === '*' ? '*' : convertCharsWithMap(group, SUPERSCRIPT_CHARS));

  text = text
    .replace(/[{}]/g, '')
    .replace(/\s+/g, ' ')
    .trim();

  return text;
}

function buildInlineMathRuns(text, options = {}) {
  const {
    bold = false,
    normalEastAsia = 'SimHei',
    normalItalic = false,
    mathItalic = true,
  } = options;

  const runs = [];
  const pattern = /\$([^$]+)\$/g;
  let lastIndex = 0;
  let match;

  const normalRun = (value) => new TextRun({
    text: value,
    bold,
    italics: normalItalic,
    font: { ascii: 'Cambria Math', eastAsia: normalEastAsia, hAnsi: 'Cambria Math' },
  });

  const mathRun = (value) => new TextRun({
    text: simplifyInlineLatex(value),
    bold,
    italics: mathItalic,
    font: { ascii: 'Cambria Math', eastAsia: normalEastAsia, hAnsi: 'Cambria Math' },
  });

  while ((match = pattern.exec(text)) !== null) {
    if (match.index > lastIndex) {
      runs.push(normalRun(text.slice(lastIndex, match.index)));
    }

    if (match[1].trim()) {
      runs.push(mathRun(match[1]));
    }

    lastIndex = match.index + match[0].length;
  }

  if (lastIndex < text.length) {
    runs.push(normalRun(text.slice(lastIndex)));
  }

  if (runs.length === 0) {
    runs.push(normalRun(text));
  }

  return runs;
}

function plainHeading(level, text) {
  return new Paragraph({
    heading: level,
    indent: { firstLine: 0 },
    children: buildInlineMathRuns(text, { bold: true, normalEastAsia: 'SimHei' }),
  });
}

function centeredText(text, options = {}) {
  const {
    size = 24,
    bold = false,
    spacing = { before: 0, after: 120 },
  } = options;

  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing,
    indent: { firstLine: 0 },
    children: [new TextRun({
      text,
      bold,
      size,
      font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' },
    })],
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    children: containsMath(text) || containsCitation(text)
      ? parseInlineContentWithCitations(text)
      : [new TextRun(text)],
  });
}

function parseHtmlTable(html) {
  const rowMatches = [...html.matchAll(/<tr>(.*?)<\/tr>/gsi)];
  const rows = rowMatches.map((rowMatch) => {
    const cellMatches = [...rowMatch[1].matchAll(/<t[dh]>(.*?)<\/t[dh]>/gsi)];
    return cellMatches.map((cellMatch) =>
      cellMatch[1]
        .replace(/<[^>]+>/g, '')
        .replace(/&nbsp;/g, ' ')
        .trim()
    );
  }).filter((row) => row.length);

  if (rows.length < 2) {
    throw new Error('HTML table requires at least one header row and one body row');
  }

  return {
    headers: rows[0],
    rows: rows.slice(1),
  };
}

function normalizeCaptionText(text) {
  return text
    .replace(/^([\u56fe\u8868])(\d)/, '$1 $2')
    .replace(/\s+/g, ' ')
    .trim();
}

function getImageType(imagePath) {
  const ext = path.extname(imagePath).toLowerCase();
  if (ext === '.jpg') return 'jpg';
  if (ext === '.jpeg') return 'jpeg';
  if (ext === '.png') return 'png';
  if (ext === '.gif') return 'gif';
  if (ext === '.bmp') return 'bmp';
  throw new Error(`Unsupported image format: ${imagePath}`);
}

function getImageDimensions(buffer, imagePath) {
  const ext = path.extname(imagePath).toLowerCase();

  if (ext === '.png') {
    return {
      width: buffer.readUInt32BE(16),
      height: buffer.readUInt32BE(20),
    };
  }

  if (ext === '.jpg' || ext === '.jpeg') {
    let offset = 2;
    while (offset < buffer.length) {
      if (buffer[offset] !== 0xFF) {
        offset += 1;
        continue;
      }
      const marker = buffer[offset + 1];
      const length = buffer.readUInt16BE(offset + 2);
      if (marker >= 0xC0 && marker <= 0xC3) {
        return {
          height: buffer.readUInt16BE(offset + 5),
          width: buffer.readUInt16BE(offset + 7),
        };
      }
      offset += 2 + length;
    }
  }

  throw new Error(`Unable to read image dimensions: ${imagePath}`);
}

function imageBlock(imagePath) {
  const data = fs.readFileSync(imagePath);
  const { width, height } = getImageDimensions(data, imagePath);
  const maxWidth = 460;
  const maxHeight = 360;
  const scale = globalThis.Math.min(maxWidth / width, maxHeight / height, 1);

  return new Paragraph({
    alignment: AlignmentType.CENTER,
    indent: { firstLine: 0 },
    spacing: { before: 120, after: 60 },
    children: [new ImageRun({
      type: getImageType(imagePath),
      data,
      transformation: {
        width: globalThis.Math.round(width * scale),
        height: globalThis.Math.round(height * scale),
      },
      altText: {
        title: path.basename(imagePath),
        description: path.basename(imagePath),
        name: path.basename(imagePath),
      },
    })],
  });
}

function parseFormulaBlock(lines, startIndex) {
  const formulaLines = [];
  let i = startIndex + 1;
  while (i < lines.length && lines[i].trim() !== '$$') {
    formulaLines.push(lines[i]);
    i += 1;
  }

  const rawLatex = formulaLines.join(' ').trim();
  const tagMatch = rawLatex.match(/\\tag\{([^}]+)\}\s*$/);
  const number = tagMatch ? tagMatch[1] : '';
  const latex = tagMatch ? rawLatex.slice(0, tagMatch.index).trim() : rawLatex;

  return {
    block: formula(latex, number),
    nextIndex: i,
  };
}

function buildContentFromMarkdown(markdownPath) {
  const raw = fs.readFileSync(markdownPath, 'utf8');
  const lines = raw.split(/\r?\n/);
  const content = [];
  const baseDir = path.dirname(markdownPath);
  let i = 0;
  let inReferences = false;
  let abstractInserted = false;
  let generatedToc = false;
  let mainHeadingCount = 0;

  const flushParagraph = (paragraphLines) => {
    if (!paragraphLines.length) return;
    const rawText = paragraphLines.join(' ').replace(/\s+/g, ' ').trim();
    if (!rawText) return;

    const leadingTabs = (paragraphLines[0].match(/^\t+/) || [''])[0].length;
    const leadingSpaces = (paragraphLines[0].match(/^ +/) || [''])[0].length;
    const left = leadingTabs > 0 ? leadingTabs * 420 : (leadingSpaces >= 4 ? 420 : 0);
    content.push(left > 0 ? bodyIndented(rawText, left) : body(rawText));
  };

  while (i < lines.length) {
    const rawLine = lines[i];
    const line = rawLine.trim();

    if (!line) {
      i += 1;
      continue;
    }

    if (line === '## \u76ee\u5f55') {
      i += 1;
      while (i < lines.length && !lines[i].startsWith('# ')) {
        i += 1;
      }
      continue;
    }

    if (line.startsWith('# ')) {
      content.push(centeredText(line.slice(2).trim(), {
        size: 36,
        bold: true,
        spacing: { before: 0, after: 240 },
      }));
      i += 1;
      continue;
    }

    if (/^##\s+.+\(name\)\s*$/.test(line)) {
      content.push(centeredText(line.replace(/^##\s+/, ''), {
        size: 24,
        spacing: { before: 0, after: 180 },
      }));
      i += 1;
      continue;
    }

    if (/^##\s*\u6458\u8981[:\uff1a]?\s*$/.test(line)) {
      content.push(centeredText('\u6458\u8981', { bold: true, spacing: { before: 0, after: 120 } }));
      i += 1;
      abstractInserted = true;
      continue;
    }

    if (/^\u5173\u952e\u8bcd[:\uff1a]/.test(line)) {
      content.push(new Paragraph({
        indent: { firstLine: 0 },
        children: [
          new TextRun({ text: '\u5173\u952e\u8bcd\uff1a', bold: true }),
          ...parseInlineContentWithCitations(line.replace(/^\u5173\u952e\u8bcd[:\uff1a]\s*/, '')),
        ],
      }));
      if (abstractInserted && !generatedToc) {
        content.push(pageBreak());
        content.push(new TableOfContents('\u76ee\u5f55', { hyperlink: true, headingStyleRange: '1-3' }));
        content.push(pageBreak());
        generatedToc = true;
      }
      i += 1;
      continue;
    }

    if (line === '$$') {
      const { block, nextIndex } = parseFormulaBlock(lines, i);
      content.push(block);
      i = nextIndex + 1;
      continue;
    }

    const imageMatch = line.match(/^!\[[^\]]*]\(([^)]+)\)$/);
    if (imageMatch) {
      const imagePath = path.resolve(baseDir, imageMatch[1]);
      content.push(imageBlock(imagePath));
      let j = i + 1;
      while (j < lines.length && !lines[j].trim()) {
        j += 1;
      }
      if (j < lines.length && /^\u56fe/.test(lines[j].trim())) {
        content.push(figCaption(normalizeCaptionText(lines[j].trim())));
        i = j + 1;
      } else {
        i += 1;
      }
      continue;
    }

    if (line.startsWith('<table>')) {
      const { headers, rows } = parseHtmlTable(line);
      if (content.length && i > 0 && /^\u8868/.test(lines[i - 1].trim())) {
        content.pop();
        content.push(tableCaption(normalizeCaptionText(lines[i - 1].trim())));
      }
      const columnWidths = headers.length === 2
        ? [1800, CONTENT_W - 1800]
        : Array(headers.length).fill(globalThis.Math.floor(CONTENT_W / headers.length));
      if (headers.length > 2) {
        columnWidths[columnWidths.length - 1] = CONTENT_W - columnWidths.slice(0, -1).reduce((sum, n) => sum + n, 0);
      }
      content.push(threeLineTable(headers, rows, columnWidths));
      i += 1;
      continue;
    }

    if (line === '---') {
      i += 1;
      continue;
    }

    if (/^##\s+\u53c2\u8003\u6587\u732e\s*$/.test(line)) {
      inReferences = true;
      content.push(pageBreak());
      content.push(plainHeading(HeadingLevel.HEADING_1, '\u53c2\u8003\u6587\u732e'));
      i += 1;
      continue;
    }

    if (/^##\s+/.test(line)) {
      const text = line.replace(/^##\s+/, '').trim();
      if (/^[\u4e00-\u5341]+\u3001/.test(text)) {
        if (mainHeadingCount > 0) {
          content.push(pageBreak());
        }
        mainHeadingCount += 1;
      }
      content.push(plainHeading(HeadingLevel.HEADING_1, text));
      inReferences = false;
      i += 1;
      continue;
    }

    if (/^###\s+/.test(line)) {
      content.push(plainHeading(HeadingLevel.HEADING_2, line.replace(/^###\s+/, '').trim()));
      i += 1;
      continue;
    }

    if (/^####\s+/.test(line)) {
      content.push(plainHeading(HeadingLevel.HEADING_3, line.replace(/^####\s+/, '').trim()));
      i += 1;
      continue;
    }

    if (/^\[\d+\]\s+/.test(line) && inReferences) {
      content.push(ref(line.replace(/^\[\d+\]\s+/, '').trim()));
      i += 1;
      continue;
    }

    if (/^- /.test(line)) {
      content.push(bullet(line.replace(/^- /, '').trim()));
      i += 1;
      continue;
    }

    const paragraphLines = [rawLine];
    let j = i + 1;
    while (j < lines.length) {
      const nextLine = lines[j];
      const nextTrimmed = nextLine.trim();
      if (
        !nextTrimmed ||
        nextTrimmed === '$$' ||
        nextTrimmed === '---' ||
        /^#/.test(nextTrimmed) ||
        /^!\[/.test(nextTrimmed) ||
        nextTrimmed.startsWith('<table>') ||
        /^- /.test(nextTrimmed) ||
        (/^\[\d+\]\s+/.test(nextTrimmed) && inReferences)
      ) {
        break;
      }
      paragraphLines.push(nextLine);
      j += 1;
    }
    flushParagraph(paragraphLines);
    i = j;
  }

  return content;
}

// ?????????????????????????????????????????????????????????????????????????????
// ISSUE 3 FIX: Manual heading numbering for synchronized counters
// ?????????????????????????????????????????????????????????????????????????????

/** Reset heading counters (call at start of document) */
function resetHeadingCounters() {
  currentChapter = 0;
  currentSection = 0;
  currentSubsection = 0;
}

/** H1 - Level 1 heading (manual Chinese numbering: 一、二、三) */
function h1Manual(text) {
  currentChapter++;
  currentSection = 0;
  currentSubsection = 0;
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    indent: { firstLine: 0 },
    children: buildInlineMathRuns(text, { bold: true, normalEastAsia: 'SimHei' }),
  });
}

/** H1 with auto-numbering (→ 1  2  3) - use when you want Arabic numerals */
function h1(text) {
  currentChapter++;
  currentSection = 0;
  currentSubsection = 0;
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    indent: { firstLine: 0 },
    children: [
      new TextRun({
        text: `${currentChapter} `,
        bold: true,
        font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' },
      }),
      ...buildInlineMathRuns(text, { bold: true, normalEastAsia: 'SimHei' }),
    ],
  });
}

/** H2 - Level 2 heading (manual numbering: chapter.section, e.g., 1.1, 2.3) */
function h2(text) {
  currentSection++;
  currentSubsection = 0;
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    indent: { firstLine: 0 },
    children: [
      new TextRun({
        text: `${currentChapter}.${currentSection} `,
        bold: true,
        font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' },
      }),
      ...buildInlineMathRuns(text, { bold: true, normalEastAsia: 'SimHei' }),
    ],
  });
}

/** H3 - Level 3 heading (manual numbering: chapter.section.subsection, e.g., 1.1.1) */
function h3(text) {
  currentSubsection++;
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    indent: { firstLine: 0 },
    children: [
      new TextRun({
        text: `${currentChapter}.${currentSection}.${currentSubsection} `,
        bold: true,
        font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' },
      }),
      ...buildInlineMathRuns(text, { bold: true, normalEastAsia: 'SimHei' }),
    ],
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Helper: captions - ISSUE 4 FIX: Mixed fonts for English/numbers
// ─────────────────────────────────────────────────────────────────────────────

/** Figure caption (below figure). label e.g. "图 1-1 系统架构" */
function figCaption(label) {
  return new Paragraph({
    style: 'FigureCaption',
    children: buildInlineMathRuns(label, { bold: true, normalEastAsia: 'SimSun' }),
  });
}

/** Table caption (above table). label e.g. "表 1-1 符号说明" */
function tableCaption(label) {
  return new Paragraph({
    style: 'TableCaption',
    children: buildInlineMathRuns(label, { bold: true, normalEastAsia: 'SimSun' }),
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
    const w = globalThis.Math.floor(CONTENT_W / n);
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
        spacing:      { before: 120, after: 120, line: 360, lineRule: LineRuleType.AUTO },
        outlineLevel: 1,
      },
    },
    {
      id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' }, size: 24, bold: true },
      paragraph: {
        alignment:    AlignmentType.LEFT,
        indent:       { firstLine: 0 },
        spacing:      { before: 120, after: 120, line: 360, lineRule: LineRuleType.AUTO },
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

const NUMBERING = {
  config: [
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
  ],
};

// ─────────────────────────────────────────────────────────────────────────────
// ██  CONTENT SECTION — Edit below this line  ████████████████████████████████
// ─────────────────────────────────────────────────────────────────────────────

// Reset counters before building content
resetHeadingCounters();

const CONTENT = INPUT_MARKDOWN
  ? buildContentFromMarkdown(INPUT_MARKDOWN)
  : [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing:   { before: 0, after: 240 },
        indent:    { firstLine: 0 },
        children:  [new TextRun({ text: '????', bold: true, size: 36,
                                  font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' } })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        indent:    { firstLine: 0 },
        children:  [new TextRun({ text: '??', bold: true })],
      }),
      body('???????????Q-learning????????????????'),
      new Paragraph({
        indent:    { firstLine: 0 },
        children:  [
          new TextRun({ text: '\u5173\u952e\u8bcd\uff1a', bold: true }),
          new TextRun('?????Q-learning?????????'),
        ],
      }),
      pageBreak(),
      h1Manual('????'),
      blank(),
      body('?????Reinforcement Learning??????????????'),
      h2('????'),
      blank(),
      body('??????????????????[1][2]'),
      h3('????'),
      blank(),
      body('?????????????????'),
      blank(),
      body('Q-Learning ??????:'),
      blank(),
      formula('Q_n(x, a) = (1 - \\alpha_n) Q_{n-1}(x, a) + \\alpha_n [r_n + \\gamma V_{n-1}(y_n)]', 1),
      blank(),
      blank(),
      tableCaption('? 1-1 ????'),
      threeLineTable(
        ['??', '??'],
        [
          ['S',   '????????????????'],
          ['A',   '?????????????????'],
          ['Q?(x,a)', 'Q?????n??????x???a???'],
        ],
        [1800, 7270]
      ),
      blank(),
      h1Manual('????'),
      blank(),
      body('????????? Q-learning ???'),
      pageBreak(),
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        indent:  { firstLine: 0 },
        children: [new TextRun('????')],
      }),
      blank(),
      ref('Watkins C J C H, Dayan P. Q-learning[J]. Machine learning, 1992, 8(3): 279-292.'),
      ref('Sutton R S, Barto A G. Reinforcement Learning: An Introduction[M]. 2nd ed. Cambridge: MIT Press, 2018.'),
    ];

// ?????????????????????????????????????????????????????????????????????????????
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
