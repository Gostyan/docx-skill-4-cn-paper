/**
 * mathml-to-docx.js - Convert MathML to docx Math components
 * Based on vace/markdown-docx (MIT License)
 */

'use strict';

const {
  MathRun, MathFraction, MathRadical, MathSuperScript, MathSubScript,
  MathSubSuperScript, MathSum, MathIntegral, XmlComponent,
} = require('docx');
const { XMLParser } = require('fast-xml-parser');

let LO_COMPAT = false;

// ─────────────────────────────────────────────────────────────────────────────
// OMML Matrix helpers
// ─────────────────────────────────────────────────────────────────────────────

class MathMatrixElement extends XmlComponent {
  constructor(children) {
    super('m:e');
    for (const child of children) this.root.push(child);
  }
}

class MathMatrixRow extends XmlComponent {
  constructor(cells) {
    super('m:mr');
    for (const cell of cells) this.root.push(new MathMatrixElement(cell));
  }
}

class MathMatrix extends XmlComponent {
  constructor(rows) {
    super('m:m');
    for (const row of rows) this.root.push(new MathMatrixRow(row));
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Main conversion function
// ─────────────────────────────────────────────────────────────────────────────

function mathmlToDocxChildren(mathml, opts) {
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '',
    textNodeName: 'text',
    preserveOrder: true,
    trimValues: false,
  });
  const json = parser.parse(mathml);
  const mathNode = findFirst(json, 'math');
  LO_COMPAT = !!(opts && opts.libreOfficeCompat);

  if (!mathNode) return [];
  const semantics = findFirst(childrenOf(mathNode), 'semantics');
  const root = semantics ? findFirst(childrenOf(semantics), 'mrow') || semantics : findFirst(childrenOf(mathNode), 'mrow') || mathNode;
  return walkChildren(childrenOf(root));
}

// ─────────────────────────────────────────────────────────────────────────────
// Walk children
// ─────────────────────────────────────────────────────────────────────────────

function walkChildren(nodes) {
  let out = [];
  for (let i = 0; i < nodes.length; i++) {
    const n = nodes[i];
    out = out.concat(walkNode(n));
  }
  return out;
}

// ─────────────────────────────────────────────────────────────────────────────
// Walk single node
// ─────────────────────────────────────────────────────────────────────────────

function walkNode(node) {
  const tag = tagName(node);
  if (!tag) {
    const t = node.text?.toString() || '';
    return t ? [new MathRun(t)] : [];
  }
  const kids = childrenOf(node);

  switch (tag) {
    case 'mrow':
      return walkChildren(kids);
      
    case 'mi':
    case 'mn':
      return textFrom(kids);
      
    case 'mo':
      return textFrom(kids);
    
    case 'mtext':
      return textFrom(kids);
      
    case 'msup': {
      const [base, sup] = firstN(kids, 2);
      return [new MathSuperScript({ children: walkNode(base), superScript: walkNode(sup) })];
    }
    
    case 'msub': {
      const [base, sub] = firstN(kids, 2);
      return [new MathSubScript({ children: walkNode(base), subScript: walkNode(sub) })];
    }
    
    case 'msubsup': {
      const [base, sub, sup] = firstN(kids, 3);
      return [new MathSubSuperScript({ children: walkNode(base), subScript: walkNode(sub), superScript: walkNode(sup) })];
    }
    
    case 'mfrac': {
      const [num, den] = firstN(kids, 2);
      return [new MathFraction({ numerator: walkNode(num), denominator: walkNode(den) })];
    }
    
    case 'msqrt': {
      const [body] = firstN(kids, 1);
      return [new MathRadical({ children: walkNode(body) })];
    }
    
    case 'mroot': {
      const [body, degree] = firstN(kids, 2);
      return [new MathRadical({ children: walkNode(body), degree: walkNode(degree) })];
    }
    
    case 'munder': {
      const [base, under] = firstN(kids, 2);
      const opText = directText(childrenOf(base));
      if (opText.includes('∑')) {
        return [new MathSum({ children: [], subScript: walkNode(under), superScript: [] })];
      }
      if (opText.includes('∫')) {
        return [new MathIntegral({ children: [], subScript: walkNode(under), superScript: [] })];
      }
      // Fallback: render as text
      return [...walkNode(base), ...walkNode(under)];
    }
    
    case 'mover': {
      const [base, over] = firstN(kids, 2);
      const opText = directText(childrenOf(base));
      if (opText.includes('∑')) {
        return [new MathSum({ children: [], subScript: [], superScript: walkNode(over) })];
      }
      if (opText.includes('∫')) {
        return [new MathIntegral({ children: [], subScript: [], superScript: walkNode(over) })];
      }
      return [...walkNode(base), ...walkNode(over)];
    }
    
    case 'munderover': {
      const [base, under, over] = firstN(kids, 3);
      const opText = directText(childrenOf(base));
      if (opText.includes('∑')) {
        if (LO_COMPAT) {
          return naryAsSubSup('∑', walkNode(under), walkNode(over), []);
        }
        return [new MathSum({ children: [], subScript: walkNode(under), superScript: walkNode(over) })];
      }
      if (opText.includes('∫')) {
        if (LO_COMPAT) {
          return naryAsSubSup('∫', walkNode(under), walkNode(over), []);
        }
        return [new MathIntegral({ children: [], subScript: walkNode(under), superScript: walkNode(over) })];
      }
      if (opText.includes('∏')) {
        if (LO_COMPAT) {
          return naryAsSubSup('∏', walkNode(under), walkNode(over), []);
        }
        return [new MathRun('∏'), ...walkNode(under), ...walkNode(over)];
      }
      return [...walkNode(base), ...walkNode(under), ...walkNode(over)];
    }
    
    case 'mtable': {
      const rows = kids.filter((k) => tagName(k) === 'mtr');
      if (LO_COMPAT) {
        // LibreOffice compatibility: use bracket notation
        const parts = [];
        parts.push(new MathRun('['));
        rows.forEach((row, ri) => {
          if (ri > 0) parts.push(new MathRun('; '));
          const cells = childrenOf(row).filter((c) => tagName(c) === 'mtd');
          cells.forEach((cell, ci) => {
            if (ci > 0) parts.push(new MathRun(', '));
            parts.push(...walkChildren(childrenOf(cell)));
          });
        });
        parts.push(new MathRun(']'));
        return parts;
      }
      const rowsCells = rows.map((row) => {
        const cells = childrenOf(row).filter((c) => tagName(c) === 'mtd');
        return cells.map((cell) => walkChildren(childrenOf(cell)));
      });
      return [new MathMatrix(rowsCells)];
    }
    
    default:
      return walkChildren(kids);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Helper functions
// ─────────────────────────────────────────────────────────────────────────────

function tagName(node) {
  const keys = Object.keys(node).filter((k) => k !== 'text' && k !== ':@');
  return keys[0] || null;
}

function childrenOf(node) {
  const tag = tagName(node);
  if (!tag) return [];
  const val = node[tag];
  return Array.isArray(val) ? val : (val ? [val] : []);
}

function textFrom(nodes) {
  const texts = nodes.map((n) => (n.text ?? '').toString()).join('');
  return texts ? [new MathRun(texts)] : [];
}

function directText(nodes) {
  return nodes.map((n) => (n.text ?? '').toString()).join('');
}

function naryAsSubSup(op, lower, upper, body) {
  return [new MathSubSuperScript({ children: [new MathRun(op)], subScript: lower, superScript: upper }), ...body];
}

function findFirst(nodes, name) {
  for (const n of nodes) {
    if (tagName(n) === name) return n;
    const inner = findFirst(childrenOf(n), name);
    if (inner) return inner;
  }
  return null;
}

function firstN(nodes, n) {
  return nodes.slice(0, n);
}

module.exports = { mathmlToDocxChildren };
