"""Insert a numbered formula block into an unpacked DOCX document.

Creates a 3-column borderless table for each block formula:
  [ 1cm gap | formula (centered, OMML) | (n) right-aligned 1cm ]

LaTeX is converted to OMML via pandoc (must be installed and on PATH).
Requires: pandoc >= 2.0

Usage:
    # Insert after the paragraph that contains <anchor text>:
    python scripts/formula.py unpacked/ "Q_n(x,a) = r + \\gamma \\max" 1 --anchor "由此可得"

    # Output XML snippet to stdout only (no file modification):
    python scripts/formula.py --latex "E=mc^2" --number 2

    # Read LaTeX from stdin:
    echo "E=mc^2" | python scripts/formula.py unpacked/ - 3 --anchor "其中"
"""

import argparse
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

# ── Layout constants (A4, 2.5 cm margins) ────────────────────────────────────
CONTENT_WIDTH = 9070   # DXA  (11906 - 2×1418)
GAP_WIDTH     = 567    # DXA  (1 cm)
NUM_WIDTH     = 567    # DXA  (1 cm)
FORMULA_WIDTH = CONTENT_WIDTH - GAP_WIDTH - NUM_WIDTH   # 7936 DXA

# Word / OOXML namespaces
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


# ── LaTeX → OMML conversion ──────────────────────────────────────────────────

def latex_to_omml(latex: str) -> str:
    """Convert a LaTeX math string to an OMML <m:oMath> XML fragment.

    Uses pandoc: write a minimal .docx containing ``$<latex>$``, unzip it,
    and extract the first <m:oMath> element from document.xml.

    Raises RuntimeError if pandoc is not available or conversion fails.
    """
    # Wrap in display math so pandoc generates a block equation
    markdown_src = f"$$\n{latex}\n$$\n"

    with tempfile.TemporaryDirectory() as tmp:
        md_path   = Path(tmp) / "eq.md"
        docx_path = Path(tmp) / "eq.docx"
        md_path.write_text(markdown_src, encoding="utf-8")

        result = subprocess.run(
            ["pandoc", str(md_path), "-o", str(docx_path)],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"pandoc failed:\n{result.stderr}\n"
                "Make sure pandoc is installed: https://pandoc.org/installing.html"
            )

        # Extract OMML from the generated docx
        with zipfile.ZipFile(docx_path) as zf:
            doc_xml = zf.read("word/document.xml").decode("utf-8")

    omml = _extract_first_omath(doc_xml)
    if not omml:
        raise RuntimeError(
            "Could not extract <m:oMath> from pandoc output.\n"
            "Verify that pandoc supports math conversion on this system."
        )
    return omml


def _extract_first_omath(xml_text: str) -> str | None:
    """Return the first <m:oMath>…</m:oMath> block as a string, or None."""
    # Use a simple regex; the element may contain nested tags.
    pattern = re.compile(
        r"<(?:m:)?oMath\b[^>]*>.*?</(?:m:)?oMath>",
        re.DOTALL
    )
    m = pattern.search(xml_text)
    if not m:
        return None
    fragment = m.group(0)
    # Normalise namespace prefix to m:
    if not fragment.startswith("<m:oMath"):
        fragment = re.sub(r"^<oMath", "<m:oMath", fragment)
        fragment = re.sub(r"</oMath>$", "</m:oMath>", fragment)
    # Ensure the math namespace is declared on the element
    if 'xmlns:m=' not in fragment:
        fragment = fragment.replace(
            "<m:oMath",
            f'<m:oMath xmlns:m="{MATH_NS}"',
            1
        )
    return fragment


# ── Table XML builder ─────────────────────────────────────────────────────────

_NO_BORDER = '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>' \
             '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>' \
             '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>' \
             '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>' \
             '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>' \
             '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'


def build_formula_table(omml: str, number: int | str) -> str:
    """Return the XML string for a 3-column formula table row."""
    num_text = f"({number})"

    return f"""\
<w:tbl
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:m="{MATH_NS}">
  <w:tblPr>
    <w:tblStyle w:val="TableNormal"/>
    <w:tblW w:w="{CONTENT_WIDTH}" w:type="dxa"/>
    <w:tblBorders>{_NO_BORDER}</w:tblBorders>
    <w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/>
      <w:left w:w="0" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/>
      <w:right w:w="0" w:type="dxa"/>
    </w:tblCellMar>
    <w:tblLook w:val="0000"/>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="{GAP_WIDTH}"/>
    <w:gridCol w:w="{FORMULA_WIDTH}"/>
    <w:gridCol w:w="{NUM_WIDTH}"/>
  </w:tblGrid>
  <w:tr>
    <!-- Column 1: 1cm gap -->
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="{GAP_WIDTH}" w:type="dxa"/>
        <w:tcBorders>{_NO_BORDER}</w:tcBorders>
      </w:tcPr>
      <w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr></w:p>
    </w:tc>
    <!-- Column 2: formula, centered -->
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="{FORMULA_WIDTH}" w:type="dxa"/>
        <w:tcBorders>{_NO_BORDER}</w:tcBorders>
      </w:tcPr>
      <w:p>
        <w:pPr>
          <w:jc w:val="center"/>
          <w:spacing w:line="240" w:lineRule="auto"/>
          <w:ind w:firstLine="0"/>
        </w:pPr>
        {omml}
      </w:p>
    </w:tc>
    <!-- Column 3: equation number, right-aligned -->
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="{NUM_WIDTH}" w:type="dxa"/>
        <w:tcBorders>{_NO_BORDER}</w:tcBorders>
      </w:tcPr>
      <w:p>
        <w:pPr>
          <w:jc w:val="right"/>
          <w:spacing w:line="240" w:lineRule="auto"/>
          <w:ind w:firstLine="0"/>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:vertAlign w:val="baseline"/>
          </w:rPr>
          <w:t xml:space="preserve">{num_text}</w:t>
        </w:r>
      </w:p>
    </w:tc>
  </w:tr>
</w:tbl>"""


# ── document.xml insertion ────────────────────────────────────────────────────

def insert_after_anchor(doc_xml_path: Path, anchor: str, table_xml: str) -> bool:
    """Insert table_xml after the first <w:p> whose text contains anchor.

    Returns True on success, False if anchor not found.
    """
    content = doc_xml_path.read_text(encoding="utf-8")

    # Find the paragraph that contains the anchor text
    # We look for <w:p ...>...</w:p> spanning across the anchor string.
    para_pattern = re.compile(r"(<w:p\b[^>]*>(?:(?!<w:p\b).)*?</w:p>)", re.DOTALL)
    match_pos = None
    match_end = None
    for m in para_pattern.finditer(content):
        # Strip tags to get plain text
        plain = re.sub(r"<[^>]+>", "", m.group(0))
        if anchor in plain:
            match_pos = m.start()
            match_end = m.end()
            break

    if match_end is None:
        return False

    # Insert the table XML right after the closing </w:p>
    new_content = content[:match_end] + "\n" + table_xml + content[match_end:]
    doc_xml_path.write_text(new_content, encoding="utf-8")
    return True


# ── CLI ───────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Insert a numbered formula block into an unpacked DOCX."
    )
    # Two modes:
    #   (a) snippet-only: --latex + --number
    #   (b) insert mode : unpacked_dir, latex_or_dash, number [--anchor]
    parser.add_argument(
        "unpacked_dir", nargs="?",
        help="Path to the unpacked DOCX directory (omit for snippet-only output)"
    )
    parser.add_argument(
        "latex_positional", nargs="?", metavar="LATEX",
        help="LaTeX formula string, or '-' to read from stdin"
    )
    parser.add_argument(
        "number_positional", nargs="?", type=str, metavar="NUMBER",
        help="Equation number (e.g. 1, 2, A.1)"
    )
    parser.add_argument("--latex", help="LaTeX formula string (alternative to positional)")
    parser.add_argument("--number", dest="number_opt", help="Equation number (alternative to positional)")
    parser.add_argument(
        "--anchor",
        help="Insert the formula table after the paragraph containing this text"
    )
    args = parser.parse_args()

    # Resolve latex and number from positional or named args
    latex  = args.latex  or args.latex_positional
    number = args.number_opt or args.number_positional

    if not latex:
        parser.error("LaTeX formula is required (positional or --latex).")
    if not number:
        parser.error("Equation number is required (positional or --number).")

    # Read from stdin if "-"
    if latex == "-":
        latex = sys.stdin.read().strip()

    # Convert LaTeX → OMML
    print(f"[formula] Converting LaTeX to OMML via pandoc…", file=sys.stderr)
    try:
        omml = latex_to_omml(latex)
    except RuntimeError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    table_xml = build_formula_table(omml, number)

    # Snippet-only mode
    if not args.unpacked_dir:
        print(table_xml)
        return

    unpacked = Path(args.unpacked_dir)
    doc_xml  = unpacked / "word" / "document.xml"
    if not doc_xml.exists():
        print(f"Error: {doc_xml} not found.", file=sys.stderr)
        sys.exit(1)

    if not args.anchor:
        # No anchor → just print snippet and hint
        print(table_xml)
        print(
            "\n[formula] No --anchor given; snippet printed above.\n"
            "          Paste it into word/document.xml at the desired location.",
            file=sys.stderr
        )
        return

    print(f"[formula] Inserting after paragraph containing: \"{args.anchor}\"", file=sys.stderr)
    success = insert_after_anchor(doc_xml, args.anchor, table_xml)
    if success:
        print(f"[formula] ✓ Inserted formula ({number}) into {doc_xml}", file=sys.stderr)
    else:
        print(
            f"[formula] ✗ Anchor text not found: \"{args.anchor}\"\n"
            f"          Snippet printed below — paste manually.",
            file=sys.stderr
        )
        print(table_xml)
        sys.exit(1)


if __name__ == "__main__":
    main()
