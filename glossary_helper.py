#!/usr/bin/env python3
"""
Glossary Helper for translate-with-glossary skill.

Three subcommands:
  load     - Convert .xlsx glossary to structured JSON
  protect  - Replace glossary source terms in text with placeholders
  restore  - Replace placeholders in translated text with target terms
"""

import sys
import json
import argparse


def load_glossary(xlsx_path, output_path=None):
    """Load bilingual glossary from .xlsx and save as JSON."""
    import openpyxl

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    glossary = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        src, tgt = row[0], row[1]
        if src and tgt:
            src = str(src).strip()
            tgt = str(tgt).strip()
            if src and tgt:
                glossary[src] = tgt

    sorted_sources = sorted(glossary.keys(), key=len, reverse=True)

    result = {
        "glossary": glossary,
        "sorted_sources": sorted_sources,
        "count": len(glossary),
    }

    if output_path:
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)

    return result


def protect_terms(text, glossary_json_or_dict, direction="auto"):
    """Find glossary source terms in text and replace with placeholders.

    Args:
        text: source text string
        glossary_json_or_dict: either a file path (str) to glossary JSON,
                              or an already-loaded dict
        direction: "auto", "cn2en", or "en2cn"

    Returns:
        dict with protected_text, mapping, direction, matched_count, matched_terms
    """
    # Accept either a file path or an already-loaded dict
    if isinstance(glossary_json_or_dict, str):
        with open(glossary_json_or_dict, "r", encoding="utf-8") as f:
            gl_data = json.load(f)
    else:
        gl_data = glossary_json_or_dict

    glossary = gl_data["glossary"]
    reverse_glossary = {v: k for k, v in glossary.items()}

    if direction == "auto":
        cjk_count = sum(
            1 for c in text if "一" <= c <= "鿿" or "㐀" <= c <= "䶿"
        )
        ascii_count = sum(1 for c in text if c.isascii() and c.isalpha())
        direction = "cn2en" if cjk_count > ascii_count else "en2cn"

    if direction == "cn2en":
        source_terms = glossary
    else:
        source_terms = reverse_glossary

    # Pre-filter: only keep terms that actually appear in the document text
    matching = {t: v for t, v in source_terms.items() if t in text}
    sorted_sources = sorted(matching.keys(), key=len, reverse=True)

    mapping = {}
    matched_terms = []
    protected_text = text
    idx = 0

    for term in sorted_sources:
        if term in protected_text:
            placeholder = f"⟨T{idx}⟩"  # ⟨T0⟩
            protected_text = protected_text.replace(term, placeholder)
            mapping[placeholder] = matching[term]
            matched_terms.append(
                {"source": term, "target": matching[term], "placeholder": placeholder}
            )
            idx += 1

    result = {
        "protected_text": protected_text,
        "mapping": mapping,
        "direction": direction,
        "matched_count": len(matched_terms),
        "matched_terms": matched_terms,
    }

    return result


def restore_terms(translated_text, protected_data_or_path):
    """Replace placeholders in translated text with target translations.

    Args:
        translated_text: translated text string with placeholders
        protected_data_or_path: either a file path (str) to the protected JSON,
                               or an already-loaded dict from protect_terms()

    Returns:
        string with placeholders replaced by target translations
    """
    if isinstance(protected_data_or_path, str):
        with open(protected_data_or_path, "r", encoding="utf-8") as f:
            protected_data = json.load(f)
    else:
        protected_data = protected_data_or_path

    mapping = protected_data["mapping"]

    result = translated_text
    for placeholder, translation in mapping.items():
        result = result.replace(placeholder, translation)

    return result


def main():
    parser = argparse.ArgumentParser(description="Glossary Helper for document translation")
    subparsers = parser.add_subparsers(dest="command", required=True)

    p_load = subparsers.add_parser("load", help="Load glossary from xlsx")
    p_load.add_argument("xlsx", help="Path to glossary .xlsx file")
    p_load.add_argument("--output", "-o", default=None, help="Output JSON path")

    p_protect = subparsers.add_parser("protect", help="Protect glossary terms with placeholders")
    p_protect.add_argument("text", help="Path to extracted text file (.txt)")
    p_protect.add_argument("glossary_json", help="Path to glossary JSON file")
    p_protect.add_argument("--output", "-o", default=None, help="Output JSON path")
    p_protect.add_argument(
        "--direction", "-d", default="auto", choices=["auto", "cn2en", "en2cn"]
    )

    p_restore = subparsers.add_parser("restore", help="Restore placeholders with translations")
    p_restore.add_argument("translated_text", help="Path to translated text file")
    p_restore.add_argument("protected_json", help="Path to protected JSON from protect step")
    p_restore.add_argument("--output", "-o", default=None, help="Output text file path")

    args = parser.parse_args()

    if args.command == "load":
        output = args.output or args.xlsx.replace(".xlsx", "_glossary.json")
        result = load_glossary(args.xlsx, output)
        print(json.dumps({"status": "ok", "count": result["count"], "output": output}))

    elif args.command == "protect":
        output = args.output or args.text.replace(".txt", "_protected.json")
        result = protect_terms(args.text, args.glossary_json, args.direction)
        with open(output, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(
            json.dumps(
                {
                    "status": "ok",
                    "direction": result["direction"],
                    "matched_count": result["matched_count"],
                    "output": output,
                },
                ensure_ascii=False,
            )
        )

    elif args.command == "restore":
        output = args.output or args.translated_text.replace(".txt", "_restored.txt")
        restored = restore_terms(args.translated_text, args.protected_json)
        with open(output, "w", encoding="utf-8") as f:
            f.write(restored)
        print(
            json.dumps(
                {"status": "ok", "restored_count": len(json.load(open(args.protected_json))["mapping"]), "output": output},
                ensure_ascii=False,
            )
        )


if __name__ == "__main__":
    main()
