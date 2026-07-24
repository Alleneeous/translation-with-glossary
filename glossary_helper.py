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
import re
import argparse


def _sanitize_entry(src, tgt):
    """Clean a single glossary row. Returns (src, tgt) or None to skip.

    Cleaning rules:
    - Skip enumerated series (source contains `/` or `……`): these are template
      placeholders (e.g. `原告一/原告二/原告三……`), not real matchable terms.
    - Target with ` / ` (space-slash-space) is an editorial alternative
      annotation; keep the first alternative only.
    - Target ending with `(s)`/`(es)` is a plural-note annotation; strip it
      and let grammar decide singular/plural.
    - Preserve legitimate parentheticals (company suffixes, English
      abbreviations, city names): those don't match the strip patterns above.
    """
    if "/" in src or "……" in src:
        return None

    if " / " in tgt:
        tgt = tgt.split(" / ", 1)[0].strip()

    tgt = re.sub(r"\s*\((?:s|es)\)\s*$", "", tgt).strip()

    if not tgt:
        return None
    return src, tgt


def sanitize_glossary_dict(data):
    """Apply _sanitize_entry rules to an already-loaded glossary dict.

    Used when the glossary comes from a pre-serialized JSON blob (e.g. the
    embedded compressed base64 in app.py) rather than a fresh xlsx load.
    Mutates and returns `data`.
    """
    raw = data.get("glossary", {})
    clean = {}
    cleaned = 0
    skipped = 0
    for src, tgt in raw.items():
        sanitized = _sanitize_entry(src, tgt)
        if sanitized is None:
            skipped += 1
            continue
        clean_src, clean_tgt = sanitized
        if clean_tgt != tgt:
            cleaned += 1
        clean[clean_src] = clean_tgt

    data["glossary"] = clean
    data["sorted_sources"] = sorted(clean.keys(), key=len, reverse=True)
    data["raw_count"] = len(raw)
    data["cleaned_count"] = cleaned
    data["skipped_count"] = skipped
    data["count"] = len(clean)
    return data


def load_glossary(xlsx_path, output_path=None):
    """Load bilingual glossary from .xlsx and save as JSON."""
    import openpyxl

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    glossary = {}
    raw_count = 0
    cleaned_count = 0
    skipped_count = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        src, tgt = row[0], row[1]
        if not (src and tgt):
            continue
        src = str(src).strip()
        tgt = str(tgt).strip()
        if not (src and tgt):
            continue
        raw_count += 1
        original_tgt = tgt
        sanitized = _sanitize_entry(src, tgt)
        if sanitized is None:
            skipped_count += 1
            continue
        clean_src, clean_tgt = sanitized
        if clean_tgt != original_tgt:
            cleaned_count += 1
        glossary[clean_src] = clean_tgt

    sorted_sources = sorted(glossary.keys(), key=len, reverse=True)

    result = {
        "glossary": glossary,
        "sorted_sources": sorted_sources,
        "count": len(glossary),
        "raw_count": raw_count,
        "cleaned_count": cleaned_count,
        "skipped_count": skipped_count,
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

    Two safety passes on top of the raw replace:
    1. If the AI wrote the placeholder without a space next to a Latin word
       (common when the source was CJK-adjacent), insert a space so we don't
       end up with "PlaintiffShengdike Technology".
    2. Collapse adjacent duplicate words like "jurisdiction jurisdiction"
       that arise when a short glossary target (e.g. 管辖 → jurisdiction)
       collides with the AI's own translation of the surrounding word.
    """
    if isinstance(protected_data_or_path, str):
        with open(protected_data_or_path, "r", encoding="utf-8") as f:
            protected_data = json.load(f)
    else:
        protected_data = protected_data_or_path

    mapping = protected_data["mapping"]

    result = translated_text
    for placeholder, translation in mapping.items():
        def _sub(m, _t=translation, _src=result):
            before_idx = m.start() - 1
            after_idx = m.end()
            left = _src[before_idx] if before_idx >= 0 else ""
            right = _src[after_idx] if after_idx < len(_src) else ""
            prefix = " " if (left.isalnum() and _t[:1].isalpha()) else ""
            suffix = " " if (right.isalnum() and _t[-1:].isalpha()) else ""
            return prefix + _t + suffix
        result = re.sub(re.escape(placeholder), _sub, result)

    # Dedupe adjacent identical words that are also glossary targets.
    # These are the cases where a short glossary target (e.g. 管辖 →
    # jurisdiction) collides with the AI's own natural translation of
    # a surrounding word. Restricting to glossary targets avoids touching
    # legit English like "had had" or "that that".
    target_words = {
        w.lower()
        for tgt in mapping.values()
        for w in re.findall(r"\w+", tgt)
        if len(w) >= 3
    }
    if target_words:
        def _dedupe(m):
            return m.group(1) if m.group(1).lower() in target_words else m.group(0)
        result = re.sub(
            r"\b(\w{3,})(?:\s+\1)+\b",
            _dedupe,
            result,
            flags=re.IGNORECASE,
        )

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
