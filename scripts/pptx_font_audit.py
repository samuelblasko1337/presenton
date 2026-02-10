import argparse
import json
import os
import zipfile
import re
from collections import Counter
from pathlib import Path
import xml.etree.ElementTree as ET

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _read_zip_entry(zf: zipfile.ZipFile, name: str) -> bytes | None:
    try:
        with zf.open(name) as f:
            return f.read()
    except KeyError:
        return None


def _parse_embedded_typefaces(presentation_xml: bytes | None) -> list[str]:
    if not presentation_xml:
        return []
    root = ET.fromstring(presentation_xml)
    typefaces: list[str] = []
    for font in root.findall(".//p:embeddedFontLst/p:embeddedFont/p:font", NS):
        tf = font.get("typeface")
        if tf:
            typefaces.append(tf)
    return sorted(set(typefaces))


def _parse_embedded_font_styles(presentation_xml: bytes | None) -> dict[str, dict[str, bool]]:
    if not presentation_xml:
        return {}
    root = ET.fromstring(presentation_xml)
    styles: dict[str, dict[str, bool]] = {}
    for embedded in root.findall(".//p:embeddedFontLst/p:embeddedFont", NS):
        font_el = embedded.find("p:font", NS)
        if font_el is None:
            continue
        tf = font_el.get("typeface")
        if not tf:
            continue
        styles[tf] = {
            "has_regular": embedded.find("p:regular", NS) is not None,
            "has_bold": embedded.find("p:bold", NS) is not None,
            "has_italic": embedded.find("p:italic", NS) is not None,
            "has_boldItalic": embedded.find("p:boldItalic", NS) is not None,
        }
    return styles


def _parse_relationships(rels_xml: bytes | None) -> dict[str, str]:
    if not rels_xml:
        return {}
    root = ET.fromstring(rels_xml)
    rels: dict[str, str] = {}
    for rel in root.findall(f".//{{{PKG_REL_NS}}}Relationship"):
        rid = rel.get("Id")
        target = rel.get("Target")
        if rid and target:
            rels[rid] = target
    return rels


def _resolve_ppt_target(target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    if target.startswith("ppt/"):
        return target
    return f"ppt/{target}"


def _extract_utf16le_strings(data: bytes, min_len: int = 3, max_items: int = 20) -> list[str]:
    strings: list[str] = []
    current: list[int] = []
    data_len = len(data)
    i = 0
    while i + 1 < data_len:
        lo = data[i]
        hi = data[i + 1]
        if hi == 0 and 32 <= lo <= 126:
            current.append(lo)
        else:
            if len(current) >= min_len:
                strings.append(bytes(current).decode("ascii", errors="ignore"))
                if len(strings) >= max_items:
                    break
            current = []
        i += 2
    if len(strings) < max_items and len(current) >= min_len:
        strings.append(bytes(current).decode("ascii", errors="ignore"))
    # de-dupe while preserving order
    seen = set()
    out: list[str] = []
    for s in strings:
        if s not in seen:
            seen.add(s)
            out.append(s)
    return out


def _filter_name_strings(strings: list[str], max_len: int = 40, max_items: int = 10) -> list[str]:
    out: list[str] = []
    for s in strings:
        if len(s) > max_len:
            continue
        if not any(ch.isalpha() for ch in s):
            continue
        out.append(s)
        if len(out) >= max_items:
            break
    return out


def _gather_embedded_font_binaries(
    presentation_xml: bytes | None,
    rels_xml: bytes | None,
    zf: zipfile.ZipFile,
) -> list[dict]:
    if not presentation_xml:
        return []
    rels = _parse_relationships(rels_xml)
    root = ET.fromstring(presentation_xml)
    entries: list[dict] = []
    for embedded in root.findall(".//p:embeddedFontLst/p:embeddedFont", NS):
        font_el = embedded.find("p:font", NS)
        typeface = font_el.get("typeface") if font_el is not None else None
        for style_tag in ("regular", "bold", "italic", "boldItalic"):
            style_el = embedded.find(f"p:{style_tag}", NS)
            if style_el is None:
                continue
            rid = style_el.get(f"{{{R_NS}}}id")
            target = rels.get(rid) if rid else None
            part_name = _resolve_ppt_target(target) if target else None
            data = _read_zip_entry(zf, part_name) if part_name else None
            strings = _extract_utf16le_strings(data) if data else []
            entries.append(
                {
                    "typeface": typeface,
                    "style": style_tag,
                    "rId": rid,
                    "part_name": part_name,
                    "utf16_strings": _filter_name_strings(strings),
                }
            )
    return entries


def _parse_theme_fonts(theme_xml: bytes | None) -> tuple[str | None, str | None]:
    if not theme_xml:
        return None, None
    root = ET.fromstring(theme_xml)
    major_el = root.find(".//a:themeElements/a:fontScheme/a:majorFont/a:latin", NS)
    minor_el = root.find(".//a:themeElements/a:fontScheme/a:minorFont/a:latin", NS)
    major = major_el.get("typeface") if major_el is not None else None
    minor = minor_el.get("typeface") if minor_el is not None else None
    return major, minor


def _iter_slide_entries(zf: zipfile.ZipFile) -> list[str]:
    slide_entries = [
        name
        for name in zf.namelist()
        if name.startswith("ppt/slides/slide") and name.endswith(".xml")
    ]
    return sorted(slide_entries)


def _latin_has_typeface(latin_el: ET.Element | None) -> bool:
    if latin_el is None:
        return False
    return bool(latin_el.get("typeface"))


def _paragraph_has_def_latin(p_el: ET.Element) -> bool:
    def_rpr = p_el.find("a:pPr/a:defRPr", NS)
    latin = def_rpr.find("a:latin", NS) if def_rpr is not None else None
    return _latin_has_typeface(latin)


def _paragraph_has_end_latin(p_el: ET.Element) -> bool:
    end_rpr = p_el.find("a:endParaRPr", NS)
    latin = end_rpr.find("a:latin", NS) if end_rpr is not None else None
    return _latin_has_typeface(latin)


def _paragraph_has_typeface(p_el: ET.Element) -> bool:
    return _paragraph_has_def_latin(p_el) or _paragraph_has_end_latin(p_el)


def _run_has_typeface(r_el: ET.Element) -> bool:
    rpr = r_el.find("a:rPr", NS)
    if rpr is None:
        return False
    latin = rpr.find("a:latin", NS)
    return _latin_has_typeface(latin)


def _extract_slide_index(slide_name: str) -> int | None:
    match = re.search(r"slide(\d+)\.xml$", slide_name)
    if not match:
        return None
    return int(match.group(1))


def _counter_to_sorted_dict(counter: Counter[str]) -> dict[str, int]:
    return {
        k: counter[k]
        for k in sorted(counter.keys(), key=lambda x: (-counter[x], x))
    }


def _top_items(counter: Counter[str], limit: int = 10) -> str:
    if not counter:
        return ""
    items = sorted(counter.items(), key=lambda x: (-x[1], x[0]))[:limit]
    return ", ".join(f"{k}:{v}" for k, v in items)


def _group_styles(counter: Counter[tuple[str, str]]) -> dict[str, set[str]]:
    grouped: dict[str, set[str]] = {}
    for (tf, style) in counter.keys():
        grouped.setdefault(tf, set()).add(style)
    return grouped


def _truthy_attr(value: str | None) -> bool:
    if value is None:
        return False
    return value.strip().lower() in {"1", "true", "t", "on", "yes"}


def _required_style_name(is_bold: bool, is_italic: bool) -> str:
    if is_bold and is_italic:
        return "boldItalic"
    if is_bold:
        return "bold"
    if is_italic:
        return "italic"
    return "regular"


def audit_pptx(pptx_path: Path) -> dict:
    if not pptx_path.exists():
        raise FileNotFoundError(str(pptx_path))

    with zipfile.ZipFile(pptx_path, "r") as zf:
        presentation_xml = _read_zip_entry(zf, "ppt/presentation.xml")
        presentation_rels_xml = _read_zip_entry(zf, "ppt/_rels/presentation.xml.rels")
        embedded_typefaces = _parse_embedded_typefaces(presentation_xml)
        embedded_styles = _parse_embedded_font_styles(presentation_xml)
        embedded_font_binaries = _gather_embedded_font_binaries(
            presentation_xml, presentation_rels_xml, zf
        )
        theme_xml = _read_zip_entry(zf, "ppt/theme/theme1.xml")
        theme_major, theme_minor = _parse_theme_fonts(theme_xml)

        requested_counts_raw: Counter[str] = Counter()
        requested_counts_faces: Counter[str] = Counter()
        requested_counts_tokens: Counter[str] = Counter()
        requested_counts_resolved: Counter[str] = Counter()
        missing_runs = 0
        missing_paragraphs = 0
        missing_paragraphs_empty = 0
        missing_paragraphs_nonempty = 0
        missing_runs_with_text = 0
        missing_paragraph_locations: list[dict] = []
        unsupported_style_counts: Counter[tuple[str, str]] = Counter()
        unsupported_style_violations: list[dict] = []

        for slide_name in _iter_slide_entries(zf):
            slide_xml = _read_zip_entry(zf, slide_name)
            if not slide_xml:
                continue
            root = ET.fromstring(slide_xml)
            slide_index = _extract_slide_index(slide_name)

            for latin in root.findall(".//a:latin", NS):
                tf = latin.get("typeface")
                if tf:
                    requested_counts_raw[tf] += 1
                    if tf.startswith("+"):
                        requested_counts_tokens[tf] += 1
                        if tf == "+mn-lt" and theme_minor:
                            requested_counts_resolved[theme_minor] += 1
                        elif tf == "+mj-lt" and theme_major:
                            requested_counts_resolved[theme_major] += 1
                    else:
                        requested_counts_faces[tf] += 1
                        requested_counts_resolved[tf] += 1

            shapes_with_tx = root.findall(".//p:sp[p:txBody]", NS)
            for shape_idx, shape in enumerate(shapes_with_tx, start=1):
                nv = shape.find("p:nvSpPr/p:cNvPr", NS)
                shape_id = nv.get("id") if nv is not None else None
                shape_name = nv.get("name") if nv is not None else None
                tx_body = shape.find("p:txBody", NS)
                if tx_body is None:
                    continue
                paragraphs = tx_body.findall("a:p", NS)
                for p_idx, p in enumerate(paragraphs, start=1):
                    has_def = _paragraph_has_def_latin(p)
                    has_end = _paragraph_has_end_latin(p)
                    p_has = has_def or has_end
                    runs = p.findall(".//a:r", NS)
                    run_typefaces = set()
                    run_texts = []
                    for r_idx, r in enumerate(runs, start=1):
                        rpr = r.find("a:rPr", NS)
                        latin = rpr.find("a:latin", NS) if rpr is not None else None
                        tf = latin.get("typeface") if latin is not None else None
                        if tf:
                            run_typefaces.add(tf)
                        t_el = r.find("a:t", NS)
                        if t_el is not None and t_el.text:
                            run_texts.append(t_el.text)
                        if rpr is not None and tf:
                            style_flags = embedded_styles.get(tf)
                            if style_flags:
                                is_bold = _truthy_attr(rpr.get("b"))
                                is_italic = _truthy_attr(rpr.get("i"))
                                required_style = _required_style_name(is_bold, is_italic)
                                has_required = style_flags.get(f"has_{required_style}", False)
                                if not has_required:
                                    unsupported_style_counts[(tf, required_style)] += 1
                                    snippet = t_el.text[:40] if t_el is not None and t_el.text else ""
                                    unsupported_style_violations.append(
                                        {
                                            "slide_file": slide_name,
                                            "shape_id": shape_id,
                                            "shape_name": shape_name,
                                            "paragraph_index": p_idx,
                                            "run_index": r_idx,
                                            "typeface": tf,
                                            "bold": is_bold,
                                            "italic": is_italic,
                                            "snippet": snippet,
                                        }
                                    )

                    if not p_has:
                        missing_paragraphs += 1
                        run_count = len(runs)
                        snippet_text = "".join(run_texts)
                        is_empty = run_count == 0 and snippet_text == ""
                        if is_empty:
                            missing_paragraphs_empty += 1
                        else:
                            missing_paragraphs_nonempty += 1
                            missing_paragraph_locations.append(
                                {
                                    "slide_file": slide_name,
                                    "slide_index": slide_index,
                                    "shape_index": shape_idx,
                                    "shape_id": shape_id,
                                    "shape_name": shape_name,
                                    "paragraph_index": p_idx,
                                    "has_runs": run_count > 0,
                                    "run_count": run_count,
                                    "run_typefaces": sorted(run_typefaces),
                                    "snippet": snippet_text[:40],
                                    "has_defRPr_latin": has_def,
                                    "has_endParaRPr_latin": has_end,
                                    "xpaths": f"/p:sld/p:cSld/p:spTree/p:sp[{shape_idx}]/p:txBody/a:p[{p_idx}]",
                                }
                            )

                    for r in runs:
                        r_has = _run_has_typeface(r)
                        if not (r_has or p_has):
                            missing_runs += 1
                            t_el = r.find("a:t", NS)
                            if t_el is not None and t_el.text:
                                missing_runs_with_text += 1

        requested_typefaces = _counter_to_sorted_dict(requested_counts_raw)
        requested_faces = _counter_to_sorted_dict(requested_counts_faces)
        requested_theme_tokens = _counter_to_sorted_dict(requested_counts_tokens)
        requested_resolved = _counter_to_sorted_dict(requested_counts_resolved)

        embedded_set = set(embedded_typefaces)
        requested_set = set(requested_counts_faces.keys())
        unknown_requested = sorted(requested_set - embedded_set)

        report = {
            "pptx_path": str(pptx_path),
            "embedded_typefaces": embedded_typefaces,
            "embedded_font_styles": embedded_styles,
            "embedded_font_binaries": embedded_font_binaries,
            "theme_majorLatin": theme_major,
            "theme_minorLatin": theme_minor,
            "requested_typefaces": requested_typefaces,
            "requested_faces": requested_faces,
            "requested_theme_tokens": requested_theme_tokens,
            "requested_resolved": requested_resolved,
            "missing_typeface_runs": {
                "runs": missing_runs,
                "paragraphs": missing_paragraphs,
                "total": missing_runs + missing_paragraphs,
            },
            "missing_paragraphs_empty": missing_paragraphs_empty,
            "missing_paragraphs_nonempty": missing_paragraphs_nonempty,
            "missing_runs_with_text": missing_runs_with_text,
            "missing_paragraph_locations": missing_paragraph_locations,
            "unsupported_style_usage": {
                "counts": {
                    tf: {style: unsupported_style_counts[(tf, style)] for style in sorted(styles)}
                    for tf, styles in _group_styles(unsupported_style_counts).items()
                },
                "violations": unsupported_style_violations,
            },
            "unknown_requested": unknown_requested,
            "counts": {
                "embedded": len(embedded_typefaces),
                "requested_unique": len(requested_counts_raw),
                "requested_faces_unique": len(requested_counts_faces),
                "requested_theme_tokens_unique": len(requested_counts_tokens),
                "requested_resolved_unique": len(requested_counts_resolved),
                "unknown_requested": len(unknown_requested),
            },
        }

    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit PPTX embedded vs requested fonts")
    parser.add_argument("--pptx", required=True, help="Path to PPTX")
    parser.add_argument("--out", help="Output JSON report path")
    args = parser.parse_args()

    pptx_path = Path(args.pptx)
    report = audit_pptx(pptx_path)

    out_path = Path(args.out) if args.out else None
    if out_path:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(report, indent=2), encoding="utf-8")

    print("PPTX Font Audit")
    print(f"  PPTX: {pptx_path}")
    print(f"  Theme major/minor: {report['theme_majorLatin']} / {report['theme_minorLatin']}")
    print(f"  Embedded typefaces ({len(report['embedded_typefaces'])}): {', '.join(report['embedded_typefaces'])}")
    if report.get("embedded_font_binaries"):
        print("  Embedded font binaries (sample):")
        for entry in report["embedded_font_binaries"][:10]:
            names = ", ".join(entry.get("utf16_strings", []))
            print(
                f"    - {entry.get('typeface')} {entry.get('style')} "
                f"{entry.get('part_name')} :: {names}"
            )
    print(f"  Requested faces ({len(report['requested_faces'])}): {', '.join(report['requested_faces'].keys())}")
    print(
        f"  Requested theme tokens ({len(report['requested_theme_tokens'])}): "
        f"{', '.join(report['requested_theme_tokens'].keys())}"
    )
    print(f"  Requested raw ({len(report['requested_typefaces'])}): {', '.join(report['requested_typefaces'].keys())}")
    top_requested = _top_items(Counter(report["requested_typefaces"]))
    top_resolved = _top_items(Counter(report["requested_resolved"]))
    print(f"  Top requested (raw): {top_requested if top_requested else '(none)'}")
    print(f"  Top resolved: {top_resolved if top_resolved else '(none)'}")
    missing = report["missing_typeface_runs"]
    print(f"  Missing typeface runs: runs={missing['runs']} paragraphs={missing['paragraphs']} total={missing['total']}")
    print(f"  Unknown requested ({len(report['unknown_requested'])}): {', '.join(report['unknown_requested'])}")
    print(
        f"  missing paragraphs: {missing['paragraphs']} (empty={report['missing_paragraphs_empty']} nonempty={report['missing_paragraphs_nonempty']})"
    )
    print(f"  missing runs with text: {report['missing_runs_with_text']}")
    for loc in report["missing_paragraph_locations"][:10]:
        snippet = loc.get("snippet", "")
        print(
            f"    - {loc.get('slide_file')} "
            f"s{loc.get('shape_index')} p{loc.get('paragraph_index')} "
            f"{snippet}"
        )
    unsupported = report["unsupported_style_usage"]["counts"]
    if unsupported:
        print("  Unsupported style usage:")
        for tf in sorted(unsupported.keys()):
            styles = unsupported[tf]
            parts = [f"{style}={styles[style]}" for style in sorted(styles.keys())]
            print(f"    - {tf}: " + ", ".join(parts))
    else:
        print("  Unsupported style usage: none")
    violations = report["unsupported_style_usage"]["violations"]
    if violations:
        print("  Top unsupported style violations:")
        for v in violations[:10]:
            snippet = v.get("snippet", "")
            print(
                f"    - {v.get('slide_file')} "
                f"s{v.get('shape_id')} p{v.get('paragraph_index')} r{v.get('run_index')} "
                f"{v.get('typeface')} b={v.get('bold')} i={v.get('italic')} "
                f"{snippet}"
            )

    if report["unknown_requested"]:
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
