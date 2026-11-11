from __future__ import annotations

import argparse
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def qn(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def ensure_child(parent: etree._Element, local: str) -> etree._Element:
    found = parent.find(f"w:{local}", namespaces=NS)
    if found is not None:
        return found
    child = etree.Element(qn(local))
    parent.append(child)
    return child


def patch_settings(root: etree._Element) -> int:
    changed = 0

    proof_state = root.find("w:proofState", namespaces=NS)
    if proof_state is None:
        proof_state = etree.Element(qn("proofState"))
        root.insert(0, proof_state)
        changed += 1
    for attr, value in (("spelling", "clean"), ("grammar", "clean")):
        if proof_state.get(qn(attr)) != value:
            proof_state.set(qn(attr), value)
            changed += 1

    for local in ("hideSpellingErrors", "hideGrammaticalErrors"):
        flag = root.find(f"w:{local}", namespaces=NS)
        if flag is None:
            flag = etree.Element(qn(local))
            root.insert(0, flag)
            changed += 1
        if flag.get(qn("val")) is not None:
            flag.attrib.pop(qn("val"))
            changed += 1

    return changed


def patch_word_xml(root: etree._Element) -> int:
    changed = 0

    for node in list(root.xpath(".//w:proofErr", namespaces=NS)):
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
            changed += 1

    # Cover unformatted runs as well as explicitly formatted runs.
    for run in root.xpath(".//w:r", namespaces=NS):
        if run.find("w:rPr", namespaces=NS) is None:
            run.insert(0, etree.Element(qn("rPr")))
            changed += 1

    for rpr in root.xpath(".//w:rPr", namespaces=NS):
        flag = rpr.find("w:noProof", namespaces=NS)
        if flag is None:
            flag = etree.SubElement(rpr, qn("noProof"))
            changed += 1
        if flag.get(qn("val")) is not None:
            flag.attrib.pop(qn("val"))
            changed += 1

    for doc_defaults in root.xpath(".//w:docDefaults", namespaces=NS):
        rpr_default = ensure_child(ensure_child(doc_defaults, "rPrDefault"), "rPr")
        if rpr_default.find("w:noProof", namespaces=NS) is None:
            rpr_default.append(etree.Element(qn("noProof")))
            changed += 1

    return changed


def word_xml_parts(base: Path) -> set[str]:
    word_dir = base / "word"
    parts = {
        "word/document.xml",
        "word/styles.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/comments.xml",
        "word/settings.xml",
    }
    if word_dir.exists():
        parts.update(f"word/{path.name}" for path in word_dir.glob("header*.xml"))
        parts.update(f"word/{path.name}" for path in word_dir.glob("footer*.xml"))
    return parts


def visible_text(docx_path: Path) -> str:
    text_parts = [
        "word/document.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/comments.xml",
    ]
    out: list[str] = []
    with zipfile.ZipFile(docx_path) as zf:
        names = set(zf.namelist())
        text_parts.extend(name for name in names if name.startswith("word/header") and name.endswith(".xml"))
        text_parts.extend(name for name in names if name.startswith("word/footer") and name.endswith(".xml"))
        for part in sorted(set(text_parts)):
            if part not in names:
                continue
            root = etree.fromstring(zf.read(part))
            out.extend(node.text or "" for node in root.iter(qn("t")))
            out.extend("\t" for _ in root.iter(qn("tab")))
            out.extend("\n" for _ in root.iter(qn("br")))
    return "".join(out)


def clean_docx(input_path: Path, output_path: Path) -> dict[str, int]:
    stats: dict[str, int] = {}

    with tempfile.TemporaryDirectory() as tmp_name:
        tmp_path = Path(tmp_name)
        with zipfile.ZipFile(input_path) as zf:
            zf.extractall(tmp_path)

        parser = etree.XMLParser(remove_blank_text=False, recover=False)
        for part in sorted(word_xml_parts(tmp_path)):
            path = tmp_path / part
            if not path.exists():
                continue
            root = etree.parse(str(path), parser).getroot()
            changed = patch_settings(root) if part == "word/settings.xml" else patch_word_xml(root)
            if changed:
                path.write_bytes(
                    etree.tostring(
                        root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )
                )
            stats[part] = changed

        temp_output = output_path.with_suffix(output_path.suffix + ".tmp")
        if temp_output.exists():
            temp_output.unlink()
        with zipfile.ZipFile(temp_output, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for path in tmp_path.rglob("*"):
                if path.is_file():
                    zf.write(path, path.relative_to(tmp_path).as_posix())
        shutil.move(str(temp_output), str(output_path))

    return stats


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Remove Word proofing underlines from a DOCX without changing visible text."
    )
    parser.add_argument("input", type=Path, help="Source .docx file")
    parser.add_argument("output", type=Path, nargs="?", help="Cleaned output .docx file")
    parser.add_argument("--in-place", action="store_true", help="Overwrite the input .docx")
    parser.add_argument("--verify-text", action="store_true", help="Fail if visible text changes")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = args.input.resolve()
    if not input_path.exists():
        raise SystemExit(f"Input not found: {input_path}")
    if input_path.suffix.lower() != ".docx":
        raise SystemExit("Input must be a .docx file")

    if args.in_place:
        output_path = input_path
    elif args.output:
        output_path = args.output.resolve()
    else:
        output_path = input_path.with_name(f"{input_path.stem}_no-proofing-lines.docx")

    before_text = visible_text(input_path) if args.verify_text else None

    if args.in_place:
        work_path = input_path.with_name(f"{input_path.stem}.proofing-clean.tmp.docx")
        shutil.copy2(input_path, work_path)
        stats = clean_docx(work_path, work_path)
        shutil.move(str(work_path), str(input_path))
    else:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        stats = clean_docx(input_path, output_path)

    if args.verify_text:
        after_text = visible_text(output_path)
        if before_text != after_text:
            raise SystemExit("Visible text changed; output should be inspected before use.")
        print("visible_text_equal: True")

    print(f"output: {output_path}")
    for part, changed in sorted(stats.items()):
        print(f"{part}: {changed}")


if __name__ == "__main__":
    main()
