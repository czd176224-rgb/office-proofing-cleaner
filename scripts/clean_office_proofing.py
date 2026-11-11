"""One non-overwriting entry point for the Word and PowerPoint proofing workflows."""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import tempfile
import xml.etree.ElementTree as ET
import zipfile

from repair_ppt import check_input, inspect_archive, repair, sha256

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
TRUE = {None, "true", "1", "on"}


def validate_word(path):
    path = Path(path).resolve(strict=True)
    if path.suffix.lower() != ".docx" or path.name.startswith("~$"):
        raise ValueError("Expected an ordinary .docx file, not a lock file")
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if len(names) != len(set(names)) or z.testzip() is not None:
            raise ValueError("Invalid or duplicate ZIP parts")
        if any(n.startswith("_xmlsignatures/") or n.lower().endswith("vbaproject.bin") for n in names):
            raise ValueError("Signed or macro-containing files require separate handling")
        if "word/document.xml" not in names:
            raise ValueError("Not a Word document")
        for n in names:
            if n.endswith(".xml") and b"<!DOCTYPE" in z.read(n):
                raise ValueError("DTD-bearing document parts are not supported")
    return path


def word_state(path):
    result = {"proof_error_markers": 0, "runs": 0, "runs_explicitly_no_proof": 0,
              "hide_spelling_errors": False, "hide_grammar_errors": False}
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if not n.startswith("word/") or not n.endswith(".xml"):
                continue
            root = ET.fromstring(z.read(n))
            result["proof_error_markers"] += len(list(root.iter(f"{{{W}}}proofErr")))
            for run in root.iter(f"{{{W}}}r"):
                result["runs"] += 1
                flag = run.find("w:rPr/w:noProof", NS)
                result["runs_explicitly_no_proof"] += flag is not None and flag.get(f"{{{W}}}val") in TRUE
            if n == "word/settings.xml":
                for element, key in (("hideSpellingErrors", "hide_spelling_errors"),
                                     ("hideGrammaticalErrors", "hide_grammar_errors")):
                    flag = root.find("w:" + element, NS)
                    result[key] = flag is not None and flag.get(f"{{{W}}}val") in TRUE
    return result


def normalized_word(data):
    root = ET.fromstring(data)
    allowed = {f"{{{W}}}{name}" for name in
               ("proofErr", "noProof", "proofState", "hideSpellingErrors", "hideGrammaticalErrors")}
    for parent in root.iter():
        for e in list(parent):
            if e.tag in allowed:
                parent.remove(e)
    for local in ("rPr", "rPrDefault"):
        for parent in root.iter():
            for e in list(parent):
                if e.tag == f"{{{W}}}{local}" and not e.attrib and not list(e):
                    parent.remove(e)
    for e in root.iter():
        if e.tag not in {f"{{{W}}}t", f"{{{W}}}delText"} and e.text is not None and not e.text.strip():
            e.text = None
        if e.tail is not None and not e.tail.strip():
            e.tail = None
        e.attrib = dict(sorted(e.attrib.items()))
    return ET.tostring(root)


def next_output(source):
    target = source.with_name(source.stem + "_校对处理副本" + source.suffix.lower())
    index = 2
    while target.exists() or target.with_suffix(".validation.json").exists():
        target = source.with_name(source.stem + f"_校对处理副本_{index}" + source.suffix.lower())
        index += 1
    return target


def clean(source, output=None, language="auto"):
    source = Path(source).resolve(strict=True)
    extension = source.suffix.lower()
    if extension not in {".docx", ".pptx"}:
        raise ValueError("Supported file types: .docx and .pptx")
    target = Path(output).resolve() if output else next_output(source)
    report_path = target.with_suffix(".validation.json")
    if source == target or target.exists() or report_path.exists():
        raise FileExistsError("Refusing to overwrite source, existing output, or report")
    if target.suffix.lower() != extension:
        raise ValueError("Output extension must match the input")
    if extension == ".pptx":
        return repair(source, target, language=language)
    if language != "auto":
        raise ValueError("--language applies to PPT repair, not the Word suppression mode")

    from clean_word_proofing import clean_docx, visible_text
    source = validate_word(source)
    target.parent.mkdir(parents=True, exist_ok=True)
    original_hash = sha256(source)
    before = word_state(source)
    original_text = visible_text(source)
    changed = []
    with tempfile.TemporaryDirectory(prefix="office-proofing-", dir=target.parent) as tmp:
        staged = Path(tmp) / "candidate.docx"
        stats = clean_docx(source, staged)
        if visible_text(staged) != original_text:
            raise ValueError("Visible Word text changed; no output published")
        with zipfile.ZipFile(source) as a, zipfile.ZipFile(staged) as b:
            names_a = {n for n in a.namelist() if not n.endswith("/")}
            names_b = {n for n in b.namelist() if not n.endswith("/")}
            if names_a != names_b or b.testzip() is not None:
                raise ValueError("Word package members changed")
            for n in sorted(names_a):
                old, new = a.read(n), b.read(n)
                if old == new:
                    continue
                if not n.startswith("word/") or not n.endswith(".xml") or normalized_word(old) != normalized_word(new):
                    raise ValueError(f"Non-proofing content or formatting changed: {n}")
                changed.append(n)
        after = word_state(staged)
        if after["proof_error_markers"] or after["runs"] != after["runs_explicitly_no_proof"]:
            raise ValueError("Not all Word run proofing properties were cleaned")
        if not after["hide_spelling_errors"] or not after["hide_grammar_errors"]:
            raise ValueError("Document-level hiding flags are unavailable or incomplete")
        if sha256(source) != original_hash:
            raise ValueError("Source changed during processing; no output published")
        os.link(staged, target)
    report = {"source": str(source), "output": str(target), "format": "docx",
              "mode": "document_level_proofing_suppression", "source_sha256": original_hash,
              "output_sha256": sha256(target), "source_unchanged": sha256(source) == original_hash,
              "visible_text_preserved": True, "nonproof_formatting_preserved": True,
              "unrelated_parts_byte_identical": True, "changed_parts": changed,
              "application_settings_modified": False, "document_proofing_settings_modified": True,
              "future_proofing_preserved": False, "before": before, "after": after,
              "part_change_counts": stats, "word_application_verified": False}
    with report_path.open("x", encoding="utf-8") as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2)
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path, nargs="?")
    parser.add_argument("--inspect", action="store_true")
    parser.add_argument("--language", default="auto", help="PPT only: auto, preserve, or a language tag")
    args = parser.parse_args()
    if args.inspect:
        if args.output:
            parser.error("--inspect does not accept an output path")
        if args.input.suffix.lower() == ".pptx":
            source = check_input(args.input)
            with zipfile.ZipFile(source) as archive:
                result = dict(source=str(source), format="pptx", **inspect_archive(archive))
        else:
            source = validate_word(args.input)
            result = dict(source=str(source), format="docx", **word_state(source))
    else:
        result = clean(args.input, args.output, args.language)
        result = {key: result[key] for key in ("source", "output", "changed_parts", "source_unchanged")}
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
