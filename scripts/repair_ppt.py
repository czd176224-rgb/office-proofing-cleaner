"""Repair PPTX proofing metadata without changing application settings (stdlib only)."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import tempfile
import xml.dom.minidom as DOM
import xml.etree.ElementTree as ET
from xml.parsers import expat
import zipfile

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PROPS = {"rPr", "defRPr", "endParaRPr"}
TRUE = {"1", "true"}
QUOTES = set("“”‘’")
HAN = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\U00020000-\U0002fa1f]")


def sha256(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def parts(archive):
    return [n for n in archive.namelist() if n.startswith("ppt/") and n.endswith(".xml")]


def targets(data):
    """Match namespace-aware DOM nodes to byte spans; never serialize a whole part."""
    encoding = re.search(br"<\?xml[^>]*encoding\s*=\s*['\"]([^'\"]+)", data[:200], re.I)
    if data.startswith((b"\xff\xfe", b"\xfe\xff")) or encoding and encoding.group(1).lower() not in {b"utf-8", b"utf8"}:
        raise ValueError("This helper supports UTF-8 OOXML parts only")
    parser = expat.ParserCreate(namespace_separator="|")
    stack, found = [], []

    def start(name, attrs):
        uri, _, local = name.rpartition("|")
        in_run = any(item[0] == A + "|r" for item in stack)
        selected = uri == A and (local == "r" or local in PROPS and not in_run)
        position = parser.CurrentByteIndex
        cursor, quote = position, None
        while cursor < len(data):
            char = data[cursor]
            if quote is not None:
                if char == quote:
                    quote = None
            elif char in (34, 39):
                quote = char
            elif char == 62:
                break
            cursor += 1
        empty_end = cursor + 1 if data[position:cursor].rstrip().endswith(b"/") else None
        stack.append((name, position, selected, empty_end))

    def end(name):
        _, pos, selected, empty_end = stack.pop()
        if selected:
            stop = empty_end if empty_end is not None else data.index(b">", parser.CurrentByteIndex) + 1
            found.append((pos, stop))

    def reject_doctype(*args):
        raise ValueError("OOXML parts containing DTDs are not supported")

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.StartDoctypeDeclHandler = reject_doctype
    parser.Parse(data, True)
    document = DOM.parseString(data)
    nodes = []

    def walk(node, in_run=False):
        if node.nodeType != node.ELEMENT_NODE:
            return
        is_run = node.namespaceURI == A and node.localName == "r"
        if node.namespaceURI == A and (is_run or node.localName in PROPS and not in_run):
            nodes.append(node)
        for child in node.childNodes:
            walk(child, in_run or is_run)

    walk(document.documentElement)
    spans = sorted(found)
    if len(nodes) != len(spans):
        raise ValueError("Could not match DrawingML targets to XML byte spans")
    return [(start, stop, node) for (start, stop), node in zip(spans, nodes)]


def direct(node, local):
    return next((c for c in node.childNodes if c.nodeType == c.ELEMENT_NODE
                 and c.namespaceURI == A and c.localName == local), None)


def text_of(node):
    return "".join(c.data for c in node.childNodes
                   if c.nodeType in (c.TEXT_NODE, c.CDATA_SECTION_NODE))


def choose_language(text, current, mode, warnings):
    if mode == "preserve":
        return None
    if mode != "auto":
        return mode
    han = len(HAN.findall(text))
    if not han:
        if not current and re.search(r"[A-Za-z]", text) and not any(ord(c) > 127 and c.isalpha() for c in text):
            return "en-US"
        return None
    if han < len(re.findall(r"[A-Za-z]", text)):
        warnings.add("Mixed text with more Latin letters than Han characters: language retained")
        return None
    if not current or current.lower().startswith("en-"):
        return "zh-CN"
    if not current.lower().startswith("zh-"):
        warnings.add("Explicit non-Chinese language on Chinese text: review before changing")
    return None


def reset(prop):
    prop.setAttribute("dirty", "1")
    prop.setAttribute("err", "0")


def patch_part(data, mode, scope, observed_errors=None):
    replacements, changes, warnings = [], [], set()
    for start, stop, node in targets(data):
        run = node.localName == "r"
        prop = direct(node, "rPr") if run else node
        marked = prop is not None and prop.getAttribute("err") in TRUE
        run_text = direct(node, "t") if run else None
        observed = run_text is not None and text_of(run_text) in (observed_errors or set())
        if scope == "marked" and not (marked or observed):
            continue
        if not run:
            clone = node.cloneNode(True)
            reset(clone)
            replacements.append((start, stop, clone.toxml(encoding="utf-8")))
            changes.append({"kind": node.localName, "old_error": marked})
            continue
        text_node = direct(node, "t")
        if text_node is None:
            warnings.add("A DrawingML run without a direct text element was skipped")
            continue
        if any(c.nodeType == c.ELEMENT_NODE and not (c.namespaceURI == A and c.localName in {"rPr", "t"})
               for c in node.childNodes):
            warnings.add("Nonstandard DrawingML run skipped")
            continue
        value = text_of(text_node)
        old_lang = prop.getAttribute("lang") if prop is not None else ""
        lang = choose_language(value, old_lang, mode, warnings)
        split = bool(lang and lang.lower().startswith("zh-") and lang != old_lang
                     and (not old_lang or old_lang.lower().startswith("en-")))
        chunks = re.split("([“”‘’])", value) if split else [value]
        output = []
        for chunk in (c for c in chunks if c or len(chunks) == 1):
            clone = node.cloneNode(True)
            cp = direct(clone, "rPr")
            if cp is None:
                qname = (clone.prefix + ":" if clone.prefix else "") + "rPr"
                cp = clone.ownerDocument.createElementNS(A, qname)
                clone.insertBefore(cp, clone.firstChild)
            reset(cp)
            target_lang = (old_lang or "en-US") if split and chunk in QUOTES else lang
            if target_lang:
                cp.setAttribute("lang", target_lang)
                if target_lang.lower().startswith("zh-") and not cp.hasAttribute("altLang"):
                    cp.setAttribute("altLang", "en-US")
            ct = direct(clone, "t")
            while ct.firstChild:
                ct.removeChild(ct.firstChild)
            ct.appendChild(clone.ownerDocument.createTextNode(chunk))
            output.append(clone.toxml(encoding="utf-8"))
        replacements.append((start, stop, b"".join(output)))
        changes.append({"kind": "r", "old_error": marked, "old_language": old_lang or None,
                        "new_language": lang or old_lang or None, "quote_runs_preserved": split,
                        "error_observed_in_powerpoint": observed, "excerpt": value[:100]})
    for start, stop, new in reversed(replacements):
        data = data[:start] + new + data[stop:]
    return data, changes, sorted(warnings)


def normalized(data):
    """Compare non-proofing structure after merging equivalent adjacent text runs."""
    root = ET.fromstring(data)
    for e in root.iter():
        if e.tag in {f"{{{A}}}{n}" for n in PROPS}:
            for key in ("lang", "altLang", "dirty", "err"):
                e.attrib.pop(key, None)
        if e.tag != f"{{{A}}}t" and e.text is not None and not e.text.strip():
            e.text = None
        if e.tail is not None and not e.tail.strip():
            e.tail = None
        e.attrib = dict(sorted(e.attrib.items()))
    for parent in root.iter():
        for e in list(parent):
            if e.tag in {f"{{{A}}}{n}" for n in PROPS} and not e.attrib and not list(e):
                parent.remove(e)
    for para in root.iter(f"{{{A}}}p"):
        prev = None
        for run in list(para):
            if run.tag != f"{{{A}}}r":
                prev = None
                continue
            prop = run.find(f"{{{A}}}rPr")
            signature = ET.tostring(prop) if prop is not None else b""
            text = run.find(f"{{{A}}}t")
            if text is None:
                prev = None
                continue
            signature += repr(sorted(text.attrib.items())).encode()
            if prev is not None and prev[0] == signature:
                prev[1].text = (prev[1].text or "") + (text.text or "")
                para.remove(run)
            else:
                prev = (signature, text)
    return ET.tostring(root)


def inspect_archive(archive):
    result = {"slides": len([n for n in archive.namelist()
                             if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)]),
              "error_nodes": 0, "pending_nodes": 0, "no_proof_attributes": 0,
              "no_proof_languages": 0, "marked_runs": []}
    for name in parts(archive):
        for _, _, node in targets(archive.read(name)):
            prop = direct(node, "rPr") if node.localName == "r" else node
            if prop is None:
                continue
            error = prop.getAttribute("err") in TRUE
            result["error_nodes"] += error
            result["pending_nodes"] += prop.getAttribute("dirty") in TRUE
            result["no_proof_attributes"] += prop.getAttribute("noProof") in TRUE
            result["no_proof_languages"] += prop.getAttribute("lang").lower() in {"zxx", "1024"}
            if error:
                t = direct(node, "t") if node.localName == "r" else None
                result["marked_runs"].append({"part": name, "language": prop.getAttribute("lang") or None,
                                              "excerpt": text_of(t)[:100] if t is not None else None})
    return result


def check_input(path):
    path = Path(path).resolve(strict=True)
    if path.suffix.lower() != ".pptx" or path.name.startswith("~$"):
        raise ValueError("Only ordinary .pptx files are supported; skip lock files")
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)):
            raise ValueError("Duplicate ZIP part names are not supported")
        if any(n.startswith("_xmlsignatures/") for n in names):
            raise ValueError("Digitally signed presentations require separate handling")
        if any(n.lower().endswith("vbaproject.bin") for n in names):
            raise ValueError("Macro-containing presentations require separate handling")
        if "ppt/presentation.xml" not in names or archive.testzip() is not None:
            raise ValueError("Invalid PPTX archive")
    return path


def default_output(source):
    output = source.with_name(source.stem + "_校对修复副本.pptx")
    i = 2
    while output.exists() or output.with_suffix(".validation.json").exists():
        output = source.with_name(source.stem + f"_校对修复副本_{i}.pptx")
        i += 1
    return output


def repair(source, output=None, language="auto", scope="marked", errors_from=None, recheck_parts=None):
    source = check_input(source)
    output = Path(output).resolve() if output else default_output(source)
    report_path = output.with_suffix(".validation.json")
    if source == output or output.exists() or report_path.exists():
        raise FileExistsError("Refusing to overwrite source, output, or existing validation report")
    if output.suffix.lower() != ".pptx":
        raise ValueError("Output must have .pptx extension")
    if language.lower() in {"zxx", "1024"} or language not in {"auto", "preserve"} and not re.fullmatch(r"[A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,8})+", language):
        raise ValueError("Use auto, preserve, or a proofing language such as zh-CN or en-US")
    if scope not in {"marked", "all"}:
        raise ValueError("scope must be marked or all")
    output.parent.mkdir(parents=True, exist_ok=True)
    original_hash = sha256(source)
    recheck_parts = set(recheck_parts or [])
    with zipfile.ZipFile(source) as archive:
        if not recheck_parts.issubset(set(parts(archive))):
            raise ValueError("Every --recheck-part must name an existing ppt/*.xml part")
    observed = {}
    if errors_from is not None:
        observed_path = check_input(errors_from)
        with zipfile.ZipFile(observed_path) as checked, zipfile.ZipFile(source) as original:
            for name in parts(checked):
                errors = set()
                for _, _, node in targets(checked.read(name)):
                    prop = direct(node, "rPr") if node.localName == "r" else node
                    text = direct(node, "t") if node.localName == "r" else None
                    if prop is not None and text is not None and prop.getAttribute("err") in TRUE:
                        errors.add(text_of(text))
                if not errors:
                    continue
                if name not in original.namelist():
                    raise ValueError(f"Observed errors belong to a different presentation: {name}")
                def visible(data):
                    return "".join(e.text or "" for e in ET.fromstring(data).iter(f"{{{A}}}t"))
                if visible(original.read(name)) != visible(checked.read(name)):
                    raise ValueError(f"Observed snapshot has different text: {name}")
                available = {text_of(t) for _, _, node in targets(original.read(name))
                             if node.localName == "r" and (t := direct(node, "t")) is not None}
                if not errors.issubset(available):
                    raise ValueError(f"Observed text run boundaries differ; review manually: {name}")
                observed[name] = errors
    changed, details, warnings = [], [], set()
    with tempfile.TemporaryDirectory(prefix="ppt-proofing-", dir=output.parent) as tmp:
        staged = Path(tmp) / "candidate.pptx"
        with zipfile.ZipFile(source) as archive, zipfile.ZipFile(staged, "w") as destination:
            before = inspect_archive(archive)
            text_parts = set(parts(archive))
            destination.comment = archive.comment
            for item in archive.infolist():
                old = archive.read(item.filename)
                new = old
                if item.filename in text_parts:
                    effective_scope = "all" if item.filename in recheck_parts else scope
                    new, edits, notes = patch_part(old, language, effective_scope, observed.get(item.filename))
                    warnings.update(notes)
                    if new != old:
                        if normalized(old) != normalized(new):
                            raise ValueError(f"Text or non-proofing formatting changed: {item.filename}")
                        changed.append(item.filename)
                        details.extend(dict(part=item.filename, **edit) for edit in edits)
                destination.writestr(item, new)
        with zipfile.ZipFile(source) as original, zipfile.ZipFile(staged) as candidate:
            if original.namelist() != candidate.namelist() or candidate.testzip() is not None:
                raise ValueError("Archive validation failed")
            for name in original.namelist():
                if name not in changed and original.read(name) != candidate.read(name):
                    raise ValueError(f"An unrelated part changed: {name}")
            after = inspect_archive(candidate)
        if sha256(source) != original_hash:
            raise ValueError("Source changed during processing; no result published")
        # Publish a completely validated file atomically, without replacing anything.
        os.link(staged, output)
    report = {"source": str(source), "output": str(output), "source_sha256": original_hash,
              "output_sha256": sha256(output), "source_unchanged": sha256(source) == original_hash,
              "language_mode": language, "scope": scope, "changed_parts": changed,
              "observed_errors_from": str(Path(errors_from).resolve()) if errors_from else None,
              "rechecked_parts": sorted(recheck_parts),
              "changed_nodes": len(details), "text_and_nonproof_formatting_preserved": True,
              "unrelated_parts_byte_identical": True, "before": before, "after": after,
              "changes": details, "warnings": sorted(warnings), "powerpoint_verified": False,
              "status": "file_metadata_repaired_powerpoint_verification_required"}
    with report_path.open("x", encoding="utf-8") as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2)
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    inspect = sub.add_parser("inspect")
    inspect.add_argument("input", type=Path)
    fix = sub.add_parser("repair")
    fix.add_argument("input", type=Path)
    fix.add_argument("output", type=Path, nargs="?")
    fix.add_argument("--language", default="auto")
    fix.add_argument("--scope", choices=["marked", "all"], default="marked")
    fix.add_argument("--errors-from", type=Path, help="A PowerPoint-checked snapshot with identical text")
    fix.add_argument("--recheck-part", action="append", default=[], help="Recheck all text in this XML part only; repeat as needed")
    args = parser.parse_args()
    if args.command == "inspect":
        path = check_input(args.input)
        with zipfile.ZipFile(path) as archive:
            result = dict(source=str(path), **inspect_archive(archive))
    else:
        result = repair(args.input, args.output, args.language, args.scope, args.errors_from, args.recheck_part)
        result = {key: result[key] for key in ("source", "output", "changed_parts", "changed_nodes", "warnings", "status")}
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
