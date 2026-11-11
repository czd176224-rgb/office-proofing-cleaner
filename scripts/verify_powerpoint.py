"""Verify a repaired PPTX using disposable PowerPoint copies, never the originals."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import time
import xml.etree.ElementTree as ET
import zipfile

from repair_ppt import A, TRUE, check_input, inspect_archive, parts, sha256

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
TEST_SHAPE = "ppt-proofing-repair-negative-control"
TEST_WORD = "qzxwvvv"


def proofing_settings():
    """Read only relevant registry values; retain digests rather than their contents."""
    import re
    import winreg
    values = {}
    roots = [r"Software\Microsoft\Office\16.0\PowerPoint\Options",
             r"Software\Microsoft\Office\16.0\Common\ProofingTools",
             r"Software\Microsoft\Shared Tools\Proofing Tools"]

    def walk(path):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
                subkeys, count, _ = winreg.QueryInfoKey(key)
                for i in range(count):
                    name, value, kind = winreg.EnumValue(key, i)
                    if re.search(r"spell|gram|proof|error|check|hide", name, re.I):
                        raw = repr((kind, value)).encode("utf-8")
                        values[path + "\\" + name] = hashlib.sha256(raw).hexdigest()
                for i in range(subkeys):
                    walk(path + "\\" + winreg.EnumKey(key, i))
        except FileNotFoundError:
            pass
    for root in roots:
        walk(root)
    return values


def open_state(app):
    return {p.FullName: {"saved": int(p.Saved), "windows": p.Windows.Count}
            for p in app.Presentations}


def snapshot_state(presentation, path):
    presentation.SaveCopyAs(str(path), 24)
    with zipfile.ZipFile(path) as archive:
        return inspect_archive(archive)


def negative_marked(path):
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
        for shape in root.iter(f"{{{P}}}sp"):
            identity = shape.find(f"{{{P}}}nvSpPr/{{{P}}}cNvPr")
            if identity is not None and identity.get("name") == TEST_SHAPE:
                text = "".join(t.text or "" for t in shape.iter(f"{{{A}}}t"))
                return text == TEST_WORD and any(r.get("err") in TRUE for r in shape.iter(f"{{{A}}}rPr"))
    return False


def visible_text(path):
    """Detect edits in test copies, allowing only the verifier's own named textbox."""
    result = {}
    with zipfile.ZipFile(path) as archive:
        for name in parts(archive):
            root = ET.fromstring(archive.read(name))
            for parent in root.iter():
                for shape in list(parent):
                    if shape.tag != f"{{{P}}}sp":
                        continue
                    identity = shape.find(f"{{{P}}}nvSpPr/{{{P}}}cNvPr")
                    if identity is not None and identity.get("name") == TEST_SHAPE:
                        parent.remove(shape)
            text = "".join(t.text or "" for t in root.iter(f"{{{A}}}t"))
            if text:
                result[name] = text
    return result


def verify(source, repaired, directory, spellcheck=False):
    if sys.platform != "win32":
        raise RuntimeError("PowerPoint verification requires Windows")
    import win32com.client
    from PIL import Image, ImageChops

    source, repaired = check_input(source), check_input(repaired)
    if source == repaired:
        raise ValueError("Source and repaired file must be different")
    directory = Path(directory).resolve()
    directory.mkdir(parents=True, exist_ok=False)
    hashes = {str(p): sha256(p) for p in (source, repaired)}
    settings_before = proofing_settings()
    app = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
    original_windows = open_state(app)
    opened = []
    expected_text = {}
    result = {"source": str(source), "repaired": str(repaired), "passed": False,
              "spellcheck_requested": spellcheck, "slide_pixels": [],
              "reopened_proofing_state": None, "new_misspelling_detected": None}

    def open_copy(path, visible=False):
        expected_text[str(path)] = visible_text(path)
        doc = app.Presentations.Open(str(path), False, False, visible)
        opened.append(doc)
        return doc

    def close_copy(doc):
        # Save only our new test copy. Never discard an edit made while it was visible.
        path = Path(doc.FullName)
        if path.parent != directory or str(path) not in expected_text:
            raise RuntimeError("Refusing to close a presentation not created by this verifier")
        doc.Save()
        if visible_text(path) != expected_text[str(path)]:
            result.setdefault("concurrent_edits_preserved_in_open_copies", []).append(str(path))
            opened.remove(doc)
            raise RuntimeError("A test copy was edited concurrently; its contents were saved and the window was left open")
        doc.Close()
        opened.remove(doc)

    def run_spellcheck(doc):
        doc.Windows(1).Activate()
        if app.ActivePresentation.FullName != doc.FullName:
            raise RuntimeError("Test presentation is not active; spelling command not issued")
        app.CommandBars.ExecuteMso("Spelling")
        if app.ActivePresentation.FullName != doc.FullName:
            raise RuntimeError("Active presentation changed during spell checking")

    try:
        snapshots = []
        counts = []
        for label, input_path in (("source", source), ("repaired", repaired)):
            path = directory / (label + "_snapshot.pptx")
            shutil.copyfile(input_path, path)
            snapshots.append(path)
            doc = open_copy(path)
            counts.append(doc.Slides.Count)
            width = 1600
            height = round(width * doc.PageSetup.SlideHeight / doc.PageSetup.SlideWidth)
            for index in range(1, doc.Slides.Count + 1):
                doc.Slides(index).Export(str(directory / f"{label}_{index:03}.png"), "PNG", width, height)
            close_copy(doc)
        result["slides"] = counts[1]
        result["slide_count_matches"] = counts[0] == counts[1]
        if result["slide_count_matches"]:
            for index in range(1, counts[0] + 1):
                with Image.open(directory / f"source_{index:03}.png") as a, Image.open(directory / f"repaired_{index:03}.png") as b:
                    same = a.size == b.size and ImageChops.difference(a.convert("RGB"), b.convert("RGB")).getbbox() is None
                result["slide_pixels"].append({"slide": index, "identical": same})
        result["all_slide_pixels_identical"] = result["slide_count_matches"] and all(x["identical"] for x in result["slide_pixels"])
        if spellcheck:
            doc = open_copy(snapshots[1], visible=True)
            run_spellcheck(doc)
            checked = directory / "reopened_checked.pptx"
            for attempt in range(5):
                state = snapshot_state(doc, checked)
                if state["pending_nodes"] == 0:
                    break
                time.sleep(1)
            result["reopened_proofing_state"] = state
            close_copy(doc)

            # Test a new, obvious misspelling only in a separate disposable copy.
            negative = directory / "negative_input.pptx"
            shutil.copyfile(snapshots[1], negative)
            doc = open_copy(negative, visible=True)
            shape = doc.Slides(1).Shapes.AddTextbox(1, 10, 10, 180, 40)
            shape.Name = TEST_SHAPE
            shape.TextFrame.TextRange.Text = TEST_WORD
            shape.TextFrame.TextRange.LanguageID = 1033
            run_spellcheck(doc)
            checked = directory / "negative_checked.pptx"
            for attempt in range(5):
                snapshot_state(doc, checked)
                result["new_misspelling_detected"] = negative_marked(checked)
                if result["new_misspelling_detected"]:
                    break
                time.sleep(1)
            close_copy(doc)
    except Exception as exc:
        result["verification_error"] = str(exc)
    finally:
        close_errors = []
        for doc in list(opened):
            try:
                close_copy(doc)
            except Exception as exc:
                close_errors.append(str(exc))
        if close_errors:
            result["test_copy_close_errors"] = close_errors
        result["original_files_unchanged"] = all(sha256(Path(p)) == h for p, h in hashes.items())
        try:
            current = open_state(app)
            result["original_open_presentations_preserved"] = all(current.get(p) == state for p, state in original_windows.items())
        except Exception as exc:
            result["original_open_presentations_preserved"] = None
            result["open_window_check_error"] = str(exc)
        result["proofing_registry_values_unchanged"] = proofing_settings() == settings_before
        result["proofing_registry_value_count"] = len(settings_before)
        state = result["reopened_proofing_state"]
        proof_passed = bool(spellcheck and state is not None and state["error_nodes"] == 0
                            and state["pending_nodes"] == 0 and state["no_proof_attributes"] == 0
                            and state["no_proof_languages"] == 0 and result["new_misspelling_detected"])
        result["passed"] = bool(proof_passed and result.get("all_slide_pixels_identical")
                                and result["original_files_unchanged"]
                                and result["original_open_presentations_preserved"]
                                and result["proofing_registry_values_unchanged"]
                                and not result.get("verification_error") and not close_errors
                                and not result.get("concurrent_edits_preserved_in_open_copies"))
        with (directory / "verification.json").open("x", encoding="utf-8") as stream:
            json.dump(result, stream, ensure_ascii=False, indent=2)
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("repaired", type=Path)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--spellcheck", action="store_true")
    args = parser.parse_args()
    result = verify(args.source, args.repaired, args.output_dir, args.spellcheck)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
