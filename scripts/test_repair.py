"""Run isolated regression checks using generated, disposable PPTX packages."""
from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile

from repair_ppt import A, repair, sha256


def document(content, prefix="a"):
    return (f'<root xmlns:{prefix}="{A}"><{prefix}:p>' + content + f'</{prefix}:p></root>').encode()


def package(path, slide, note=None, signed=False):
    with zipfile.ZipFile(path, "w") as z:
        z.comment = b"keep-package-comment"
        z.writestr("[Content_Types].xml", b"<Types/>")
        z.writestr("ppt/presentation.xml", b"<presentation/>")
        z.writestr("ppt/slides/slide1.xml", slide)
        z.writestr("ppt/media/image1.bin", bytes(range(256)))
        if note:
            z.writestr("ppt/notesSlides/notesSlide1.xml", note)
        if signed:
            z.writestr("_xmlsignatures/sig1.xml", b"<Signature/>")


class RepairTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix="ppt-proofing-unit-")
        self.addCleanup(self.tmp.cleanup)
        self.base = Path(self.tmp.name)
        self.source = self.base / "source.pptx"

    def test_text_quotes_formatting_and_unrelated_parts(self):
        xml = document('<d:r><d:rPr err="true" dirty="0" b="1"><d:solidFill><d:srgbClr val="163C3E"/></d:solidFill></d:rPr><d:t>中文“组织方式”与 A&amp;B</d:t></d:r>', "d")
        package(self.source, xml)
        before = sha256(self.source)
        report = repair(self.source)
        with zipfile.ZipFile(self.source) as a, zipfile.ZipFile(report["output"]) as b:
            self.assertEqual(a.comment, b.comment)
            self.assertEqual(a.read("ppt/media/image1.bin"), b.read("ppt/media/image1.bin"))
            root = ET.fromstring(b.read("ppt/slides/slide1.xml"))
            self.assertEqual("".join(t.text or "" for t in root.iter(f"{{{A}}}t")), '中文“组织方式”与 A&B')
            for run in root.iter(f"{{{A}}}r"):
                rp, text = run.find(f"{{{A}}}rPr"), run.find(f"{{{A}}}t").text
                self.assertEqual(rp.get("b"), "1")
                self.assertEqual(rp.get("lang"), "en-US" if text in {"“", "”"} else "zh-CN")
                self.assertIsNone(rp.get("noProof"))
        self.assertEqual(before, sha256(self.source))

    def test_explicit_languages_and_preserve_mode(self):
        xml = document('<a:r><a:rPr lang="ja-JP" err="1"/><a:t>中文名称</a:t></a:r><a:r><a:rPr lang="en-US" err="1"/><a:t>qzxwvvv</a:t></a:r>')
        package(self.source, xml)
        report = repair(self.source)
        self.assertTrue(report["warnings"])
        with zipfile.ZipFile(report["output"]) as z:
            root = ET.fromstring(z.read("ppt/slides/slide1.xml"))
            self.assertEqual([r.get("lang") for r in root.iter(f"{{{A}}}rPr")], ["ja-JP", "en-US"])
        report = repair(self.source, language="preserve")
        self.assertEqual(report["changed_nodes"], 2)

    def test_no_overwrites_and_numbered_outputs(self):
        package(self.source, document('<a:r><a:rPr err="1"/><a:t>中文</a:t></a:r>'))
        first, second = repair(self.source), repair(self.source)
        self.assertNotEqual(first["output"], second["output"])
        digest = sha256(first["output"])
        with self.assertRaises(FileExistsError):
            repair(self.source, first["output"])
        with self.assertRaises(FileExistsError):
            repair(self.source, self.source)
        self.assertEqual(digest, sha256(first["output"]))

    def test_observed_notes_are_located_without_copying_checked_package(self):
        slide = document('<a:r><a:t>正文</a:t></a:r>')
        note = document('<a:r><a:t>原稿第02页</a:t></a:r>')
        package(self.source, slide, note)
        observed = self.base / "observed.pptx"
        package(observed, slide, document('<a:r><a:rPr err="1"/><a:t>原稿第02页</a:t></a:r>'))
        report = repair(self.source, errors_from=observed)
        self.assertEqual(report["changed_parts"], ["ppt/notesSlides/notesSlide1.xml"])
        self.assertEqual(report["changes"][0]["new_language"], "zh-CN")
        package(observed, slide, document('<a:r><a:rPr err="1"/><a:t>不同内容</a:t></a:r>'))
        with self.assertRaises(ValueError):
            repair(self.source, errors_from=observed)

    def test_scope_all_adds_properties_without_touching_field_content(self):
        package(self.source, document('<a:r><a:t>中文</a:t></a:r><a:fld id="field"><a:rPr err="1"/><a:t>7</a:t></a:fld>'))
        report = repair(self.source, scope="all")
        self.assertEqual(report["changed_nodes"], 2)
        self.assertTrue(report["text_and_nonproof_formatting_preserved"])

    def test_recheck_is_limited_to_the_requested_part(self):
        slide = document('<a:r><a:t>正文</a:t></a:r>')
        package(self.source, slide, document('<a:r><a:t>备注中文</a:t></a:r>'))
        report = repair(self.source, recheck_parts=['ppt/notesSlides/notesSlide1.xml'])
        self.assertEqual(report['changed_parts'], ['ppt/notesSlides/notesSlide1.xml'])
        with zipfile.ZipFile(report['output']) as z:
            self.assertEqual(z.read('ppt/slides/slide1.xml'), slide)
        with self.assertRaises(ValueError):
            repair(self.source, recheck_parts=['ppt/nonexistent.xml'])

    def test_self_closing_properties_next_to_parent_closing_tags(self):
        package(self.source, document('<a:r><a:rPr err="1"/><a:t>中文</a:t></a:r><a:endParaRPr err="1"/>'))
        report = repair(self.source)
        self.assertEqual(report['changed_nodes'], 2)
        self.assertEqual(report['after']['error_nodes'], 0)

    def test_signed_and_dtd_inputs_are_not_published(self):
        package(self.source, document('<a:r><a:t>中文</a:t></a:r>'), signed=True)
        with self.assertRaises(ValueError):
            repair(self.source)
        package(self.source, b'<!DOCTYPE root [<!ENTITY x "text">]><root/>')
        with self.assertRaises(ValueError):
            repair(self.source)
        self.assertEqual(list(self.base.glob('*_校对修复副本*.pptx')), [])


if __name__ == "__main__":
    unittest.main()
