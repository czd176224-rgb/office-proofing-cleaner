"""Regression checks for routing, Word content preservation, and output protection."""
from pathlib import Path
import tempfile
import unittest
import zipfile

from clean_office_proofing import clean, word_state
from clean_word_proofing import visible_text
from repair_ppt import sha256
from test_repair import document, package

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def word_package(path, signed=False):
    body = f'''<w:document xmlns:w="{W}"><w:body><w:p>
<w:proofErr w:type="spellStart"/><w:r><w:t>中文测试</w:t></w:r>
<w:r><w:rPr><w:noProof w:val="false"/><w:b/></w:rPr><w:t>qzxwvvv</w:t></w:r>
<w:proofErr w:type="spellEnd"/>
<w:ins w:id="1" w:author="Test"><w:r><w:t>修订保留</w:t></w:r></w:ins>
<w:del w:id="2" w:author="Test"><w:r><w:delText>删除修订保留</w:delText></w:r></w:del>
</w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>表格保留</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
</w:body></w:document>'''
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", b"<Types/>")
        z.writestr("word/document.xml", body.encode())
        z.writestr("word/styles.xml", f'<w:styles xmlns:w="{W}"><w:docDefaults/></w:styles>'.encode())
        z.writestr("word/settings.xml", f'<w:settings xmlns:w="{W}"><w:hideSpellingErrors w:val="false"/><w:hideGrammaticalErrors w:val="0"/></w:settings>'.encode())
        z.writestr("word/header1.xml", f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>页眉保留</w:t></w:r></w:p></w:hdr>'.encode())
        z.writestr("word/comments.xml", f'<w:comments xmlns:w="{W}"><w:comment w:id="0"><w:p><w:r><w:t>批注保留</w:t></w:r></w:p></w:comment></w:comments>'.encode())
        z.writestr("word/media/image1.bin", bytes(range(256)))
        z.writestr("word/_rels/document.xml.rels", b"<Relationships/>")
        if signed:
            z.writestr("_xmlsignatures/sig1.xml", b"<Signature/>")


class OfficeTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix="office-proofing-unit-")
        self.addCleanup(self.tmp.cleanup)
        self.base = Path(self.tmp.name)

    def test_word_preserves_body_revision_media_and_normalizes_false_flags(self):
        source = self.base / "example.docx"
        word_package(source)
        digest, text = sha256(source), visible_text(source)
        report = clean(source)
        self.assertEqual(sha256(source), digest)
        self.assertEqual(visible_text(Path(report["output"])), text)
        self.assertFalse(report["future_proofing_preserved"])
        self.assertFalse(report["application_settings_modified"])
        after = word_state(Path(report["output"]))
        self.assertEqual(after["proof_error_markers"], 0)
        self.assertEqual(after["runs"], after["runs_explicitly_no_proof"])
        self.assertTrue(after["hide_spelling_errors"] and after["hide_grammar_errors"])
        with zipfile.ZipFile(source) as a, zipfile.ZipFile(report["output"]) as b:
            self.assertEqual(a.read("word/media/image1.bin"), b.read("word/media/image1.bin"))
            self.assertEqual(a.read("word/_rels/document.xml.rels"), b.read("word/_rels/document.xml.rels"))

    def test_ppt_routes_to_language_repair(self):
        source = self.base / "example.pptx"
        package(source, document('<a:r><a:rPr err="1"/><a:t>中文文字</a:t></a:r>'))
        report = clean(source)
        self.assertEqual(report["changes"][0]["new_language"], "zh-CN")
        self.assertEqual(report["after"]["no_proof_attributes"], 0)
        self.assertTrue(report["source_unchanged"])

    def test_same_named_formats_and_repeat_calls_do_not_overwrite(self):
        word, ppt = self.base / "same.docx", self.base / "same.pptx"
        word_package(word)
        package(ppt, document('<a:r><a:t>正文</a:t></a:r>'))
        a, b, c = clean(word), clean(ppt), clean(word)
        self.assertEqual(len({a["output"], b["output"], c["output"]}), 3)
        for source in (word, ppt):
            with self.assertRaises(FileExistsError):
                clean(source, source)
        with self.assertRaises(FileExistsError):
            clean(word, a["output"])

    def test_signed_and_wrong_extensions_are_rejected(self):
        source = self.base / "signed.docx"
        word_package(source, signed=True)
        digest = sha256(source)
        with self.assertRaises(ValueError):
            clean(source)
        self.assertEqual(sha256(source), digest)
        with self.assertRaises(ValueError):
            clean(source, self.base / "wrong.pptx")
        self.assertEqual(list(self.base.glob('*_校对处理副本*')), [])


if __name__ == "__main__":
    unittest.main()
