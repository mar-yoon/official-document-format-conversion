import io
import json
import threading
import unittest
import urllib.request
import zipfile
from http.server import ThreadingHTTPServer
from xml.etree import ElementTree as ET

from app import Handler
from formatter import FormatError, W, format_docx, inspect


def fixture():
    texts = ["某某机关文件", "某发〔2026〕1号", "关于开展规范化工作的通知", "各部门：", "一、工作要求", "请按时完成。", "某某机关", "2026年10月26日"]
    ps = "".join(f'<w:p><w:r><w:t>{text}</w:t></w:r></w:p>' for text in texts)
    xml = f'<?xml version="1.0"?><w:document xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"><w:body>{ps}<w:sectPr/></w:body></w:document>'
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as zf:
        zf.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        zf.writestr("word/document.xml", xml)
        zf.writestr("word/media/image1.png", b"image-bytes")
    return out.getvalue()


class FormatterTests(unittest.TestCase):
    def test_format_preserves_media_and_sets_roles(self):
        original = fixture()
        output, report = format_docx(original)
        self.assertEqual(report["title"], "关于开展规范化工作的通知")
        self.assertEqual(report["roles"]["heading_1"], 1)
        with zipfile.ZipFile(io.BytesIO(output)) as zf:
            self.assertEqual(zf.read("word/media/image1.png"), b"image-bytes")
            document_xml = zf.read("word/document.xml")
            self.assertIn(b"xmlns:w14=", document_xml)
            root = ET.fromstring(document_xml)
        q = lambda x: f"{{{W}}}{x}"
        self.assertEqual(root.find(f".//{q('pgSz')}").get(q("w")), "11906")
        self.assertEqual(root.find(f".//{q('pgMar')}").get(q("left")), "1587")
        paragraphs = list(root.iter(q("p")))
        self.assertEqual(paragraphs[2].find(f"{q('pPr')}/{q('jc')}").get(q("val")), "center")
        self.assertEqual(paragraphs[4].find(f".//{q('rFonts')}").get(q("eastAsia")), "黑体")

    def test_invalid_input_rejected(self):
        with self.assertRaises(FormatError):
            inspect(b"not a docx")

    def test_mini_program_upload_download(self):
        server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            boundary = "test-boundary"
            body = (f"--{boundary}\r\nContent-Disposition: form-data; name=\"file\"; filename=\"sample.docx\"\r\nContent-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n").encode() + fixture() + f"\r\n--{boundary}--\r\n".encode()
            base = f"http://127.0.0.1:{server.server_port}"
            request = urllib.request.Request(base + "/api/jobs?template=gov_official", data=body, headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}, method="POST")
            with urllib.request.urlopen(request) as response:
                result = json.load(response)
            self.assertEqual(result["report"]["title"], "关于开展规范化工作的通知")
            with urllib.request.urlopen(base + result["download_path"]) as response:
                downloaded = response.read()
            self.assertTrue(zipfile.is_zipfile(io.BytesIO(downloaded)))
        finally:
            server.shutdown()
            server.server_close()


if __name__ == "__main__":
    unittest.main()
