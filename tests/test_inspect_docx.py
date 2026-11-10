
import unittest
from pathlib import Path
from zipfile import ZipFile

from src.inspect_docx import count_footnotes


class TestInspectDocx(unittest.TestCase):
    """测试 Word 文档的脚注识别功能。"""

    def test_no_footnotes(self):
        """测试没有脚注的文档。"""

        path = Path("examples/format_test.docx")

        with ZipFile(path, "r") as docx:
            count = count_footnotes(docx)

        self.assertEqual(count, 0)

    def test_missing_file(self):
        """测试不存在的文件是否正确报错。"""

        path = Path("examples/not_exist.docx")

        with self.assertRaises(FileNotFoundError):
            with ZipFile(path, "r"):
                pass


if __name__ == "__main__":
    unittest.main()
