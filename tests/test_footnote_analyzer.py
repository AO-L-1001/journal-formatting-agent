
import unittest
from pathlib import Path

from src.footnote_analyzer import (
    分析脚注,
    合并继承样式,
    解析字体,
)


项目目录 = Path(__file__).resolve().parents[1]


class 脚注检测测试(unittest.TestCase):
    """检查脚注分析器的核心功能。"""

    def test_测试文档包含两条脚注(self):
        路径 = 项目目录 / "examples" / "footnote_test.docx"

        if not 路径.is_file():
            self.skipTest("缺少脚注测试文档")

        结果 = 分析脚注(路径)

        self.assertEqual(len(结果), 2)
        self.assertEqual(
            [脚注["脚注序号"] for 脚注 in 结果],
            [1, 2],
        )
        self.assertTrue(
            all(脚注["脚注内容"] for 脚注 in 结果)
        )

    def test_不存在的文件应报错(self):
        路径 = 项目目录 / "examples" / "不存在的文件.docx"

        with self.assertRaises(FileNotFoundError):
            分析脚注(路径)

    def test_主题字体解析(self):
        字体设置 = {
            "asciiTheme": "minorHAnsi",
            "eastAsiaTheme": "minorEastAsia",
        }

        主题字体 = {
            "minor": {
                "西文": "Calibri",
                "Hans": "宋体",
            }
        }

        结果 = 解析字体(字体设置, 主题字体)

        self.assertEqual(结果["中文"], "宋体")
        self.assertEqual(结果["西文"], "Calibri")

    def test_未知样式不应导致程序崩溃(self):
        结果 = 合并继承样式("不存在的样式", {})

        self.assertEqual(结果["样式链"], [])
        self.assertIsNone(结果["字号_pt"])


if __name__ == "__main__":
    unittest.main()
