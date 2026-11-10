
from pathlib import Path
from zipfile import ZipFile, BadZipFile
import xml.etree.ElementTree as ET


def count_footnotes(docx):
    """统计正文实际引用的普通脚注数量。"""

    files = set(docx.namelist())

    if "word/footnotes.xml" not in files:
        return 0

    namespace = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    }

    # 读取正文 XML，寻找真正的脚注引用
    document_root = ET.fromstring(docx.read("word/document.xml"))

    reference_ids = {
        element.get(f"{{{namespace['w']}}}id")
        for element in document_root.findall(".//w:footnoteReference", namespace)
    }

    # 读取脚注 XML，排除分隔符等特殊记录
    footnotes_root = ET.fromstring(docx.read("word/footnotes.xml"))

    actual_ids = set()

    for footnote in footnotes_root.findall("w:footnote", namespace):
        footnote_id = footnote.get(f"{{{namespace['w']}}}id")
        footnote_type = footnote.get(f"{{{namespace['w']}}}type")

        if footnote_type is None:
            actual_ids.add(footnote_id)

    return len(reference_ids & actual_ids)

def inspect_docx(file_path):
    """检查 Word 文档的内部结构，不修改原始文件。"""

    path = Path(file_path)

    if not path.is_file():
        raise FileNotFoundError(f"找不到文件：{path}")

    if path.suffix.lower() != ".docx":
        raise ValueError("目前只支持 .docx 格式的 Word 文件")

    try:
        with ZipFile(path, "r") as docx:
            files = set(docx.namelist())

            components = {
                "正文": "word/document.xml",
                "样式": "word/styles.xml",
                "脚注": "word/footnotes.xml",
                "尾注": "word/endnotes.xml",
                "编号": "word/numbering.xml",
            }

            print(f"正在检查：{path.name}")
            print("-" * 40)

            for name, location in components.items():
                status = "存在" if location in files else "未发现"
                print(f"{name}：{status}")

            print("-" * 40)
            print(f"内部文件总数：{len(files)}")
            footnote_count = count_footnotes(docx)
            print(f"实际脚注数量：{footnote_count}")
    except BadZipFile as exc:
        raise ValueError("文件不是有效的 DOCX 压缩包") from exc


if __name__ == "__main__":
    inspect_docx("examples/format_test.docx")
