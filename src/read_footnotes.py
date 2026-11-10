
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET


def read_footnotes(file_path):
    """读取 Word 文档中正文引用的普通脚注及其文字。"""

    path = Path(file_path)

    if not path.is_file():
        raise FileNotFoundError(f"找不到文件：{path}")

    namespace = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    }

    id_attribute = f"{{{namespace['w']}}}id"
    type_attribute = f"{{{namespace['w']}}}type"

    with ZipFile(path, "r") as docx:
        files = set(docx.namelist())

        if "word/footnotes.xml" not in files:
            return []

        # 从正文中获取实际引用的脚注编号
        document_root = ET.fromstring(docx.read("word/document.xml"))

        reference_ids = [
            element.get(id_attribute)
            for element in document_root.findall(
                ".//w:footnoteReference", namespace
            )
        ]

        # 读取脚注 XML 文件
        footnotes_root = ET.fromstring(docx.read("word/footnotes.xml"))

        footnotes = {}

        for footnote in footnotes_root.findall("w:footnote", namespace):
            footnote_id = footnote.get(id_attribute)
            footnote_type = footnote.get(type_attribute)

            # 跳过分隔符等特殊脚注记录
            if footnote_type is not None:
                continue

            paragraphs = []

            for paragraph in footnote.findall("w:p", namespace):
                text = "".join(
                    element.text or ""
                    for element in paragraph.findall(".//w:t", namespace)
                )

                if text.strip():
                    paragraphs.append(text)

            footnotes[footnote_id] = "\n".join(paragraphs)

        # 按正文中的引用顺序返回脚注
    results = []

    for footnote_id in reference_ids:
        if footnote_id in footnotes:
            results.append({
                "number": len(results) + 1,  # 从1开始的显示序号
                "id": footnote_id,           # 保留Word原始ID
                "text": footnotes[footnote_id]
            })

    return results


if __name__ == "__main__":
    file_path = "examples/footnote_test.docx"

    results = read_footnotes(file_path)

    print(f"成功读取 {len(results)} 条脚注")

    for footnote in results:
        print("-" * 40)
        print(f"脚注序号：{footnote['number']}")
        print(f"Word内部ID：{footnote['id']}")
        print(f"脚注内容：{footnote['text']}")
