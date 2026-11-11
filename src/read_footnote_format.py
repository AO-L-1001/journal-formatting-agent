
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET


def read_footnote_format(file_path):
    """读取 Word 脚注中直接设置的字体格式。"""

    path = Path(file_path)

    if not path.is_file():
        raise FileNotFoundError(f"找不到文件：{path}")

    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    }

    attr = f"{{{ns['w']}}}"

    with ZipFile(path, "r") as docx:
        if "word/footnotes.xml" not in docx.namelist():
            print("文档没有脚注")
            return

        root = ET.fromstring(docx.read("word/footnotes.xml"))

        for footnote in root.findall("w:footnote", ns):
            # 跳过脚注分隔符等特殊记录
            if footnote.get(attr + "type") is not None:
                continue

            footnote_id = footnote.get(attr + "id")

            print(f"\nWord内部脚注ID：{footnote_id}")
            print("-" * 40)

            for run in footnote.findall(".//w:r", ns):
                text = "".join(
                    t.text or ""
                    for t in run.findall("w:t", ns)
                )

                if not text.strip():
                    continue

                properties = run.find("w:rPr", ns)

                font_name = None
                font_size = None
                bold = None
                italic = None

                if properties is not None:
                    fonts = properties.find("w:rFonts", ns)

                    if fonts is not None:
                        font_name = (
                                fonts.get(attr + "eastAsia")
                                or fonts.get(attr + "ascii")
                                or fonts.get(attr + "hAnsi")
                        )

                    size = properties.find("w:sz", ns)

                    if size is not None:
                        value = size.get(attr + "val")
                        if value is not None:
                            font_size = int(value) / 2

                    for tag, name in [("b", "bold"), ("i", "italic")]:
                        element = properties.find(f"w:{tag}", ns)

                        if element is not None:
                            value = element.get(attr + "val")
                            enabled = value not in ("0", "false", "off")

                            if name == "bold":
                                bold = enabled
                            else:
                                italic = enabled

                print(f"文字：{text}")
                print(f"直接设置字体：{font_name}")
                print(f"直接设置字号：{font_size} pt")
                print(f"直接设置加粗：{bold}")
                print(f"直接设置斜体：{italic}")
                print("-" * 40)


if __name__ == "__main__":
    read_footnote_format("examples/footnote_test.docx")
