
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET


def inspect_footnote_styles(file_path):
    """检查 Word 脚注使用的段落样式及样式中的字体设置。"""

    path = Path(file_path)

    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    }
    attr = f"{{{ns['w']}}}"

    with ZipFile(path, "r") as docx:
        files = set(docx.namelist())

        if "word/footnotes.xml" not in files:
            print("文档中没有脚注文件")
            return

        if "word/styles.xml" not in files:
            print("文档中没有样式定义文件")
            return

        footnotes_root = ET.fromstring(docx.read("word/footnotes.xml"))
        styles_root = ET.fromstring(docx.read("word/styles.xml"))

        styles = {}

        for style in styles_root.findall("w:style", ns):
            style_id = style.get(attr + "styleId")
            styles[style_id] = style

        for footnote in footnotes_root.findall("w:footnote", ns):
            if footnote.get(attr + "type") is not None:
                continue

            footnote_id = footnote.get(attr + "id")
            print(f"\n脚注 ID：{footnote_id}")

            for paragraph in footnote.findall("w:p", ns):
                properties = paragraph.find("w:pPr", ns)

                if properties is None:
                    print("未直接指定段落样式")
                    continue

                style_element = properties.find("w:pStyle", ns)

                if style_element is None:
                    print("未直接指定段落样式")
                    continue

                style_id = style_element.get(attr + "val")
                print(f"段落样式 ID：{style_id}")

                style = styles.get(style_id)

                if style is None:
                    print("未找到对应的样式定义")
                    continue

                name_element = style.find("w:name", ns)
                style_name = (
                    name_element.get(attr + "val")
                    if name_element is not None
                    else None
                )

                print(f"样式名称：{style_name}")

                run_properties = style.find("w:rPr", ns)

                if run_properties is None:
                    print("样式中没有直接定义字符格式")
                    continue

                fonts = run_properties.find("w:rFonts", ns)
                size = run_properties.find("w:sz", ns)

                font_name = None
                font_size = None

                if fonts is not None:
                    font_name = (
                            fonts.get(attr + "eastAsia")
                            or fonts.get(attr + "ascii")
                            or fonts.get(attr + "hAnsi")
                    )

                if size is not None:
                    value = size.get(attr + "val")
                    if value is not None:
                        font_size = int(value) / 2

                print(f"样式直接定义的字体：{font_name}")
                print(f"样式直接定义的字号：{font_size} pt")


if __name__ == "__main__":
    inspect_footnote_styles("examples/footnote_test.docx")
