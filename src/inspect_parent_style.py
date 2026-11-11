
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET


def inspect_parent_style(file_path):
    """检查脚注样式及其父样式的字体和字号。"""

    path = Path(file_path)

    if not path.is_file():
        raise FileNotFoundError(f"找不到文件：{path}")

    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    }

    def attr(name):
        return f"{{{ns['w']}}}{name}"

    with ZipFile(path, "r") as docx:
        root = ET.fromstring(docx.read("word/styles.xml"))

    styles = {
        style.get(attr("styleId")): style
        for style in root.findall("w:style", ns)
    }

    current_id = "2"
    visited = set()

    while current_id and current_id not in visited:
        visited.add(current_id)

        style = styles.get(current_id)

        if style is None:
            print(f"找不到样式：{current_id}")
            break

        name = style.find("w:name", ns)
        based_on = style.find("w:basedOn", ns)
        run_properties = style.find("w:rPr", ns)

        print("-" * 45)
        print(f"样式 ID：{current_id}")
        print(
            f"样式名称："
            f"{name.get(attr('val')) if name is not None else '未知'}"
        )

        if run_properties is not None:
            fonts = run_properties.find("w:rFonts", ns)
            size = run_properties.find("w:sz", ns)

            if fonts is not None:
                for key in (
                        "ascii", "hAnsi", "eastAsia",
                        "asciiTheme", "hAnsiTheme", "eastAsiaTheme"
                ):
                    print(f"{key}：{fonts.get(attr(key))}")
            else:
                print("该样式没有直接指定字体")

            if size is not None:
                value = size.get(attr("val"))
                print(f"字号：{int(value) / 2} pt")
            else:
                print("该样式没有直接指定字号")
        else:
            print("该样式没有直接定义文字格式")

        current_id = (
            based_on.get(attr("val"))
            if based_on is not None
            else None
        )

    print("-" * 45)
    print("样式继承关系检查完成")


if __name__ == "__main__":
    inspect_parent_style("examples/footnote_test.docx")
