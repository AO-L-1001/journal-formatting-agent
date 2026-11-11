
from zipfile import ZipFile
import xml.etree.ElementTree as ET


def inspect_style_inheritance(file_path):
    """检查脚注样式的继承关系和文档默认字体。"""

    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    }
    attr = f"{{{ns['w']}}}"

    with ZipFile(file_path, "r") as docx:
        root = ET.fromstring(docx.read("word/styles.xml"))

        styles = {
            style.get(attr + "styleId"): style
            for style in root.findall("w:style", ns)
        }

        # 查找脚注文本样式
        for style_id, style in styles.items():
            name = style.find("w:name", ns)

            if name is None:
                continue

            if name.get(attr + "val") != "footnote text":
                continue

            print(f"脚注样式 ID：{style_id}")

            based_on = style.find("w:basedOn", ns)

            if based_on is not None:
                parent_id = based_on.get(attr + "val")
                print(f"继承的父样式 ID：{parent_id}")
            else:
                print("没有明确指定父样式")

        # 检查文档默认字符格式
        defaults = root.find("w:docDefaults", ns)

        if defaults is None:
            print("没有找到文档默认格式")
            return

        run_defaults = defaults.find("w:rPrDefault/w:rPr", ns)

        if run_defaults is None:
            print("没有找到默认字符格式")
            return

        fonts = run_defaults.find("w:rFonts", ns)

        if fonts is None:
            print("没有找到默认字体定义")
            return

        print("\n文档默认字体设置：")

        for name in ["ascii", "hAnsi", "eastAsia",
                     "asciiTheme", "hAnsiTheme", "eastAsiaTheme"]:
            print(f"{name}: {fonts.get(attr + name)}")


if __name__ == "__main__":
    inspect_style_inheritance("examples/footnote_test.docx")
