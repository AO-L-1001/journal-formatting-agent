
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET


def inspect_theme_fonts(file_path):
    """读取 Word 文档主题中的主要字体和次要字体。"""

    path = Path(file_path)

    if not path.is_file():
        raise FileNotFoundError(f"找不到文件：{path}")

    namespace = {
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main"
    }

    with ZipFile(path, "r") as docx:
        theme_files = [
            name for name in docx.namelist()
            if name.startswith("word/theme/")
               and name.endswith(".xml")
        ]

        if not theme_files:
            print("文档中未发现主题字体文件")
            return

        for theme_file in theme_files:
            print(f"正在检查主题文件：{theme_file}")

            root = ET.fromstring(docx.read(theme_file))
            font_scheme = root.find(".//a:fontScheme", namespace)

            if font_scheme is None:
                print("未找到字体方案")
                continue

            for label, tag in [
                ("主要字体", "majorFont"),
                ("次要字体", "minorFont")
            ]:
                font_group = font_scheme.find(f"a:{tag}", namespace)

                if font_group is None:
                    continue

                print(f"\n{label}：")

                latin = font_group.find("a:latin", namespace)
                east_asia = font_group.find("a:ea", namespace)

                if latin is not None:
                    print(f"默认西文字体：{latin.get('typeface')}")

                if east_asia is not None:
                    print(f"默认东亚字体：{east_asia.get('typeface')}")

                for font in font_group.findall("a:font", namespace):
                    script = font.get("script")
                    typeface = font.get("typeface")

                    if script in ("Hans", "Hant"):
                        print(f"中文字体配置（{script}）：{typeface}")


if __name__ == "__main__":
    inspect_theme_fonts("examples/footnote_test.docx")
