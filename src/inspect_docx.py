
from pathlib import Path
from zipfile import ZipFile, BadZipFile


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

    except BadZipFile as exc:
        raise ValueError("文件不是有效的 DOCX 压缩包") from exc


if __name__ == "__main__":
    inspect_docx("examples/format_test.docx")
