
from pathlib import Path
from docx import Document



def read_document(file_path):
    """Read paragraph text and direct formatting from a Word file."""

    path = Path(file_path)

    if not path.is_file():
        raise FileNotFoundError(f"File not found: {path}")

    doc = Document(path)

    print(f"Document: {path.name}")
    print(f"Total paragraphs: {len(doc.paragraphs)}")
    print("-" * 50)

    for index, paragraph in enumerate(doc.paragraphs, start=1):
        if not paragraph.text.strip():
            continue

        print(f"\nParagraph {index}")
        print(f"Style: {paragraph.style.name}")
        print(f"Text: {paragraph.text}")

        spacing = paragraph.paragraph_format.line_spacing
        print(f"Direct line spacing: {spacing}")

        for run_index, run in enumerate(paragraph.runs, start=1):
            if not run.text.strip():
                continue

            font = run.font

            print(f"  Run {run_index}: {run.text!r}")
            print(f"    Font: {font.name}")
            print(
                f"    Size: "
                f"{font.size.pt if font.size else None} pt"
            )
            print(f"    Bold: {font.bold}")
            print(f"    Italic: {font.italic}")

        print("-" * 50)


if __name__ == "__main__":
    read_document("examples/format_test.docx")
