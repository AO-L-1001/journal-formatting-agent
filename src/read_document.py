
from pathlib import Path
from docx import Document


def read_document(file_path):
    """Read paragraphs and styles from a Word document."""

    path = Path(file_path)

    if not path.is_file():
        raise FileNotFoundError(f"File not found: {path}")

    doc = Document(path)

    print(f"Document: {path.name}")
    print(f"Total paragraphs: {len(doc.paragraphs)}")
    print("-" * 40)

    for index, paragraph in enumerate(doc.paragraphs, start=1):
        if not paragraph.text.strip():
            continue

        print(f"Paragraph {index}")
        print(f"Style: {paragraph.style.name}")
        print(f"Text: {paragraph.text}")
        print("-" * 40)


if __name__ == "__main__":
    read_document("examples/sample.docx")
