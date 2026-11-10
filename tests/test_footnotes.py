
from src.read_footnotes import read_footnotes


def test_footnotes():
    # 读取测试用 Word 文件
    results = read_footnotes("examples/footnote_test.docx")

    # 检查是否成功读取两条脚注
    assert len(results) == 2, "脚注数量应该为2"

    # 检查脚注显示序号是否从1开始
    assert results[0]["number"] == 1
    assert results[1]["number"] == 2

    # 检查脚注内部ID
    assert results[0]["id"] == "0"
    assert results[1]["id"] == "1"

    # 检查脚注内容是否正确
    assert "张三" in results[0]["text"]
    assert "里斯" in results[1]["text"]


if __name__ == "__main__":
    test_footnotes()
    print("所有脚注测试均已通过！")
