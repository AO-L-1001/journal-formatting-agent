
import json
from pathlib import Path

from src.footnote_analyzer import 分析脚注


def 生成脚注报告(输入路径, 输出路径):
    """分析 Word 脚注并生成 JSON 检测报告。"""

    脚注列表 = 分析脚注(输入路径)

    报告 = {
        "文档名称": Path(输入路径).name,
        "脚注总数": len(脚注列表),
        "检测范围": "脚注段落样式、继承关系及主题字体",
        "脚注": 脚注列表,
    }

    输出文件 = Path(输出路径)
    输出文件.parent.mkdir(
        parents=True, exist_ok=True
    )

    with 输出文件.open(
            "w", encoding="utf-8"
    ) as 文件:
        json.dump(
            报告,
            文件,
            ensure_ascii=False,
            indent=2,
        )

    print(f"检测完成，共发现 {len(脚注列表)} 条脚注")
    print(f"报告已生成：{输出文件}")


if __name__ == "__main__":
    生成脚注报告(
        "examples/footnote_test.docx",
        "output/footnote_report.json",
    )
