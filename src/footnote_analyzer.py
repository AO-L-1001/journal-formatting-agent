
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

NS = {"w": W, "a": A}


def 属性名(名称):
    """生成 Word XML 属性的完整名称。"""
    return f"{{{W}}}{名称}"


def 读取主题字体(压缩包):
    """读取文档主题中的主要和次要字体。"""
    结果 = {"major": {}, "minor": {}}
    路径 = "word/theme/theme1.xml"

    if 路径 not in 压缩包.namelist():
        return 结果

    根节点 = ET.fromstring(压缩包.read(路径))

    for 类型 in ("major", "minor"):
        节点 = 根节点.find(f".//a:{类型}Font", NS)

        if 节点 is None:
            continue

        拉丁字体 = 节点.find("a:latin", NS)
        东亚字体 = 节点.find("a:ea", NS)

        if 拉丁字体 is not None:
            结果[类型]["西文"] = 拉丁字体.get("typeface")

        if 东亚字体 is not None:
            结果[类型]["东亚"] = 东亚字体.get("typeface")

        for 字体 in 节点.findall("a:font", NS):
            语言 = 字体.get("script")

            if 语言 in ("Hans", "Hant"):
                结果[类型][语言] = 字体.get("typeface")

    return 结果


def 读取样式(压缩包):
    """建立 Word 样式 ID 与样式定义的对应关系。"""
    路径 = "word/styles.xml"

    if 路径 not in 压缩包.namelist():
        return {}

    根节点 = ET.fromstring(压缩包.read(路径))

    return {
        样式.get(属性名("styleId")): 样式
        for 样式 in 根节点.findall("w:style", NS)
    }


def 读取样式属性(样式):
    """读取单个样式直接定义的字体、字号和父样式。"""
    结果 = {
        "名称": None,
        "父样式": None,
        "字体设置": {},
        "字号_pt": None,
    }

    名称 = 样式.find("w:name", NS)
    父样式 = 样式.find("w:basedOn", NS)
    文字格式 = 样式.find("w:rPr", NS)

    if 名称 is not None:
        结果["名称"] = 名称.get(属性名("val"))

    if 父样式 is not None:
        结果["父样式"] = 父样式.get(属性名("val"))

    if 文字格式 is not None:
        字体 = 文字格式.find("w:rFonts", NS)
        字号 = 文字格式.find("w:sz", NS)

        if 字体 is not None:
            for 字段 in (
                    "ascii", "hAnsi", "eastAsia",
                    "asciiTheme", "hAnsiTheme", "eastAsiaTheme"
            ):
                值 = 字体.get(属性名(字段))

                if 值 is not None:
                    结果["字体设置"][字段] = 值

        if 字号 is not None:
            值 = 字号.get(属性名("val"))

            if 值 is not None:
                结果["字号_pt"] = int(值) / 2

    return 结果


def 合并继承样式(样式ID, 样式表):
    """从父样式到子样式逐层合并文字格式。"""
    样式链 = []
    已访问 = set()
    当前ID = 样式ID

    while 当前ID and 当前ID not in 已访问:
        已访问.add(当前ID)
        样式 = 样式表.get(当前ID)

        if 样式 is None:
            break

        信息 = 读取样式属性(样式)
        样式链.append((当前ID, 信息))
        当前ID = 信息["父样式"]

    字体设置 = {}
    字号 = None

    for _, 信息 in reversed(样式链):
        字体设置.update(信息["字体设置"])

        if 信息["字号_pt"] is not None:
            字号 = 信息["字号_pt"]

    return {
        "样式链": [
            {"ID": 编号, "名称": 信息["名称"]}
            for 编号, 信息 in 样式链
        ],
        "字体设置": 字体设置,
        "字号_pt": 字号,
    }


def 解析字体(字体设置, 主题字体):
    """根据显式字体和主题字体设置推断中西文字体。"""

    def 主题查找(名称, 语言):
        if not 名称:
            return None

        类型 = "major" if 名称.startswith("major") else "minor"
        return 主题字体.get(类型, {}).get(语言)

    中文 = 字体设置.get("eastAsia")

    if not 中文:
        中文 = 主题查找(
            字体设置.get("eastAsiaTheme"), "Hans"
        )

    西文 = 字体设置.get("ascii")

    if not 西文:
        西文 = 主题查找(
            字体设置.get("asciiTheme"), "西文"
        )

    if not 西文:
        西文 = 字体设置.get("hAnsi")

    if not 西文:
        西文 = 主题查找(
            字体设置.get("hAnsiTheme"), "西文"
        )

    return {"中文": 中文, "西文": 西文}


def 分析脚注(文件路径):
    """分析正文引用的脚注，返回结构化格式信息。"""
    路径 = Path(文件路径)

    if not 路径.is_file():
        raise FileNotFoundError(f"找不到文件：{路径}")

    with ZipFile(路径, "r") as 压缩包:
        文件列表 = set(压缩包.namelist())

        if "word/document.xml" not in 文件列表:
            raise ValueError("文件缺少 Word 正文 XML")

        if "word/footnotes.xml" not in 文件列表:
            return []

        主题字体 = 读取主题字体(压缩包)
        样式表 = 读取样式(压缩包)

        正文 = ET.fromstring(
            压缩包.read("word/document.xml")
        )

        引用编号 = [
            节点.get(属性名("id"))
            for 节点 in 正文.findall(
                ".//w:footnoteReference", NS
            )
        ]

        脚注根节点 = ET.fromstring(
            压缩包.read("word/footnotes.xml")
        )

        脚注记录 = {}

        for 脚注 in 脚注根节点.findall("w:footnote", NS):
            编号 = 脚注.get(属性名("id"))

            if 脚注.get(属性名("type")) is not None:
                continue

            段落文字 = []
            样式ID = None

            for 段落 in 脚注.findall("w:p", NS):
                段落属性 = 段落.find("w:pPr", NS)

                if 段落属性 is not None and 样式ID is None:
                    样式节点 = 段落属性.find("w:pStyle", NS)

                    if 样式节点 is not None:
                        样式ID = 样式节点.get(属性名("val"))

                文字 = "".join(
                    节点.text or ""
                    for 节点 in 段落.findall(".//w:t", NS)
                )

                if 文字.strip():
                    段落文字.append(文字)

            脚注记录[编号] = {
                "内容": "\n".join(段落文字),
                "样式ID": 样式ID,
            }

    结果 = []

    for 编号 in 引用编号:
        if 编号 not in 脚注记录:
            continue

        记录 = 脚注记录[编号]
        样式信息 = 合并继承样式(
            记录["样式ID"], 样式表
        )

        字体 = 解析字体(
            样式信息["字体设置"], 主题字体
        )

        结果.append({
            "脚注序号": len(结果) + 1,
            "Word内部ID": 编号,
            "脚注内容": 记录["内容"],
            "段落样式ID": 记录["样式ID"],
            "样式继承链": 样式信息["样式链"],
            "中文字体": 字体["中文"],
            "西文字体": 字体["西文"],
            "字号_pt": 样式信息["字号_pt"],
            "检测说明": "基于段落样式继承及主题字体推断，尚未合并文字级覆盖格式",
        })

    return 结果
