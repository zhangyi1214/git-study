#!/usr/bin/env python3
"""
Meta-analysis PDF data extractor.

功能：
1. 从指定文件夹批量读取 PDF（默认按文件创建时间排序，尽可能贴近“上传顺序”）。
2. 按预设字段抽取信息，缺失值统一填充为 "NR"。
3. 将干预组/对照组拆成多行导出到 Excel。
4. 检测是否提到补充材料（supplementary/appendix 等），并输出提醒文件。

说明：
- 这是“半自动抽取”脚本：会尽量通过正则和关键词抽取，复杂表格建议人工复核。
- 依赖：pandas, openpyxl, pypdf
  安装：pip install pandas openpyxl pypdf
"""

from __future__ import annotations

import argparse
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
from pypdf import PdfReader


COLUMNS = [
    "标题",
    "编号",
    "作者",
    "年份",
    "注册号",
    "是否青少年",
    "纳入的患者双相类型及人数",
    "诊断标准",
    "治疗类型",
    "治疗方案（剂量）",
    "给药方式",
    "治疗时长",
    "样本量",
    "性别（女性数量）",
    "年龄平均",
    "地区",
    "盲法",
    "基线抑郁评分（SD）",
    "终点抑郁评分（SD）",
    "抑郁变化（SD）",
    "达到反应标准人数",
    "达到缓解标准的人数",
    "转躁(人数)",
    "全因停药人数及率",
    "总体不良事件发生人数",
]

NR = "NR"


@dataclass
class Arm:
    name: str
    fields: Dict[str, str]


def clean_text(s: str) -> str:
    s = s.replace("\x00", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def read_pdf_text(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path))
    texts: List[str] = []
    for page in reader.pages:
        texts.append(page.extract_text() or "")
    return clean_text("\n".join(texts))


def first_match(text: str, patterns: List[str], flags: int = re.I) -> str:
    for p in patterns:
        m = re.search(p, text, flags)
        if m:
            if m.groups():
                return clean_text(m.group(1))
            return clean_text(m.group(0))
    return NR


def infer_year(text: str) -> str:
    m = re.search(r"\b(19\d{2}|20\d{2})\b", text)
    return m.group(1) if m else NR


def infer_title(text: str, filename: str) -> str:
    # 粗略策略：用前 300 字中最长句作为标题候选。
    head = text[:500]
    parts = re.split(r"[\n\r\.。;；]", head)
    parts = [p.strip() for p in parts if len(p.strip()) > 15]
    if parts:
        return max(parts, key=len)[:180]
    return filename


def parse_common_fields(text: str, file_stem: str) -> Dict[str, str]:
    fields = {c: NR for c in COLUMNS}
    fields["编号"] = file_stem
    fields["标题"] = infer_title(text, file_stem)

    fields["作者"] = first_match(
        text,
        [
            r"([A-Z][a-zA-Z\-]+\s+(?:et al\.?|and\s+[A-Z][a-zA-Z\-]+))",
            r"Authors?\s*[:：]\s*([^\n\r]{3,120})",
        ],
    )
    fields["年份"] = infer_year(text)
    fields["注册号"] = first_match(
        text,
        [
            r"(?:ClinicalTrials\.gov|Trial registration|Registration)\s*[:：#]?\s*(NCT\d{8})",
            r"(ChiCTR[-\w]+)",
            r"(ISRCTN\d+)",
        ],
    )

    adolescent_flag = first_match(
        text,
        [
            r"\b(adolescen\w+|youth|pediatric|children)\b",
            r"\b(青少年|儿童|未成年人)\b",
        ],
    )
    fields["是否青少年"] = "是" if adolescent_flag != NR else "否/NR"

    fields["纳入的患者双相类型及人数"] = first_match(
        text,
        [
            r"(bipolar\s+(?:I|II|1|2)[^\.;]{0,120})",
            r"(双相[^\.;]{0,120})",
        ],
    )
    fields["诊断标准"] = first_match(
        text,
        [r"\b(DSM-?5|DSM-?IV-?TR|ICD-?10|ICD-?11|MINI|SCID)\b"],
    )
    fields["性别（女性数量）"] = first_match(
        text,
        [
            r"female[s]?\s*(?:n\s*=|=)?\s*(\d+)",
            r"女性\s*(\d+)",
        ],
    )
    fields["年龄平均"] = first_match(
        text,
        [r"(?:mean\s+age|年龄)\s*(?:=|:)?\s*([0-9]+(?:\.[0-9]+)?)"],
    )
    fields["地区"] = first_match(
        text,
        [
            r"\b(multicenter|single-center|United States|China|Europe|Japan|Korea|India)\b",
            r"(中国|美国|欧洲|日本|韩国|印度|多中心|单中心)",
        ],
    )
    fields["盲法"] = first_match(
        text,
        [
            r"\b(double-blind|single-blind|open-label)\b",
            r"(双盲|单盲|开放标签)",
        ],
    )

    fields["基线抑郁评分（SD）"] = first_match(
        text,
        [
            r"baseline[^\n\r]{0,60}(MADRS|HAM-D|HDRS|BDI)[^\n\r]{0,60}",
            r"(基线[^\n\r]{0,80}抑郁[^\n\r]{0,80})",
        ],
    )
    fields["终点抑郁评分（SD）"] = first_match(
        text,
        [
            r"endpoint[^\n\r]{0,60}(MADRS|HAM-D|HDRS|BDI)[^\n\r]{0,60}",
            r"(终点[^\n\r]{0,80}抑郁[^\n\r]{0,80})",
        ],
    )
    fields["抑郁变化（SD）"] = first_match(
        text,
        [
            r"change\s+from\s+baseline[^\n\r]{0,100}",
            r"(变化值[^\n\r]{0,80})",
        ],
    )
    fields["达到反应标准人数"] = first_match(
        text,
        [r"response\s*(?:rate)?[^\d]{0,20}(\d+)", r"反应[^\d]{0,20}(\d+)"],
    )
    fields["达到缓解标准的人数"] = first_match(
        text,
        [r"remission\s*(?:rate)?[^\d]{0,20}(\d+)", r"缓解[^\d]{0,20}(\d+)"],
    )
    fields["转躁(人数)"] = first_match(
        text,
        [r"(?:mania|hypomania|switch)[^\d]{0,20}(\d+)", r"转躁[^\d]{0,20}(\d+)"],
    )
    fields["全因停药人数及率"] = first_match(
        text,
        [
            r"all-cause\s+discontinuation[^\.;]{0,80}",
            r"(总停药[^\.;]{0,80})",
        ],
    )
    fields["总体不良事件发生人数"] = first_match(
        text,
        [
            r"adverse\s+event[s]?[^\d]{0,20}(\d+)",
            r"不良事件[^\d]{0,20}(\d+)",
        ],
    )
    return fields


def mention_supplementary(text: str) -> bool:
    return bool(
        re.search(
            r"\b(supplementary|supplemental|appendix|online\s+material|eTable|eFigure|补充材料|附录)\b",
            text,
            flags=re.I,
        )
    )


def infer_arms(text: str) -> List[Arm]:
    # 尝试从文本中识别分组；未识别时仍输出干预组+对照组两行。
    arm_patterns = [
        ("干预组", r"(intervention|treatment|experimental|active\s+drug|干预组|治疗组)"),
        ("对照组", r"(control|placebo|usual\s+care|对照组|安慰剂)"),
    ]

    detected = []
    for name, pat in arm_patterns:
        if re.search(pat, text, re.I):
            detected.append(name)

    if not detected:
        detected = ["干预组", "对照组"]

    arms: List[Arm] = []
    for arm_name in detected:
        fields = {
            "治疗类型": arm_name,
            "治疗方案（剂量）": NR,
            "给药方式": NR,
            "治疗时长": NR,
            "样本量": NR,
        }

        if arm_name == "干预组":
            fields["治疗方案（剂量）"] = first_match(
                text,
                [
                    r"(?:intervention|treatment|experimental)[^\.;]{0,120}",
                    r"(?:给予|使用)[^\.;]{0,120}",
                ],
            )
        else:
            fields["治疗方案（剂量）"] = first_match(
                text,
                [
                    r"(?:control|placebo|usual care)[^\.;]{0,120}",
                    r"(?:对照|安慰剂)[^\.;]{0,120}",
                ],
            )

        fields["给药方式"] = first_match(
            text, [r"\b(oral|iv|intravenous|intramuscular|subcutaneous)\b", r"(口服|静脉|肌注|皮下)"]
        )
        fields["治疗时长"] = first_match(
            text,
            [r"(?:for|duration|week|weeks|months)[^\.;]{0,40}", r"(\d+\s*(?:周|月|天))"],
        )
        fields["样本量"] = first_match(
            text,
            [
                r"\bn\s*=\s*(\d+)\b",
                r"sample\s+size[^\d]{0,10}(\d+)",
                r"样本量[^\d]{0,10}(\d+)",
            ],
        )
        arms.append(Arm(name=arm_name, fields=fields))

    return arms


def list_pdfs(input_dir: Path, order: str) -> List[Path]:
    pdfs = [p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]
    if order == "ctime":
        pdfs.sort(key=lambda p: (p.stat().st_ctime, p.name.lower()))
    elif order == "mtime":
        pdfs.sort(key=lambda p: (p.stat().st_mtime, p.name.lower()))
    else:  # name
        pdfs.sort(key=lambda p: p.name.lower())
    return pdfs


def extract_from_folder(input_dir: Path, output_excel: Path, order: str, alerts_file: Path) -> None:
    rows: List[Dict[str, str]] = []
    supplemental_hits: List[str] = []

    pdf_files = list_pdfs(input_dir, order)
    if not pdf_files:
        raise FileNotFoundError(f"在目录中未发现 PDF 文件: {input_dir}")

    for pdf in pdf_files:
        text = read_pdf_text(pdf)
        common = parse_common_fields(text, pdf.stem)

        if mention_supplementary(text):
            supplemental_hits.append(pdf.name)

        arms = infer_arms(text)
        for arm in arms:
            row = {c: NR for c in COLUMNS}
            row.update(common)
            row.update(arm.fields)
            rows.append(row)

    df = pd.DataFrame(rows, columns=COLUMNS)
    output_excel.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_excel, index=False)

    alerts_file.parent.mkdir(parents=True, exist_ok=True)
    with open(alerts_file, "w", encoding="utf-8") as f:
        if supplemental_hits:
            f.write("以下文献提到补充材料/附录（请优先人工核查补充文件中的数据）：\n")
            for name in supplemental_hits:
                f.write(f"- {name}\n")
        else:
            f.write("未在文献文本中检测到补充材料关键词。\n")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Meta 分析 PDF 数据提取脚本")
    parser.add_argument("-i", "--input-dir", required=True, help="PDF 文件夹路径")
    parser.add_argument("-o", "--output", default="meta_extraction.xlsx", help="输出 Excel 文件路径")
    parser.add_argument(
        "--order",
        choices=["ctime", "mtime", "name"],
        default="ctime",
        help="文件处理顺序：ctime(默认,创建时间) / mtime(修改时间) / name(文件名)",
    )
    parser.add_argument(
        "--alerts",
        default="supplementary_alerts.txt",
        help="补充材料提醒输出路径",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    input_dir = Path(args.input_dir).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()
    alerts_path = Path(args.alerts).expanduser().resolve()

    if not input_dir.exists() or not input_dir.is_dir():
        raise NotADirectoryError(f"输入目录无效: {input_dir}")

    extract_from_folder(input_dir, output_path, args.order, alerts_path)
    print(f"提取完成：{output_path}")
    print(f"补充材料提醒：{alerts_path}")


if __name__ == "__main__":
    main()
