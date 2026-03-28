# Meta分析数据提取脚本

脚本：`meta_extract.py`

## 安装依赖

```bash
pip install pandas openpyxl pypdf
```

## 用法

```bash
python meta_extract.py -i /你的PDF目录 -o meta_extraction.xlsx --order ctime --alerts supplementary_alerts.txt
```

- `--order ctime`：按文件创建时间排序（默认，尽量贴近上传顺序）
- `--order mtime`：按修改时间排序
- `--order name`：按文件名排序

## 输出

1. `meta_extraction.xlsx`：包含以下表头（缺失数据填 `NR`）
   - 标题 | 编号 | 作者 | 年份 | 注册号 | 是否青少年 | 纳入的患者双相类型及人数 | 诊断标准 | 治疗类型 | 治疗方案（剂量） | 给药方式 | 治疗时长 | 样本量 | 性别（女性数量） | 年龄平均 | 地区 | 盲法 | 基线抑郁评分（SD） | 终点抑郁评分（SD） | 抑郁变化（SD） | 达到反应标准人数 | 达到缓解标准的人数 | 转躁(人数) | 全因停药人数及率 | 总体不良事件发生人数
2. `supplementary_alerts.txt`：若文献提到补充材料/附录，会在这里列出对应 PDF 文件名。

## 注意

- 脚本会尝试将干预组和对照组拆分为多行输出；若文本无法明确识别，仍会输出干预组和对照组占位行。
- PDF格式差异较大，建议对导出结果进行人工复核。
