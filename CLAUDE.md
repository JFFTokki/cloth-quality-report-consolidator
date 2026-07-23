# 项目协作规则

本项目整理`2026Q4`质检报告。业务与格式规则以`docs/`为权威来源，脚本位于`pipeline/qc_all/`。

## 硬规则

- 第一规则：报告出现繁体文字时，必须先将全文转换为简体，再进行报告字段、表头、检测项、结果和元数据识别；繁体原文另行保留用于追溯。
- 任何单元格中的全部PDF链接都必须处理。
- 不以最后修改时间丢弃旧报告；原始项名、报告号和PDF链接必须可追溯。
- 横向汇总只使用统一检测父项，条件、材料、方向、组分和具体物质保留为标准子项。
- 无法可靠归并的项目标记`待确认归属`，不进入横向正式列。
- 新的人工明确规则优先于旧映射种子；不要仅修改输出Excel。
- 新输出使用新文件名，不覆盖历史结果；缓存和旧产物未经确认不得删除。

## 命令速查

```bash
scripts/setup_environment.sh
.venv/bin/python pipeline/qc_all/test_down_items_complete.py
.venv/bin/python pipeline/qc_all/extract_table_items_checkpoint.py --limit 300
QC_DATA_PATH="tmp/<批次>/report_data.json" .venv/bin/python pipeline/qc_all/build_100_sku_parent_child_report.py
.venv/bin/python quality-report-consolidator/scripts/validate_records.py tmp/<批次>/report_data.json --downloads tmp/<批次>/download_manifest.json
node quality-report-consolidator/scripts/export_leader_workbook.mjs tmp/<批次>/report_data.json outputs/test_V2/<新文件名>.xlsx
```

## 深入文档

- `docs/01_业务逻辑.md`：数据范围、颜色、PDF和追溯。
- `docs/02_表格结构与格式.md`：领导版两张业务表、字段顺序和视觉规则。
- `docs/03_检测项目归并规则.md`：同类项、父项/子项和人工确认规则。
- `docs/04_处理流程与运行说明.md`：解析链路、环境变量和复跑命令。
- `docs/05_异常处理与验收标准.md`：真实原因状态与验收基准。
