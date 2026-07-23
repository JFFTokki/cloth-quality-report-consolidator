# Canonical schema

Use English internal keys shown in parentheses when generating JSON. Use the Chinese labels in the workbook.

## First rule

If any report text is traditional Chinese, convert the complete native/OCR text to simplified Chinese before recognizing metadata, headers, items, results, or judgments. Retain the original traditional text in source-row, page-text, or diagnostic evidence fields for audit.

## Detail sheet fields

1. 记录ID (`recordId`)
2. 记录类型 (`recordType`)
3. 来源表行 (`sourceRow`)
4. 报告序号 (`reportIndex`)
5. 报告单号 (`reportNumber`)
6. 检测报告文件名 (`reportFileName`)
7. 报告内编号 (`reportInternalId`)
8. 来源货号/款号 (`sourceProductCode`)
9. 报告签发日期 (`reportIssueDate`)
10. 报告内货号/款号 (`reportProductCode`)
11. 颜色 (`color`)
12. 报告内版单号 (`plateNumber`)
13. 面料/物料编号 (`materialNumber`)
14. 样品类型 (`sampleType`)
15. 来源判定 (`sourceJudgment`)
16. PDF页码 (`pageNumber`)
17. 表序号 (`tableNumber`)
18. 项目序号 (`itemNumber`)
19. 项目类别/表标题 (`tableTitle`)
20. 检测项目 (`rawItem`)
21. 统一检测父项 (`parentItem`)
22. 标准检测子项 (`childItem`)
23. 检测方法/判定依据 (`method`)
24. 技术要求/限值 (`requirement`)
25. 实测/测试结果 (`result`)
26. 单项判定/结论 (`judgment`)
27. 单位 (`unit`)
28. 是否有国际认证 (`hasInternationalCertification`)
29. CMA标识 (`cmaMark`)
30. CNAS标识 (`cnasMark`)
31. CAS号 (`casNumber`)
32. 报告出具机构名称 (`issuingInstitution`)
33. 报告限/检出限 (`detectionLimit`)
34. 备注 (`note`)
35. 项目原始行/页面完整原文 (`rawRow`)
36. 检测报告来源URL (`url`)
37. 简体检测项目 (`simplifiedItem`)
38. 归并状态 (`mappingStatus`)
39. 归并证据 (`mappingEvidence`, AM)
40. 报告签发日期识别状态 (`reportIssueDateStatus`, AN)
41. 报告签发日期异常原因 (`reportIssueDateReason`, AO)
42. CMA识别证据/异常原因 (`cmaRecognitionNote`, AP)
43. CNAS识别证据/异常原因 (`cnasRecognitionNote`, AQ)
44. 是否进入总览 (`includeInOverview`)
45. 检测分类 (`classification`)
46. 本地PDF路径 (`localPath`)
47. 解析器版本 (`parserVersion`)
48. 规则版本 (`ruleVersion`)
49. 诊断信息 (`diagnostics`)
50. 处理状态 (`processStatus`)
51. 来源工作表 (`sourceSheet`)
52. 来源单元格 (`sourceCell`)

Recommended record types: `项目判定汇总`, `检测结果明细`, `补充检测结果`, `表级结论`, `表格说明`, `报告信息/表格原文`, `页面完整原文`, `待复核`, `异常`, `审计摘要`.

Use `mappingStatus` values `已确认`, `自动归并`, `自动新增`, `自动归入父项（子项类型未完全识别）`, `自动归入领域父项`, `待确认归属`, `待复核`, or `解析残片`. Only `已确认`, conservatively verified `自动归并`, `自动新增`, and safe automatic parent assignments may set `includeInOverview=true`; `待确认归属`, `待复核`, and `解析残片` must not enter the overview.

Report-level field rules:

- Extract `reportIssueDate`, `issuingInstitution`, `cmaMark`, and `cnasMark` once per PDF and inherit the same values to all records from that PDF.
- Format a confirmed `reportIssueDate` as `yyyy-mm-dd`. Record the source label, original date text, page, and recognition evidence in `diagnostics`.
- Populate `reportIssueDateStatus` with `已识别`, `未发现`, or `待复核`. Populate `reportIssueDateReason` for every blank or `待复核` date, and also include the selected label/page evidence for recognized dates.
- Use `有`, `未发现`, or `待复核` for `cmaMark` and `cnasMark`. These values describe mark appearance in the report only, not certificate validity or accredited scope.
- Set `hasInternationalCertification=是` when either `cmaMark` or `cnasMark` is `有`; otherwise set it to `否`. This is only a summary of visible marks.
- Populate `cmaRecognitionNote` and `cnasRecognitionNote` with the evidence chain result, such as native text hit, OCR hit, visual mark evidence, top-region image review, negative authorization text, or the concrete reason a mark is `未发现` or `待复核`.
- Preserve a detected CAS value in `casNumber`, including multiple CAS numbers when one result row contains multiple substances. Keep the original source text in `rawItem` or `rawRow`.

## Overview sheet fields

- 货号
- 检测项目总数
- 基础检测项目数
- 基础检测项目清单
- 功能性检测项目数
- 功能性检测项目清单
- 涉及报告数
- 报告单号
- 检测报告文件名
- 对应版单号
- 样品类型
- 总体判定
- 待复核项目数
- 异常报告数

Count and list only distinct confirmed `parentItem` values. Keep every `childItem` and every report-level result in the detail sheet.

## Leader workbook visual contract

Use the supplied leader workbook as the visual template for both sheets:

- hide gridlines and use `TableStyleMedium2` with filter buttons;
- use Carlito 11 pt for body cells, wrapped and top-aligned;
- use a dark-blue `#17365D` title band with 18 pt bold white text;
- use a pale-yellow `#FFF4CE` instruction band with `#7A4E00` text;
- use a blue `#1F4E78` header row with 11 pt bold white centered text;
- use the leader row heights: overview title/instruction/header/data `34/28/38/58`, detail title/instruction/header/data `34/32/44/54`;
- keep the leader blue-white banded rows and its judgment, record-type, and functional-item conditional colors;
- do not freeze panes in the leader-layout export;
- mark fields absent from the leader source workbook with subtle gray `#D9D9D9` headers and light-gray `#F2F2F2` body cells, not orange;
- style the four metadata diagnosis columns `报告签发日期识别状态`, `报告签发日期异常原因`, `CMA识别证据/异常原因`, and `CNAS识别证据/异常原因` with the same subtle gray `#D9D9D9` header and light-gray `#F2F2F2` body cells;
- place `颜色` immediately after `报告内货号/款号` and use the normal leader table colors for that column;
- place `统一检测父项` and `标准检测子项` immediately after `检测项目` and use the normal leader table colors for those columns.
- place `报告签发日期` immediately before `报告内货号/款号`;
- place `是否有国际认证`, `CMA标识`, and `CNAS标识` immediately before `CAS号`, in that order, using normal leader table colors;
- place the four gray diagnosis columns `报告签发日期识别状态`, `报告签发日期异常原因`, `CMA识别证据/异常原因`, and `CNAS识别证据/异常原因` immediately after column AM, occupying AN:AQ;
- place `报告出具机构名称` immediately after `CAS号` and use the normal leader table colors.
- move the secondary trace fields `处理状态`, `来源工作表`, and `来源单元格` to the far right of `质检全项目明细`.

## Audit summary fields

Store these as export metadata or as audit records in `质检全项目明细`; do not create a third business sheet by default.

- parserVersion
- ruleVersion
- sourceSheetCount
- sourceRowCount
- manifestRelationshipCount
- uniqueReportCount
- downloadedCount
- failedCount
- pageCount
- extractedPageCount
- structuredItemCount
- pendingReviewCount
- excludedFragmentCount
- blankItemCount
- unparsedSuspectCount
- exceptionCount
- generatedAt
