# Textile test classification

Classify by the confidently normalized parent item, not by a child material, direction, condition, component, or analyte. Keep the rule list versioned and allow explicit user or company overrides.

## Functional tests

Match terms including:

- 远红外, 红外升温, 吸湿发热, 发热, 升温, 保温, 蓄热, 暖感;
- 凉感, 接触瞬间凉感;
- 抗菌, 抑菌, 抗病毒, 防螨, 防霉, 消臭;
- 防紫外, 紫外线;
- 防泼水, 拒水, 防水, 耐水压, 静水压, 拒油, 防油, 防污, 易去污;
- 吸湿速干, 速干, 透湿, 透气, 防风;
- 防钻绒, 钻绒;
- 防静电, 抗静电, 负离子;
- 阻燃, 防火, 自清洁, 驱蚊, 防蚊.

## Basic tests

Default a valid confirmed parent item to basic when it does not match a functional term. Typical examples:

- pH, 甲醛, 可分解致癌芳香胺, 邻苯二甲酸酯, 重金属;
- 纤维含量, 成分定性, 羽绒成分测定;
- 耐水/汗渍/摩擦/皂洗/唾液/光色牢度;
- 起毛起球, 缝子纰裂, 断裂/撕破/顶破强力, 耐磨;
- 尺寸变化, 水洗/干洗外观;
- 常规燃烧安全要求 unless the user explicitly treats flame-retardant claims as functional.

## Exclusions

Never classify supplier/company/address, brand, order quantity, product/style/plate number, sample stage/description, dates, report notes, judgment-basis prose, laboratory address, or execution-standard metadata as test items.

Do not classify a row whose parent assignment is `待复核` or `解析残片`; retain it in detail and keep it out of overview counts.

## Overrides

Use a supplied company taxonomy over this default. Apply the newest explicit human decision before older mapping seeds. Record every override, evidence, rule version, and effective date in audit metadata.
