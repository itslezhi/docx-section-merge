---
name: docx-section-merge
description: 以用户指定的一份 DOCX 项目模板为母版，将多份子文档逐章合并，保留正文与技术对象，统一模板格式和编号，并用真实 Word 批注记录模型判断。用于任务书、建议书、申报书的多稿合并；不用于无模板的整份文档从零编写。
---

# 项目任务书多子文档合并

以一份项目模板和一份或多份子文档 Word 生成新的合稿。模板提供章节、格式和填写要求，子文档提供技术事实与内容。保持原文件不变；默认完整保留非重复内容。用户已明确授权的范围、顺序、删除、综合和格式例外在本次会话内持续有效。

**命名约定。** 泛指输入文件类型时统一称“子文档”；当示例、批注来源或映射记录需要区分多份输入时，统一使用“子文档A、子文档B、子文档C……”命名。

## 核心规则

每次先读 [合并契约](references/merge-contract.md)。它是本 Skill 内授权、内容处置与异常处理的唯一规则来源；各专项文件只解释实现，不能扩大权限。

- 使用当前模板的角色原型写入；普通正文、固定表单和技术对象走各自的写入路径，禁止导入供体全局样式。
- 从每份原始 DOCX 自动枚举所有内容单元，再逐章登记去向。不得先挑出“重要内容”，再用挑选结果证明完整覆盖。
- 只有重新计算后确证完全重复，或用户明确授权，才能删除实质内容；相近、较短或被长稿包含均不足以删除。
- `retention` 单独记录原文去向，`uses` 单独记录综合用途。原文完整保留与作为总体综合依据可以同时成立。
- 模型补标题、加衔接、综合、改变语义位置时生成真实 Word 批注。低风险歧义可继续生成审阅稿；涉及技术事实或承诺的未决变动保持原文，只暂停受影响项。
- 图片必须核对所在位置与实际媒体的对应关系；公式必须核对结构；编号必须按来源和对象身份精确映射，禁止全局数字替换或以屏蔽全部编号代替核验。
- 新任务只用 v3 计划。输入、输出和审阅证据绑定文件哈希；缺少计划版本、源块、证据或必需验收项时不能判通过。

## 执行流程

1. **盘点。** 识别唯一模板与全部子文档，冻结输入哈希。用 `merge_pipeline.py inventory` 从源文件生成完整登记表，用 `catalog` 查询模板/输出单元；识别模板全部说明并逐页核对。子文档含多项研究内容时按实际章节映射，不按文件数推断研究项数。
2. **确定计划。** 用模板建立章节与格式原型；补全各源单元的 `retention`、目标章节及独立的 `uses`。登记授权、总体综合范围、编号对象身份和批注；不得通过删登记项绕过内容保留。
3. **受控生成。** 使用统一入口 `merge_pipeline.py run`，从干净模板建立候选，依次调用任务适配器的内容写入、编号、批注及机器校验。普通段落、内嵌图片和原生公式可复用 `TemplateWriter`；复杂对象由已审查的 Documents 适配器处理，仍走同一验收入口。
4. **审阅。** 机器通过后完成模板约束、有效格式、编号引用、语义和逐页视觉检查。保留问题位置的批注；检查实际目标编辑器可用性，不能把未完成的检查写成通过。
5. **交付。** 用 `release --mode review` 交付审阅稿；只有未决裁定全部解决时才用 `--mode final`。交付级别与批注是否清除分别控制，“最终文件”不等于清除批注。默认只交付新的 DOCX。

命令参数、完整计划格式及适配器接口见 [v3 计划与统一入口](references/plan-v3-and-pipeline.md)。不要要求用户填写内部 JSON。

## 按需读取

| 触发条件 | 文件 / Skill |
| --- | --- |
| 每次合并最先读取 | [合并契约](references/merge-contract.md)、[模板说明约束](references/template-instruction-contract.md)、[v3 计划与统一入口](references/plan-v3-and-pipeline.md) |
| 逐稿逐章映射及内容处置 | [章节映射](references/research-doc-section-mapping.md)、[内容保留](references/content-reconciliation.md) |
| 向模板写入内容 | [样式隔离](references/style-isolation-pipeline.md)、[格式规范化](references/format-normalization.md) |
| 总体章节综合 | [总体综合](references/synthesis-sections.md)；需要撰写时调用 `research-project-document-v2`，限定在已授权槽位 |
| 新绘、改绘、修复图件 | [图件处理](references/diagram-generation-repair.md)、`minimal-card-flowchart`；仅在图件工作触发时加载 |
| 模型新增、拆合、衔接、重定位或歧义 | [批注模式](references/review-comments-mode.md) |
| 长表、定高容器或跨页 | [分页](references/table-pagination.md) |
| 标题、列表、图表、公式或引文编号变化 | [编号引用](references/renumbering.md) |
| 复杂对象迁移 | [OOXML 移植](references/ooxml-transplant.md) |
| 交付前 | [交付 QA](references/qa.md) |

实际编辑 DOCX 时使用 `documents` Skill，按对象读取 `tasks/multi_doc_merge.md`、`tasks/style_lint_normalize.md`、`tasks/headings_numbering.md`、`tasks/images_figures.md`、`tasks/captions_crossrefs.md`、`tasks/fields_update.md`、`tasks/verify_render.md`；批注读取 comments 工作流。已有母版时不套用 Documents 的从零创作预设，不改变母版纸张、表格或字体来适配通用设计。

## 完成标准与边界

所有源单元有去向、每项删除有依据、内容/对象/编号核验通过、模板格式通过、批注闭合且逐页可见，才完成交付。机器通过仅表示对应机器检查完成；人工语义判断与逐页视觉检查必须单独有证据。

模板不唯一、文件无法读取或缺失必要源文件时暂停依赖它们的工作。局部对象暂不支持时保留输入和已通过检查的候选，说明具体缺项并继续独立部分；不得移除该对象后声称合稿完整。详细异常分级见合并契约。

更新本 Skill 的脚本后运行 `python scripts/run_regression.py`；修改用法或核心处置规则后，另用隔离的小型任务检验完整流程。
