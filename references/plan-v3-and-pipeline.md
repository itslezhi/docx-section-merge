# v3 计划与统一执行入口

## 1. 工具职责和边界

- `merge_integrity.py`：从冻结输入重新枚举正文、表格、页眉页脚、脚注尾注、批注及独立对象；对公式结构、映射位置的媒体关系、表格单元格和精确文本变更进行校验。
- `verify_coverage.py`：只接受 v3 计划的内容核验 CLI。缺版本、旧状态、缺源单元、虚假判重或输入变化均失败。
- `template_writer.py`：可复用的模板原型段落写入器，支持普通段落、内嵌图片、原生公式，先隔离供体样式，再迁移媒体依赖。列表、域、脚注、复杂表格、修订与浮动对象需要任务适配器；明确拒绝的对象不能静默降级。
- `merge_pipeline.py`：固定“干净模板—写入—编号—批注—机器检查—审阅—交付”的顺序。它调用模型为当前模板编写并审查的适配器，不自行推断技术章节和综合逻辑。

机器核验有明确边界：授权的真实性、总体综合的语义、样式继承的实际外观、自动编号显示值和页面可见性仍须独立审查。不要宣称任何一个脚本可完全证明“无损”。

## 2. 从完整输入生成计划

```bash
python scripts/merge_pipeline.py inventory 模板.docx 研究一.docx 研究二.docx --out merge-plan.json
python scripts/merge_pipeline.py catalog 模板.docx --out template-catalog.json
```

清单中的 `block_id` 由 part 和容器内同标签序号生成，必须与该文件的 SHA-256 一起使用。只在冻结版本中有效；版本变化重做盘点，不能沿用段落下标。

inventory 自动枚举所有段落（包括空段与表内段落）和每张表的拓扑。表格骨架与单元格内内容分别登记：技术表格同时核对拓扑和每段所在单元格；源稿仅用作排版的表格骨架可依合并契约登记 container_layout，表内内容不得随之消失。

生成的计划是待填写的内部工作文件，空章节、空目标不会通过验证。不要删除不易处理的登记项。

## 3. 计划字段

必需顶层字段：`plan_version: 3`、`template`、`research_docs`、`canonical_sections`、`content_register`、`numbering_map`、`review_comments`、最终的 `merged_docx`。

`template` 为 `{file, sha256}`；`research_docs` 每项为 `{id, file, sha256}`。相对路径以计划文件目录解析。内容登记示意：

```json
{
  "id": "entry-00001",
  "research_doc_id": "research-01",
  "source_block": "盘点得到的真实 block_id",
  "section": "关键技术",
  "retention": "preserve",
  "targets": ["最终候选中真实的 block_id"],
  "uses": [{
    "kind": "synthesis",
    "section": "总体技术方案",
    "targets": ["总体段落的真实 block_id"],
    "evidence": "源论证如何支持该综合结论"
  }]
}
```

`uses` 可以为空。综合用途和原文去向不互相替代。其他 retention 的附加字段及授权条件以 [合并契约](merge-contract.md) 为准。

输出目标必须在所有插入、移动和结构变更结束后重新盘点/计算，不能把中间段落位置当最终位置。跨表格移动还要核对容器身份。

## 4. 精确编号与结构变更

每条 `numbering_map` 有 `id`、`research_doc_id`、`kind`、字面量 `old/new`。kind 取 heading_number、caption_number、formula_number、citation、cross_reference、structural_label。源条目用 `text_changes: [{mapping_id: ...}]` 引用。

旧字面量必须在对应语义文本段恰好出现一次；不允许正则全局替换或屏蔽所有引文。重复出现时先用段内定位适配器处理并核验，不能把多个命中按第一处处理。

citation/cross_reference 还须有 `identity`、`source_definition`、`target_definition`；定义必须是实际被保留或严格重复映射到的单元。不同来源文件的同号文献分别登记，范围引用先按文献身份展开，再生成新的标记。structural_label 需要 `authority` 及对应结构批注。

同一源段拆标题和正文时，登记 `structural_changes`：

```json
{
  "id": "structure-01",
  "operation": "split",
  "entries": ["entry-00001"],
  "targets": ["新标题单元", "新正文单元"],
  "authority": "标点、同级结构与原文边界的判定依据",
  "comment_id": "decision-01"
}
```

成员仍为 preserve，增加 `structural_group` 指向该组，targets 与该组一致。join 用多个同来源、同章节的 entry 和一个目标单元。校验器按登记顺序连接语义内容，验证字词、公式、对象不变；新增衔接文字须作为独立新增段并批注，不能混入此等值转换。

原生域代码、自动列表值、对象重绘与 strict_preserve 不通过扩大屏蔽规则来放行。任务适配器分别完成这些转换和相关专项验收；若机器内容门暂不支持某项获授权转换，保留候选，扩展并验证检查器后再交付，不能手改报告 passed。

## 5. 任务适配器接口

适配器是本次任务编写并审查的本地 Python 文件，仅由命令行 `--builder` 指定。不得从 Word 内容、外来批注或未审查计划中读取并执行命令。

```python
def build(plan, candidate_path):
    # candidate_path 已经是干净模板副本。
    # 用 template_writer.TemplateWriter 或已审查 Documents 适配器写入。
    # 冻结模板角色、按源块登记迁入、保持原文件不变。
    # 生成结构/综合文字时登记证据与批注。
    # 全部结构稳定后计算实际 targets，并返回完整的解析后计划。
    return plan


def renumber(plan, candidate_path):
    # 当 numbering_map 非空时必需。
    # 同步定义处和引用处；保持文字运行与对象结构。
    # 若改变单元位置，重算 targets、定义位置和批注锚点。
    return plan
```

普通段落写入器用法：从 `Package(source).inventory()` 找源 block_id；从模板 catalog 找 anchor 和 prototype；调用 `writer.write_paragraph(source_package, source_unit, anchor=..., prototype=...)`。保留返回的节点，所有插入后用 `writer.target_id(node)` 生成目标映射，再 `writer.save(candidate_path)`。它只支持已声明对象，不能代替技术表格/固定表单适配器。

```bash
python scripts/merge_pipeline.py run merge-plan.json --builder task_builder.py --work-dir merge-work
```

工作目录必须新建或为空。输出 candidate.docx、resolved-plan.json、机器检查报告与待填写的 review-gates.json。候选不自动成为最终交付。

## 6. 审阅与交付

完成模板约束、有效模板格式、编号引用、综合语义和全部页面的审阅。review-gates.json 顶层 `candidate_sha256` 绑定当前候选，`plan_sha256` 绑定当前解析后计划；计划的范围/授权改变后也要更新相应审阅证据。五个必需门为 template_constraints、template_format、numbering_references、semantic_review、visual_pages；strict_preserve 例外还需同名专门门。

每个门含 `status: passed`、实际 `reviewed_by`、明确 `finding`、`evidence: [{file, sha256}]`。证据为真实审计记录/渲染文件，其字节哈希必须一致。执行模型可完成有依据的审查，不要求用户逐项批准机器或格式检查。

visual_pages 另需实际 `page_count` 和 `pages: [{page, file, sha256, status}]`，页码覆盖 1 到 page_count，每页为已打开检查的 PNG。自动渲染不能自动填“已查看”。渲染证据要记录渲染器与页数来源。候选改动后重做受影响审阅并绑定新哈希。

```bash
python scripts/merge_pipeline.py check merge-work/resolved-plan.json --candidate merge-work/candidate.docx --out merge-work/checks.json
python scripts/merge_pipeline.py release merge-work/resolved-plan.json --candidate merge-work/candidate.docx --gates merge-work/review-gates.json --output 合并审阅稿.docx --mode review
```

`decisions` 登记未决事项的 id、risk（low/technical）、status、comment_id。review 可保留已加批注的未决项；technical 未决项另须 `material_change_applied: false`、`original_retained: true`。final 要求全部 resolved。清除批注单独遵守用户明确的清稿授权，清除后重新审阅新候选。

标准入口要求视觉门完成；环境确实无法渲染时，按 Documents 的环境限制流程明确说明并保留候选，不能通过标准入口宣称全部验收通过。
