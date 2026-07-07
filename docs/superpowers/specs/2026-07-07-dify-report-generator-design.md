# Dify AI 应用 — 报告模板、字段映射与 Word 导出 设计文档

## 概述

在 Dify AI 应用中实现报告生成功能，支持用户上传 .docx 模板（含占位符）、通过 Dify 工作流填充数据、最终导出为 Word 文档。

## 方案选择

**方案二：Dify 工作流内实现**

完全在 Dify 工作流内完成，利用 Dify 的 LLM 节点 + Code Node 实现模板填充和 Word 导出。

## 完整工作流节点编排

```
[开始] ──▶ [LLM 节点: 报告分析] ──▶ [Code Node: 模板填充] ──▶ [结束]
```

---

## 节点 1：开始节点 — 变量定义

配置应用的所有输入变量，用户在工作流启动时填写/上传。

### 变量列表

| 变量名 | 类型 | 是否必填 | 说明 |
|--------|------|---------|------|
| `report_template` | File | 是 | 用户上传的 .docx 模板文件 |
| `customer_name` | TextInput | 是 | 客户/主体名称 |
| `eval_date` | TextInput | 是 | 评估日期 |
| `photo` | File | 否 | 现场照片（可选） |

### 开始节点配置截图说明

```
开始节点
├── 输入表单
│   ├── [文件上传] 报告模板 *      ← report_template (File)
│   ├── [文本框]   客户/主体名称 *  ← customer_name (TextInput)
│   ├── [文本框]   评估日期 *       ← eval_date (TextInput)
│   └── [文件上传] 现场照片        ← photo (File)
│
└── 输出变量 (供下游节点使用)
    ├── {{input.report_template}}   ← 模板文件引用
    ├── {{input.customer_name}}     ← 客户名称值
    ├── {{input.eval_date}}         ← 评估日期值
    └── {{input.photo}}            ← 图片文件引用
```

---

## 节点 2：LLM 节点 — 报告内容生成

### 基础配置

| 配置项 | 值 |
|--------|-----|
| **节点名称** | 报告内容生成 |
| **模型** | gpt-4o / claude-3.5-sonnet（推荐，JSON 输出质量高） |
| **温度** | 0.2（控制一致性，确保 JSON 格式稳定） |
| **输出格式** | JSON（启用 Dify JSON 输出模式） |

### 上下文 / 变量引用

| 变量引用 | 在 Prompt 中的占位符 |
|----------|---------------------|
| `input.customer_name` | `{{input.customer_name}}` |
| `input.eval_date` | `{{input.eval_date}}` |

### System Prompt

```
你是一名专业的报告分析师。
根据客户信息生成评估报告内容。
输出必须严格遵循指定的 JSON 格式。
```

### User Prompt

```
请为以下客户生成评估报告内容：

客户名称: {{input.customer_name}}
评估日期: {{input.eval_date}}

请严格按照以下 JSON 格式输出，不要包含任何其他内容：
{
  "score": <数字, 0-100>,
  "conclusion": "<评估结论文本，可包含多个段落，段落之间用 \\n 分隔>",
  "eval_table": [
    {"dimension": "<维度名称>", "score": <分数>, "note": "<说明>"}
  ],
  "suggestion": "<改进建议>"
}

要求：
- eval_table 数组应包含 3-5 个评估维度
- conclusion 应详细、具体
- 输出必须是合法的 JSON，不可包含注释或额外文字
```

### 输出变量

| 输出变量 | 类型 | 说明 |
|----------|------|------|
| `llm_output.text` | String | LLM 原始输出文本（JSON 字符串） |

> **注意**: 如果 Dify LLM 节点支持 JSON 模式，可以直接解析为结构化变量。如果不支持，则 Code Node 中做 `json.loads()` 解析。

---

## 节点 3：Code Node — 模板填充与 Word 生成

### 基础配置

| 配置项 | 值 |
|--------|-----|
| **节点名称** | 模板填充引擎 |
| **运行环境** | Python 3 |
| **超时时间** | 60 秒 |

### 输入变量

| 变量名 | 来源 | 类型 | 说明 |
|--------|------|------|------|
| `template_file` | `{{input.report_template}}` | File | 用户上传的 .docx 模板 |
| `llm_raw_output` | `{{llm_output.text}}` | String | LLM 输出的 JSON 字符串 |
| `customer_name` | `{{input.customer_name}}` | String | 客户名称 |
| `eval_date` | `{{input.eval_date}}` | String | 评估日期 |
| `photo_file` | `{{input.photo}}` | File | 现场照片（可选） |

### 输出变量

| 输出变量 | 类型 | 说明 |
|----------|------|------|
| `word_output` | File | 生成的 .docx 报告文件 |

### 完整 Python 代码

```python
import json
import re
import os
from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement


def main(template_file: str, llm_raw_output: str,
         customer_name: str, eval_date: str,
         photo_file: str = None) -> dict:
    """
    Dify Code Node 入口函数。

    参数:
        template_file: 上传的 .docx 模板文件路径 (Dify 自动传入)
        llm_raw_output: LLM 节点输出的 JSON 字符串
        customer_name: 客户名称
        eval_date: 评估日期
        photo_file: 现场照片文件路径 (Dify 自动传入，可选)
    返回:
        dict: {"word_output": "/tmp/generated_report.docx"}
    """
    # ============================================================
    # 1. 解析 LLM 输出
    # ============================================================
    try:
        llm_data = json.loads(llm_raw_output)
    except json.JSONDecodeError:
        # 如果 LLM 输出不是标准 JSON，尝试从中提取 JSON 块
        json_match = re.search(r'\{.*\}', llm_raw_output, re.DOTALL)
        if json_match:
            llm_data = json.loads(json_match.group())
        else:
            raise ValueError("LLM 输出无法解析为 JSON")

    # ============================================================
    # 2. 构建扁平化的变量字典
    # ============================================================
    # input.xxx 开头的变量
    input_vars = {
        "input.customer_name": customer_name,
        "input.eval_date": eval_date,
    }

    # llm_output.xxx 开头的变量 (从 LLM JSON 中提取)
    llm_vars = {}
    for key, value in llm_data.items():
        llm_vars[f"llm_output.{key}"] = value

    # 合并所有变量
    all_vars = {**input_vars, **llm_vars}

    # ============================================================
    # 3. 加载模板并处理
    # ============================================================
    doc = Document(template_file)

    # 3.1 处理所有段落
    _process_all_paragraphs(doc, all_vars, photo_file)

    # 3.2 处理所有表格 (模板中已有的表格)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _process_single_paragraph(para, all_vars, photo_file)

    # ============================================================
    # 4. 保存输出
    # ============================================================
    output_path = '/tmp/generated_report.docx'
    doc.save(output_path)

    return {'word_output': output_path}


def _process_all_paragraphs(doc, all_vars, photo_file):
    """遍历文档所有段落，处理占位符"""
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # 检查是否包含图片占位符
        if '{{image:' in para.text:
            _handle_image_placeholder(para, photo_file)
            continue

        # 检查普通占位符
        for key, value in all_vars.items():
            placeholder = '{{' + key + '}}'
            if placeholder not in para.text:
                continue

            if isinstance(value, list) and len(value) > 0:
                # 数组值 → 在段落后插入 Word 表格
                _insert_table_after_paragraph(doc, para, value)
                # 清空占位符文本
                for run in para.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, '')
            elif isinstance(value, str) and '\n' in value:
                # 含换行的文本 → 替换为多段
                _replace_with_multi_paragraph(para, placeholder, value)
            else:
                # 普通文本/数字 → 直接替换
                for run in para.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, str(value))

    # 清理空段落（表格插入后原段落可能只剩空文本）
    _cleanup_empty_paragraphs(doc)


def _process_single_paragraph(para, all_vars, photo_file):
    """处理单个段落中的占位符替换（用于表格内单元格）"""
    for key, value in all_vars.items():
        placeholder = '{{' + key + '}}'
        if placeholder not in para.text:
            continue
        for run in para.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, str(value))


def _insert_table_after_paragraph(doc, para, data_list):
    """
    在段落后插入一个 Word 表格。
    data_list: [{"dimension": "...", "score": ..., "note": "..."}, ...]
    """
    if not data_list or not isinstance(data_list, list):
        return

    # 从第一个元素提取表头
    headers = list(data_list[0].keys())

    # 获取段落所在的父级（body）
    parent = para._element.getparent()
    # 获取段落的索引位置
    para_index = list(parent).index(para._element)

    # 创建表格
    rows_count = len(data_list) + 1  # +1 表头行
    cols_count = len(headers)
    table = doc.add_table(rows=rows_count, cols=cols_count)
    table.style = 'Light Grid Accent 1'

    # 填充表头
    for col_idx, header in enumerate(headers):
        cell = table.rows[0].cells[col_idx]
        cell.text = header
        # 加粗表头
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True

    # 填充数据行
    for row_idx, item in enumerate(data_list):
        for col_idx, header in enumerate(headers):
            cell = table.rows[row_idx + 1].cells[col_idx]
            cell.text = str(item.get(header, ''))

    # 将表格插入到段落之后
    parent.insert(para_index + 1, table._element)


def _handle_image_placeholder(para, photo_file):
    """处理图片占位符 {{image:xxx}} — 将占位符替换为图片"""
    if not photo_file or not os.path.exists(photo_file):
        # 没有图片文件，替换为占位文本
        for run in para.runs:
            run.text = re.sub(r'\{\{image:\w+\}\}', '[图片未提供]', run.text)
        return

    # 清空段落文本
    for run in para.runs:
        run.text = ''

    # 在段落中插入图片
    run = para.add_run()
    try:
        run.add_picture(photo_file, width=Inches(5))
    except Exception:
        # 图片插入失败，留空
        run.text = ''


def _replace_with_multi_paragraph(para, placeholder, text):
    """将含换行的文本替换为多个段落"""
    lines = text.split('\n')

    if not lines:
        return

    # 第一行替换原段落中的占位符
    for run in para.runs:
        if placeholder in run.text:
            run.text = run.text.replace(placeholder, lines[0])

    # 后续行作为新段落插入
    if len(lines) > 1:
        parent = para._element.getparent()
        para_index = list(parent).index(para._element)

        for line in lines[1:]:
            new_para = OxmlElement('w:p')
            new_run = OxmlElement('w:r')
            new_text = OxmlElement('w:t')
            new_text.text = line
            new_run.append(new_text)
            new_para.append(new_run)
            para_index += 1
            parent.insert(para_index, new_para)


def _cleanup_empty_paragraphs(doc):
    """删除内容为空的段落"""
    for para in doc.paragraphs:
        if para.text.strip() == '' and len(para.runs) == 0:
            para._element.getparent().remove(para._element)



```

---

## 节点 4：结束节点 — 输出配置

### 输出配置

| 配置项 | 值 |
|--------|-----|
| **输出方式** | 直接返回文件供下载 |
| **输出变量** | `{{code_node.word_output}}` |

### 用户端展示效果

用户在工作流启动后会看到：

```
┌─────────────────────────────────────┐
│  报告生成                            │
│                                     │
│  [上传] 报告模板 (.docx) *           │
│  [输入] 客户/主体名称 *              │
│  [输入] 评估日期 *                   │
│  [上传] 现场照片 (可选)              │
│                                     │
│  [▶ 开始生成]                        │
└─────────────────────────────────────┘

生成完成后:
  ↓ 点击下载 "generated_report.docx"
```

---

## 模板设计 (.docx 格式)

用户在 .docx 模板中使用 Dify 变量语法标记占位符：

| 变量语法 | 对应变量路径 | 说明 |
|----------|-------------|------|
| `{{input.customer_name}}` | 开始节点 → customer_name | 文本 |
| `{{input.eval_date}}` | 开始节点 → eval_date | 文本 |
| `{{llm_output.score}}` | LLM 节点输出 → score | 数字 |
| `{{llm_output.conclusion}}` | LLM 节点输出 → conclusion | 段落文本（可含 \n 换行） |
| `{{llm_output.eval_table}}` | LLM 节点输出 → eval_table | 数据表格（值为数组） |
| `{{llm_output.suggestion}}` | LLM 节点输出 → suggestion | 文本 |
| `{{image:input.photo}}` | 开始节点 → photo | 图片（需在模板中有图片占位符） |

### 模板示例

```
客户评估报告
═══════════════

客户名称：{{input.customer_name}}
评估日期：{{input.eval_date}}
综合评分：{{llm_output.score}} 分

评估结论：
{{llm_output.conclusion}}

评估明细：
{{llm_output.eval_table}}

改进建议：
{{llm_output.suggestion}}

现场照片：
{{image:input.photo}}
```

---

## 模板填充规则总结

| 数据类型 | 模板写法 | Code Node 行为 |
|----------|---------|---------------|
| 文本 | `{{input.customer_name}}` | 字符串直接替换 |
| 数字 | `{{llm_output.score}}` | 转字符串后替换 |
| 多段文本 | `{{llm_output.conclusion}}` | 按 `\n` 拆分为多个段落 |
| 数组/表格 | `{{llm_output.eval_table}}` | 在占位符位置插入 Word 表格（第一行表头） |
| 图片 | `{{image:input.photo}}` | 插入图片文件 |

---

## 注意事项

1. **python-docx 限制**: 复杂样式的 .docx 模板在替换后可能丢失部分格式（如字体颜色、缩进等），需要在真实模板上测试验证
2. **表格样式**: 代码中使用了 `Light Grid Accent 1` 样式，可替换为 Dify 环境中可用的其他表格样式
3. **文件大小**: Dify Code Node 对输出文件有大小限制，大文件建议控制内容量
4. **图片**: 需要 Dify 文件系统支持通过文件 ID 读取图片，如果 `photo_file` 为空则跳过
5. **占位符命名**: 模板中的占位符必须与 Dify 变量路径完全匹配，区分大小写
6. **LLM JSON 输出稳定性**: 建议使用 gpt-4o 或 claude-3.5-sonnet 并设置 temperature=0.2，确保 JSON 格式稳定
