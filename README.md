# Word AI 自动排版 + 模板合并

基于 Python 桥接 DeepSeek 大模型，对 Word 文档段落进行智能分类并自动赋样式，
最后把结果合并进官方模板（`raw_en_20260616.docx` 等），替换 `[SectionName: BodyText]` 节。

---

## 1. 流程总览（对应用户使用 4 步）

```
Step 1  打开目标 docx，手动前处理（删除无关内容 / 分节等）并保存。
          │
Step 2  一键运行 python main.py <前处理后文件.docx>
          │
          ├── 2a LLM 批量分类：把每段文本判成 H1..H9 / BODY / BULLET / NUM / ...
          └── 2b 自动赋样式：按 config.yaml style_map 把标签映射为 Word 自定义样式
          │
Step 3  基于官方模板另存为新文件，清空 BodyText 占位内容，
          把 Step 2 排版后的正文粘入 BodyText 节。
          │
Step 4  后续检查逻辑（占位，待实现，如标题层级连续性 / 表格覆盖率 / 图片题注配对）。
```

输出文件（默认位置）：

| 阶段 | 默认路径 |
|---|---|
| 步骤2 中间产物（排版但未合并） | `output/<原名>_typeset.docx` |
| 步骤3 最终交付（合并模板后） | `output/<原名>_final.docx` |

---

## 2. 环境准备（只需做一次）

### 2.1 Python 版本

要求 Python 3.10+。本机测试通过版本：**Python 3.12.2**。

```cmd
py --version
```

### 2.2 安装依赖

在任意 PowerShell / CMD 中：

```cmd
py -m pip install -r D:\github_repos\word-ai-typeset\requirements.txt
```

依赖清单（与 VBA 无任何耦合，纯 Python）：

```
python-docx >= 1.1.0    # 读写 .docx
requests    >= 2.31.0   # 调 DeepSeek API
PyYAML      >= 6.0      # 解析 config.yaml
```

### 2.3 配置 API Key

打开 `D:\github_repos\word-ai-typeset\key.txt`，把第三行：

```
sk-REPLACE_ME
```

替换成你真实的 DeepSeek API Key（`sk-` 开头的长串），**保存**即可。注释行（`#` 开头）自动忽略。

> ⚠ key.txt 不要提交到 git。建议在 `.gitignore` 中加入 `key.txt`。

### 2.4 确认模板路径

打开 `config.yaml` → `paths.template_docx`，确保指向你机器上真实的模板文件：

```yaml
paths:
  template_docx: "C:\\Users\\PC\\xwechat_files\\...\\raw_en_20260616.docx"
```

模板必须包含 `[SectionName: BodyText]` 标记段落，否则 Step 3 会报错。

---

## 3. 使用方法

### 3.1 完整流程（LLM 排版 + 模板合并）

```cmd
py D:\github_repos\word-ai-typeset\main.py "<前处理后docx路径>"
```

示例：

```cmd
py D:\github_repos\word-ai-typeset\main.py "D:\工作文档\某项目\原文.docx"
```

终端会依次打印：

```
[main] 步骤2：自动排版 ...
[typeset] 共 48 段待分类
[typeset] 完成 -> ...\原文_typeset.docx   (应用 48/48 段)
[main] 步骤3：合并模板 -> ...\原文_final.docx
[merge] 已清空占位内容 27 个块
[merge] 完成 -> ...\原文_final.docx   (插入 50 个块)
[main] 步骤4：后续检查逻辑（占位，待实现）

[main] 全部完成！最终文件: ...\原文_final.docx
```

### 3.2 只合并不排版（跳过 LLM）

当你已经手动赋好样式、或只想验证模板合并链路，加 `--skip-typeset`：

```cmd
py main.py "<已排版文件.docx>" --skip-typeset
```

### 3.3 指定输出路径

默认输出到 `config.yaml` 的 `paths.output_dir`。也可以临时指定：

```cmd
py main.py "原文.docx" --output "D:\交付\XXX项目_最终稿.docx"
```

### 3.4 指定配置文件

支持把 config.yaml 放到别处，用 `--config` 指定：

```cmd
py main.py "原文.docx" --config "D:\工作\my_config.yaml"
```

---

## 4. 配置文件详解

所有变量集中在 `config.yaml`，方便一次性改完。

### 4.1 `paths` —— 文件路径

```yaml
paths:
  template_docx: "..."   # 官方模板路径（必改为本机真实路径）
  input_docx:    "..."   # 仅作默认值，实际以命令行参数为准
  output_dir:    "..."   # 步骤2 + 步骤3 的产物目录
  key_file:      "..."   # API Key 文件路径
```

### 4.2 `template` —— 模板正文节

```yaml
template:
  section_marker: "BodyText"       # 与模板中 [SectionName: Xxx] 的 Xxx 一致
  clear_placeholder: true          # 合并前是否清空 BodyText 节占位（原 VBA 默认 true）
  cn_template: "raw_cn_20260616.docx"
  en_template: "raw_en_20260616.docx"
```

> 说明：当前步骤3 只使用 `paths.template_docx` 指定的单一模板。
> `cn_template / en_template` 保留作后续语种自动选模板扩展用。

### 4.3 `llm` —— DeepSeek 客户端

```yaml
llm:
  base_url: "https://api.deepseek.com/v1"
  model: "deepseek-chat"
  temperature: 0.0        # 分类任务确定性越高越好，建议 0~0.2
  max_tokens: 1024        # 一次批量返回 JSON 所需输出 token 数
  timeout: 30
  batch_size: 20          # 一次请求分类几段，越大越省 token，但单请求越长
```

### 4.4 `style_map` —— LLM 标签 → Word 样式名

**LLM 只能输出以下 18 种白名单标签**，它们一一映射到模板 `raw_en_20260616.docx` 中的自定义样式：

| LLM 标签 | 含义 | 对应模板样式名 |
|---|---|---|
| H1 | 一级标题 | `heading 1` |
| H2 | 二级标题 | `heading 2` |
| H3 | 三级标题 | `heading 3` |
| H4 .. H9 | 四级 ~ 九级标题 | `heading 4` ~ `heading 9` |
| BODY | 普通正文段落 | `_全文正文格式_hs` |
| BULLET | 无序列表项（圆点） | `_项目符号列表1级_圆点_hs` |
| NUM | 有序列表项（数字编号） | `_编号1_(1级)_hs` |
| TABLE_HEADER | 表格表头行 | `_表格_表头标题行_hs` |
| TABLE_BODY | 表格正文行 | `_表格_正文格式_hs` |
| CODE | 代码块 / 等宽字体段 | `_代码字体_hs` |
| IMAGE | 图片所在段 | `_图_hs` |
| CAPTION | 图/表题注 | `_图_题注_hs` |

需要加新标签时，先在 `llm.py` 的 `ALLOWED_LABELS` 集合中加入，再在 `style_map` 中映射到模板的真实样式名。

### 4.5 `typeset` —— 步骤2 行为

```yaml
typeset:
  max_text_len: 500       # 段落超长截断后再送 LLM，避免爆 token
  skip_blank: true        # 跳过空段 / 纯空白段
  fallback_label: "BODY"  # LLM 输出非法或乱码时，统一按 BODY 兜底
```

---

## 5. 完整使用示例

假设工作流是：你收到一份 `D:\工作\某项目原文.docx`。

1. **Step 1（手动，在 Word 里）**：打开 `某项目原文.docx`，删掉无关页、补分节，保存为 `某项目原文_已前处理.docx`。

2. **Step 2+3+4（一键，命令行）**：

   ```cmd
   py D:\github_repos\word-ai-typeset\main.py "D:\工作\某项目原文_已前处理.docx"
   ```

3. **取最终交付物**：打开 `D:\github_repos\word-ai-typeset\output\某项目原文_已前处理_final.docx`，核对目录、页眉页脚、封面、正文样式是否符合预期。

4. **调参**：如某段样式分类不对，可：
   - 改 `config.yaml` 里 `style_map` 映射（如 H2 要对应"标题_3_hs"）；
   - 改 `typeset.fallback_label` 或增大 `llm.temperature` 让分类更"大胆"；
   - 直接在 Word 里手动修正最终稿。

---

## 6. 项目结构

```
word-ai-typeset/
├── config.yaml           # 全局配置（所有可调参数）
├── key.txt               # DeepSeek API Key（sk-...）
├── requirements.txt      # 依赖清单
├── config.py             # 读取 config.yaml / key.txt，返回 Config 对象
├── llm.py                # DeepSeek 客户端：批量段落 → 批量标签
├── typeset.py            # 步骤2：LLM 分类 + 赋样式，输出 _typeset.docx
├── merge.py              # 步骤3：复制模板 → 清空占位 → 粘入正文
├── main.py               # 入口：解析命令行 → 串步骤2、3、4
├── output/               # (运行后生成) 中间 + 最终产物
├── test_input.docx       # (可选) 内置 29 段示例文档，用于快速跑通
└── README.md             # 本文件
```

---

## 7. 常见问题

### 7.1 输出文件里出现重复内容（模板内容出现两次）

**最常见原因**：把上一次的输出 `<原名>_final.docx` 又当成输入再跑了一次 `main.py`。
`main.py` 会把输入文件的 `<base>` 名再拼一个 `_final`，于是输出变成 `<原名>_final_final.docx`，
而 final 文件本身已含完整模板结构，再走一次合并就把模板内容二次插入。

**正确做法**：只用「前处理后的原始 docx」作为输入，不要把 `_final.docx` / `_typeset.docx` 当输入。
如需重跑，先删掉 `output/` 下的旧产物，或用 `--output` 指定一个新名字。

### 7.2 报错 `API Key 文件不存在` / `未在 ... 中找到 API Key`

- 检查 `config.yaml` 的 `paths.key_file` 路径是否存在；
- 打开对应 `key.txt`，确认存在非注释行（不以 `#` 开头）且值为真实 `sk-...`，不是占位。

### 7.2 LLM 请求返回 401 Unauthorized

- Key 填错 / 过期 / 配额已用完；
- 到 DeepSeek 控制台核对 Key 状态，或换一条新 Key；
- 确认 `base_url` 是 `https://api.deepseek.com/v1` （不要带多余尾斜杠）。

### 7.3 返回 429 Too Many Requests

- 短期请求频率超限；`typeset` 段落很多时可把 `llm.batch_size` 调小（如 10）。
- 或等待配额重置 / 升级套餐。

### 7.4 `模板中未找到标记段落: [SectionName: BodyText]`

- 模板文件损坏或被改写；
- 打开模板 docx，搜索是否真的有字面量 `BodyText`；
- 如需改成别的标记（如 `CoverPage`），改 `config.yaml` 的 `template.section_marker` 即可。

### 7.5 最终稿样式不对（段前间距、缩进、字体不对）

- 最常见原因：源 docx 与模板 docx 中**同名样式的 styleId 不一致**。
  `merge.py` 已按"样式名称反向查 styleId"重映射一次；若仍有问题，
  直接用 Word 打开最终稿，找到异常段落 → 重新应用一次对应 `_hs` 样式即可。

### 7.6 控制台中文乱码 / UnicodeEncodeError

- 本脚本已避免输出 emoji；若仍出现乱码，是 Windows CMD/PowerShell 编码为 GBK 导致，
  不影响文件正确性，只需以最终生成的 docx 为准。
- 可选修复：在 CMD 中先跑 `chcp 65001` 切 UTF-8 代码页，再运行脚本。

---

## 8. 待实现 / 扩展点

- **步骤4 检查逻辑**：`main.py` 中已预留占位，可加入：
  - 标题层级连续性（H1 → H2 → H3 正确，无跳级 H1→H3）
  - 表格样式覆盖率（表头是否全是 TABLE_HEADER）
  - 图片 ↔ 题注配对（图片段后面必须跟 CAPTION 段）
  - 术语表刷新（对应 VBA 中的 `更新术语表`）
- **按语种自动切换模板**：根据源文档中文字符占比自动选 cn/en 模板。
- **表格/图片 自动排版**：对表格首行赋 TABLE_HEADER、图片段落赋 IMAGE、下一段赋 CAPTION。

---

## 9. 快速冒烟测试

项目内置了 [test_input.docx](file:///D:/github_repos/word-ai-typeset/test_input.docx)（29 段，含标题/正文/列表/伪代码/3×3表格/题注），用于快速验证：

```cmd
:: 仅验证模板合并（不调 LLM，不消耗 token）
py D:\github_repos\word-ai-typeset\main.py D:\github_repos\word-ai-typeset\test_input.docx --skip-typeset

:: 完整跑一遍（消耗少量 token，需先填好 key）
py D:\github_repos\word-ai-typeset\main.py D:\github_repos\word-ai-typeset\test_input.docx --output D:\github_repos\word-ai-typeset\output\smoke_full.docx
```

---

**文档版本**：2026-08-13 初始版本
