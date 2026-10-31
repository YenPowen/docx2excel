# docx → Excel 汇总工具

一个 Windows 桌面小工具：批量选中多个 `.docx` 文档，把每份文档的正文提取出来，汇总进一个 Excel 表格（一行一个文件：文件名 + 文档内容）。

界面用 tkinter 写成，双击即用，不需要命令行。

## 功能

- 多选 Word 文档，可反复追加，自动去重
- 后台线程处理 + 进度条，界面上不会卡死
- 自动生成带时间戳的默认文件名（`docx汇总_20260101_120000.xlsx`）
- 写表时自动设置列宽，「文件名」40 字符、「文档内容」100 字符
- 缺依赖时弹窗提示安装命令，而不是静默闪退
- 双击 `.pyw` 运行，不弹黑色控制台窗口；已做 DPI 感知，高分屏下界面不糊

## 环境要求

- Windows
- Python 3.8+
- 依赖：`python-docx`、`pandas`、`openpyxl`

## 安装

```bash
pip install -r requirements.txt
```

## 运行

双击 `word.pyw`，或者在命令行执行：

```bash
python word.pyw
```

## 使用步骤

1. 点「选择 Word 文档 (多选)」，选中需要汇总的 `.docx`（可以多次点，继续追加）
2. 点「开始汇总到 Excel」，选择保存位置
3. 完成后弹窗提示，生成的 Excel 里「汇总」工作表包含「文件名」「文档内容」两列

## 已知限制

- 只提取正文段落（`python-docx` 的 `paragraphs`）。表格、页眉页脚、图片里的文字不会提取，需要的话可以扩展 `extract_text_from_docx`
- 输出是纯文本拼接，原文档的字体、样式、编号不会保留
- 单个文件读取失败不会中断整体流程，会在内容列写入 `[读取失败: ...]`

## 项目结构

```
word.pyw          主程序：tkinter 界面 + docx 提取 + Excel 写出
requirements.txt  依赖清单
```

## 核心代码

```python
doc = docx.Document(str(docx_path))
paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
return "\n".join(paragraphs)
```
