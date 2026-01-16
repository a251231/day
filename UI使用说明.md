# 企业微信日报生成器（UI）

## 1. 安装依赖

在项目目录下执行：

`pip install -r requirements.txt`

## 2. 启动UI

`streamlit run wecom_report_ui.py`

注意：不要用 `python wecom_report_ui.py` 直接运行，否则会出现 `missing ScriptRunContext` 且页面打不开。
如果你已经习惯用 Python 方式启动，也可以直接执行：

`python -m streamlit run wecom_report_ui.py`

## 3. 使用方式

- Excel来源：支持“选择本地文件”或“上传文件”
- 日期：
  - 日期表（默认今天）：从Excel里的“投料日期”列表选择，默认选中今天（如果今天不在表里，则默认最新一天）
  - 自动（推荐）：自动选择“最近且有关键指标数据”的日期
  - 手动日历：按日历手动指定
- 回溯天数：当某个指标当日未出数时，允许向前回溯最近 N 天，找到最近一次有数的投料批次并标注

## 4. 输出说明

- UI输出的是“企业微信消息文本”，可直接复制粘贴
- 若某指标非当日取数，会显示类似：`（2026.01.13投料批次DA2601-080~DA2601-087）`
