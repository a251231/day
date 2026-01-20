# 企业微信日报生成器（UI）

## 1. 安装依赖

在项目目录下执行：

`pip install -r requirements.txt`

## 2. 启动UI

`streamlit run wecom_report_ui.py`

注意：不要用 `python wecom_report_ui.py` 直接运行，否则会出现 `missing ScriptRunContext` 且页面打不开。
如果你已经习惯用 Python 方式启动，也可以直接执行：

`python -m streamlit run wecom_report_ui.py`

## 2.1 局域网访问（可选）

如果希望同一局域网其他电脑访问：

- 推荐：`python wecom_report_ui.py` 启动（已默认绑定 `0.0.0.0`）
- 或：`streamlit run wecom_report_ui.py --server.address=0.0.0.0 --server.port=8501`

访问地址：`http://<本机局域网IP>:8501`，并确保防火墙放行 8501 端口。

如需只允许本机访问，可设置环境变量 `WE_COM_REPORT_UI_BIND=127.0.0.1` 后再启动。

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

## 5. 使用 GitHub 工作流打包 exe

仓库已提供 GitHub Actions 工作流：`.github/workflows/build-windows-exe.yml`。

- 打包产物：
  - `daily_wecom_report.exe`：命令行版
  - `wecom_report_ui`：UI 版（文件夹形式，运行其中的 `wecom_report_ui.exe`）
- 触发方式：在 GitHub 页面 `Actions` → `build-windows-exe` → `Run workflow`
- 下载方式：工作流完成后在 `Artifacts` 下载 `windows-exe`

### 体积更小的打包（推荐）

如果你发现 exe 很大，可以使用“最小依赖 venv + 排除项”的工作流：`.github/workflows/build-windows-exe-min.yml`。

- 触发方式：`Actions` → `build-windows-exe-min` → `Run workflow`
- 产物：`Artifacts` → `windows-exe-min`
- 说明：CLI/UI 会在各自独立 venv 内构建，避免把 runner/环境里无关的大依赖打进包里；同时通过 `--exclude-module` 排除常见无关模块。

### 打包 exe “闪退”排查

如果你是双击 `wecom_report_ui.exe` 后一闪而过，常见原因是 **8501 端口被占用** 或 **打包缺少 Streamlit 元数据/静态资源**。

- 建议：用命令行启动查看报错（不会一闪而过）
  - `dist\\wecom_report_ui\\wecom_report_ui.exe`
- 当前打包入口已自动在 `8501~8600` 里选择可用端口，并强制关闭 `global.developmentMode`，避免跳转到 `3000` 端口。
