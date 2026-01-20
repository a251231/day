# 企业微信日报生成器

从 Excel 生成企业微信日报文本，支持制程/成品两张表，自动回溯最近有数的投料批次并标注。

## 功能特点

- 支持两行/三行拼接表头
- 自动推荐“最新且有关键指标数据”的日期
- 指标回溯最近有数的投料批次并标注
- UI 一键生成 + 命令行批量生成

## 安装依赖

```bash
pip install -r requirements.txt
```

## 启动方式

### UI（推荐）

```bash
streamlit run wecom_report_ui.py
```

如果用 Python 直接启动，可执行：

```bash
python -m streamlit run wecom_report_ui.py
```

#### 局域网访问（可选）

- 方式一：`python wecom_report_ui.py`（默认绑定 `0.0.0.0`）
- 方式二：

```bash
streamlit run wecom_report_ui.py --server.address=0.0.0.0 --server.port=8501
```

访问地址：`http://<本机局域网IP>:8501`，并确保防火墙放行 8501 端口。
如需只允许本机访问，可设置环境变量 `WE_COM_REPORT_UI_BIND=127.0.0.1` 后再启动。

### CLI

```bash
python daily_wecom_report.py --excel <路径> --date YYYY-MM-DD --lookback-days 7 --out report.txt
```

## Docker Compose

在项目根目录执行：

```bash
docker compose up
```

访问：`http://<宿主机IP>:8501`

> Compose 已挂载当前目录，UI 的“选择本地文件”可直接看到宿主机目录下的 Excel。

## 说明文档

- UI 使用说明见 `UI使用说明.md`
