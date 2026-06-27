# 得物对账单自动获取系统

多店铺并发下载得物平台对账单，自动提取关键数据并入库到 SQL Server。

## 功能

- **多店铺并发** — 6 家店铺并行下载，互不阻塞
- **账单下载** — 自动获取各周期对账单 Excel
- **数据提纯** — 从原始 Excel 提取关键列，生成精简文件
- **自动入库** — 解析提纯文件，写入 SQL Server 数据库
- **定时调度** — 可设置每日自动运行时间
- **GUI 界面** — PyQt6 图形界面，多标签页日志查看

## 项目结构

```
得物对账单_sqlserver版.py       # 主程序（GUI 版）
得物对账单_历史数据版.py          # 历史数据批量获取（无头模式）
.env.example                    # 配置文件模板
requirements.txt                # Python 依赖
```

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 配置

复制 `.env.example` 为 `.env`，填写数据库和 API 配置：

```ini
# SQL Server 数据库
DB_SERVER=your_server
DB_DATABASE=your_database
DB_USERNAME=your_username
DB_PORT=1433
DB_PASSWORD=your_password

# 得物开放平台 API
APP_KEY=your_app_key
APP_SECRET=your_app_secret

# 店铺凭证（cred_id, app_key, app_secret, shop_name）
```

### 3. 运行

```bash
python 得物对账单_sqlserver版.py
```

### 4. 打包 exe

```bash
pip install pyinstaller
pyinstaller 得物对账单_sqlserver版.spec
```

## 工作流程

```
下载对账单 → 提纯关键列 → 写入 SQL Server
   (Excel)      (_tiqu.xlsx)      (dw_dzd_* 表)
```

## 数据库表

- `dw_dzd_bill_overview` — 账单总览（汇总金额）
- `dw_dzd_bill_records` — 账单记录（每笔明细）
- `dw_dzd_xs` — 销售订单
- `dw_dzd_thtk` — 退货退款

## 技术栈

- Python 3.10+
- PyQt6（GUI）
- PyODBC + SQL Server
- Pandas + OpenPyXL（Excel 处理）
- Requests（API 调用）
- Tenacity（自动重试）

## 许可

MIT
