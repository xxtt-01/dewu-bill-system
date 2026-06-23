# 得物对账单系统 (Dewu Bill System)

店铺对账单自动获取与入库系统，对接得物开放平台 API，将账单 Excel 数据解析并写入 SQL Server 数据库。

## Glossary

### 核心概念
| 术语 | 定义 |
|------|------|
| **得物 (Dewu)** | 电商平台，提供开放 API 接口获取店铺对账单 |
| **对账单 (Bill)** | 店铺在得物平台的结算账单，包含周期内的所有订单交易数据 |
| **Bill No** | 账单编号，格式如 `20250630-20250706-QY-1207023276910521`，包含日期范围+店铺标识 |
| **店铺 (Shop)** | 得物平台上的店铺，系统监控 5 家店铺：北京晶睿、锋潮、欧思曼、奕氪大号、奕氪小号 |
| **应用凭证 (AppCredential)** | 得物开放平台的 API 凭证，包含 cred_id/app_key/app_secret，存储在数据库 `dewu_app_credentials` 表 |

### 数据表
| 表名 | 用途 |
|------|------|
| `dewu_app_credentials` | 得物开放平台 API 凭证（店铺标识 + Key + Secret） |
| `dw_dwd_bill_records` | 【API账单列表】存储 period_list 接口获取的账单记录，按 bill_no+店铺去重 |
| `dw_dwd_bill_records_copy1` | 【已入库判重】存储已成功导入的 bill_no，process_import 前按此表跳过 |
| `dw_dzd_xs` | 销售订单明细表：存储所有「销售订单」sheet 的行级数据 |
| `dw_dzd_thtk` | 退货退款订单明细表：存储所有「退货退款订单」sheet 的行级数据 |
| `dw_dzd_bill_overview` | 账单总览表：每账单一行，存储「账单总览」sheet 的汇总数据（23 个字段） |
| `dw_dzd_other_fee` | 本期结算其他项费用明细表（20 个金额/类型字段 + bill_no/bill_period） |
| `dw_dzd_deduction_detail` | 扣减其他费用明细表（费用类型、金额、币种） |
| `dw_dzd_cargo_damage` | 本期货损买进订单明细表（19 个订单/金额字段） |

### 数据流阶段
| 阶段 | 操作 | 输入 | 输出 |
|------|------|------|------|
| **① 下载账单** | API 拉取账单列表 → 存入 `dw_dwd_bill_records` → 下载 Excel | 得物 API | `downloads/{shop}/{bill_no}.xlsx` |
| **② 账单处理(提数)** | 提取 Excel 中全部 6 个 sheet（账单总览/销售订单/退货退款订单/本期结算其他项费用/扣减其他费用明细/本期货损买进订单） | 原始 xlsx | `tiqu/{shop}/{bill_no}_tiqu.xlsx` |
| **③ 账单入库** | 解析 tiqu 文件 → 写入 6 张明细表 | tiqu xlsx | 数据库记录 |

### 数据表字段结构

#### 账单总览表 `dw_dzd_bill_overview`
存储「账单总览」sheet 中 Row 4 的概要数据（每账单一行）：
- `bill_no` — 账单编号（按此字段去重）
- `company_name` — 公司名称
- `settlement_cycle` — 结算周期（周结）
- `bill_period` — 账单起止时间
- `product_transaction_amount` — 本期商品交易金额
- `platform_service_fee` — 本期交易类平台服务费金额
- `distribution_service_fee` — 分销服务费
- `advance_payment_recovery` / `seller_subsidy_amount` / `adjustment_item` 等
- `other_*` — 其他非交易类费用（来自「本期结算其他项费用」子列）
- `price_reduction_refund` / `price_reduction_subsidy` — 售中降价
- `total_payable_amount` — 应结总金额
- `settlement_status` — 结算状态
- `create_time` — 入库时间

#### 销售订单表 `dw_dzd_xs` 字段来源
多行表头拼接格式 `{分组}{字段名}{字段名}`，含 `bill_no`、`bill_period` 两个新增字段。

#### 退货退款订单表 `dw_dzd_thtk` 字段来源
同拼接结构，含 `bill_no`、`bill_period` 两个新增字段。

### 需要关注的点
- `import_bills`（阶段②）按 sheet 类型分 3 种表头处理：
  - **销售订单/退货退款订单**：`data.iloc[1:]` 去说明行 → `data.iloc[:3]` 合并 3 行（分组+列名+子列名）为扁平表头 → `data.iloc[3:]` 取数据
  - **账单总览**：`data.iloc[1:]` → 合并前 2 行为扁平表头 → `data.iloc[2:]` 取数据
  - **本期结算其他项费用/扣减其他费用明细/本期货损买进订单**：`data.iloc[1:]` 去说明行后第 1 行即列名，无需合并，直接写入
- `process_import`（阶段③）先检查 `dw_dwd_bill_records_copy1` 判重，已入库则跳过
- SQL Server 连接使用 `pyodbc` + `{SQL Server}` 驱动，密码通过环境变量 `DB_PASSWORD` 配置
- 批量插入每 500 条提交一次，避免长事务超时
- 已废弃的 `dw_dzd_bill_import_records` 表及其对应的 `check_if_imported`、`record_import` 函数已在 2026-06-23 清理
