## 2026-06-01 17:35: 修复 5 个潜在问题
- **文件:** `得物对账单_sqlserver版.py`
- **原因:** 代码审查发现 5 个潜在问题，逐一修复
- **决策:**
  1. **密码硬编码** — 改为 `os.environ.get('DB_PASSWORD', 默认值)`，部署时设环境变量即可覆盖，不破坏现有运行
  2. **`record_import` 占位符** — `%s` → `?`，与其他 SQL Server 查询保持一致（该函数虽然未调用，但修复以消除隐患）
  3. **倒计时标签遮挡** — `countdown_label` 从 `row=2, column=1` 移到 `row=3, column=0`，`pause_button` 移到 `row=3, column=1`，行列对齐不再重叠
  4. **`auto_run` 动态类型** — `type('AutoRun', ...)` → 正式 `class AutoRun` 定义，消除每次 `main_gui()` 调用创建新匿名类的泄漏
  5. **窗口关闭线程终止** — 添加 `WM_DELETE_WINDOW` 协议处理器，关闭前暂停自动运行并确认退出
- **影响范围:** 仅 `得物对账单_sqlserver版.py`

## 2026-06-01 18:30: 修复 OSError 日志缺失 + PyInstaller 重新打包
- **文件:** `得物对账单_sqlserver版.py`
- **原因:** fetch_api_data 的 except 漏捕 OSError，导致无法诊断得物 API 连接失败
- **决策:**
  1. 增加 `except OSError` 分支，记录 `[type(e).__name__]` + 完整 traceback
  2. 用 PyInstaller 重新打包，输出 `dist/得物对账单_sqlserver 版.exe`
- **提交:** 1f322f6（OSError 日志修复），8dd6cf0（delivery_time 映射修正）
- **重新打包:** 已重建 dist/得物对账单_sqlserver 版.exe（132MB，含上述修复）

## 2026-06-09 16:00: 新增账单总览提取入库 + 扩展字段
- **文件:**
  - `得物对账单_sqlserver版.py`
  - `数据库迁移_20260609_账单总览.sql`（新建）
  - `CONTEXT.md`
- **原因:** 需要在账单处理流程中提取「账单总览」sheet 并存储到数据库；同时在销售/退款订单表中记录所属账单信息
- **决策:**
  1. **新建表 `dw_dzd_bill_overview`** — 存储每份账单的汇总概要（23 个字段），用 `sp_addextendedproperty` 加中文注释
  2. **扩展 `dw_dzd_xs` / `dw_dzd_thtk`** — 各追加 `bill_no`（账单编号）和 `bill_period`（账单起止时间），方便查询时知道订单属于哪个账单
  3. **账单总览表按 `bill_no` 去重** — 同一账单多次入库不会重复插入
  4. **提数阶段（import_bills）** — `SHEETS_TO_KEEP` 增加「账单总览」，用与销售订单一致的 2 行表头拼接逻辑处理
  5. **入库阶段（process_import）** — `bill_period` 从账单总览 sheet 提取，传递给销售/退款订单的导入函数
- **影响范围:** 已有数据不受影响（新增字段有默认空值），新处理的账单会完整填充

## 2026-06-09 17:00: 发货时间自动降级使用业务时间
- **文件:** `得物对账单_sqlserver版.py`
- **原因:** 得物 Excel 表格曾调整格式，部分周期「发货时间」字段为空但有「业务时间」字段；另一部分没有业务时间但发货时间正常。需自动判断并降级
- **决策:**
  1. 动态检测 `订单基础信息业务时间业务时间` 列是否存在于当前 Excel 中
  2. `delivery_pos` 通过 `field_mapping.values().index('delivery_time')` 动态计算
  3. 降级条件：`delivery_time 为空 AND 业务时间列存在 AND 业务时间有值`
  4. 销售订单和退货退款订单同步增加此逻辑
- **影响范围:** 历史已入库的发货时间空值不会自动补填，后续入库会自动降级

## 2026-06-09 17:30: 代码审查修复 4 项问题
- **文件:** `得物对账单_sqlserver版.py`
- **原因:** 全面代码审查发现若干潜在问题
- **修复:**
  1. **`extract_bill_period_from_file` 硬编码列索引** — 改为用列名(`'账单起止时间账单起止时间'.get()`)取值，避免得物调整列顺序后取错值；异常时记录 warning 而不是静默吞掉
  2. **`import_bills` 中 bill_no 截断** — `split('-')[0]` 只能取到日期前缀，改为 `.replace('.xlsx', '')` 取完整编号
  3. **降级计数器** — 在销售订单和退款订单中记录实际触发了多少行的发货时间降级，便于日志排查
- **影响范围:** 仅代码健壮性提升，业务逻辑不变
