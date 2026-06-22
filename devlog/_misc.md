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

## 2026-06-09 17:45: 清理 import_bills 中的死判重代码
- **文件:** `得物对账单_sqlserver版.py`
- **原因:** `import_bills` 中使用 `check_if_imported` 查询 `dw_dzd_bill_import_records` 表做判重，但该表从未有数据写入，每次全量重处理所有文件。此判重从未生效，是死代码。
- **决策:**
  1. 删除 `import_bills` 中 `check_if_imported` 调用 + `bill_no` 提取（仅用于判重）
  2. 保留 `check_if_imported` 和 `record_import` 函数定义（不碍事，未来可能需要）
  3. 提数步骤幂等覆盖，无副作用
- **影响范围:** 无行为变化，少一次无用数据库查询

## 2026-06-09 18:00: 修复代码审查发现的 5 个问题
- **文件:** `得物对账单_sqlserver版.py`
- **修复:**
  1. **`total_records` 赋值前使用** — fallback_count 日志中引用了尚未定义的 `total_records`，触发降级时 NameError 崩溃。将 `total_records = len(records)` 移到日志之前
  2. **`dw_dzd_bill_overview` ZH 列缺失** — INSERT 中包含 ZH 但表定义没有，通过 ALTER TABLE 补充
  3. **密码日志泄露** — 连接字符串含明文密码写入日志，改为 `safe_conn_str` 脱敏
  4. **`@@IDENTITY`** — 多用户环境可能取到触发器产生的 ID，改为 `SCOPE_IDENTITY()`
  5. **空 bill_no 插入** — bill_no 为空时无警告，添加 `logging.warning`
- **影响范围:** 修复了两个潜在运行时崩溃（#1、#2）和两个安全隐患（#3、#4）

## 2026-06-09 18:15: 提取目录从 dzd_tiqu 改为 dzd_xintiqu
- **文件:** `得物对账单_sqlserver版.py`, `CONTEXT.md`
- **原因:** 新版 _tiqu 文件包含账单总览 sheet 和 bill_no/bill_period 字段，与旧格式不兼容。需保留旧 dzd_tiqu 目录中的历史文件，新流程生成到新目录
- **决策:** 仅改 `EXTRACT_DIR` 常量 `dzd_tiqu` → `dzd_xintiqu`，所有引用自动同步。GUI 启动时自动创建目录
- **操作建议:** 如需补全历史数据，手动点「账单处理」→「账单入库」即可重新处理所有原下载文件到新目录

## 2026-06-09 22:55: 修复自动循环暂停响应和递归问题
- **文件:** `得物对账单_sqlserver版.py`
- **原因:** 自动循环使用 `time.sleep(21600)` 且递归调用自身，导致点击暂停后最多等 6 小时才生效，且调用栈随运行天数增长
- **修复:**
  1. 递归 → `while` 循环，调用栈不再增长
  2. `time.sleep(N)` → `sleep_cancellable(N)`，每 1 秒检查暂停状态，点击暂停立即停止
  3. 异常后不再终止循环，等 60 秒自动重试
- **影响范围:** 自动循环行为改善，不影响手动操作

## 2026-06-09 23:33: 重新打包，程序名优化为 DewuBillSystem
- **文件:** `DewuBillSystem.exe`
- **原因:** 旧包名 `得物对账单_sqlserver 版.exe` 含空格和冗余信息。重新打包清理旧构建，优化命名
- **决策:** `DewuBillSystem.exe`（英文名避免路径编码问题，简洁清晰）
- **打包参数:** `--onefile --windowed`，含全部 hidden-imports

## 2026-06-18 16:59: 修复 exe 运行时报 cacert.pem 找不到导致 HTTPS 连接失败
- **文件:**
  - `DewuBillSystem.spec`
  - `DewuBillSystem.exe`（重新打包）
- **原因:** PyInstaller 打包时未包含 certifi 的 SSL 证书文件 `cacert.pem`，`datas=[]` 为空。运行时报 `OSError: Could not find a suitable TLS CA certificate bundle`，所有得物 API HTTPS 请求均失败
- **决策:**
  1. `.spec` 添加 `import certifi`，`datas=[(certifi.where(), 'certifi')]` 将证书文件打包进 exe
  2. `hiddenimports` 增加 `certifi`，确保模块本身也被打包
  3. 重新打包为 `DewuBillSystem.exe`
- **影响范围:** 仅影响打包后的 exe 运行，源码和开发环境不受影响

## 2026-06-18 17:30: 从源码中移除硬编码数据库密码，敏感配置迁移到 .env
- **文件:**
  - `得物对账单_sqlserver版.py`
  - `requirements.txt`（新增 python-dotenv 依赖）
  - `.env.example`（新建模板文件）
- **原因:** 私有仓库中硬编码数据库默认密码 `yike@2025` 和服务器 IP 不是最佳实践，存在泄露风险
- **决策:**
  1. 添加 `python-dotenv`，启动时自动加载 `.env` 文件
  2. `DB_CONFIG` 全部字段改为从环境变量读取，`DB_PASSWORD` 不设任何默认值
  3. 密码为空时启动即报错退出，避免运行时出现 cryptic 的 pyodbc 错误
  4. 创建 `.env.example` 模板，包含所有配置项占位
- **影响范围:** 需在服务器上创建 `.env` 文件或设置环境变量 `DB_PASSWORD`，否则程序启动时报错退出

## 2026-06-22 17:30: 全量比对 tiqu 文件与 dw_dzd_xs，补入 6 个缺失账单 + 修复 3 个损坏 tiqu 文件
- **文件:**
  - `D:\dw_dzd\dzd_tiqu\奕氪大号\20250811-20250817-QY-1207024550175701_tiqu.xlsx`（重新生成）
  - `D:\dw_dzd\dzd_tiqu\奕氪大号\20250818-20250824-QY-1207024782178141_tiqu.xlsx`（重新生成）
  - `D:\dw_dzd\dzd_tiqu\欧思曼\20260126-20260201-QY-1207031106382011_tiqu.xlsx`（重新生成）
  - `dw_dzd_xs` 表（补入 10,782 条）
  - `dw_dwd_bill_records_copy1` 表（更新 6 条导入记录）
- **原因:** 用户反馈 dw_dzd_xs 表可能缺失数据，全量比对 269 个 tiqu 文件与数据库后发现 6 个账单完全未入库（仅记录了导入状态但数据未写入）
- **决策:**
  1. 全量扫描 269 个 tiqu 文件，提取所有 order_id 与数据库比对
  2. 发现欧阳曼 5 个连续账单（20260504~20260607）+ 奕氪小号 1 个账单（20251222~20251228）共 10,782 条缺失
  3. 清除这 6 个账单的虚假导入记录后重新调用主程序入库函数补数
  4. 3 个 tiqu 文件损坏无法读取，从原始 raw 文件重新生成
  5. 全量验证通过：269 个账单 order_id 100% 匹配
- **影响范围:** dw_dzd_xs 新增 10,782 条记录，dw_dzd_thtk 同步补入退货退款数据。其余数据不受影响
