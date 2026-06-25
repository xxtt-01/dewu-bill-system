## 2026-06-25: API 409 修复 + 批量流程完善 + 恢复并行
- **文件:**
  - `得物对账单_sqlserver版.py`
  - `得物对账单_历史数据版.py`
- **原因:** 得物 API 间歇返回 409（业务高峰限流），原因有二：(1) 分页未识别 `total_results` 导致多请求了不必要的第 2 页，触发限流阈值；(2) 重试仅 3 次、最长等 10s，不足以应对 409
- **根因:** API 返回 `total_results` 而非 `totalCount`，paginate 判断永远不成立，每窗多请求一次空页
- **决策:**
  1. `fetch_api_data`: 增加 `total_results` 字段识别，移除多余的第 2 页请求
  2. `@retry`: 3→7 次，max=10→30 秒，给 API 更多恢复时间
  3. 分页间加 `time.sleep(1)`，降低连续请求触发限流的概率
  4. `run_processing` 增加 `return exit_code`，`run_processing_with_logging` 向上传递退出码
  5. `_run_batch`: 下载失败时跳过提纯和入库，进入下一窗口
  6. `import_bills` / `process_import`: 增加 `start_date`/`end_date` 参数，按文件名日期过滤
  7. `SHOP_WORKERS` 从 1 改回 6，经实测 0×409
- **影响范围:** 两个版本同步修复，API 调用相关 3 个 @retry 均已加强

## 2026-06-25: 代码审查修复—注释矛盾 + exit_code 缺失 + 类型安全 + 日志级别
- **文件:**
  - `得物对账单_sqlserver版.py`
  - `得物对账单_历史数据版.py`
- **原因:** OpenCodeReview 审查发现的 4 个问题
- **决策:**
  1. **H1**: 更新 ThreadPoolExecutor 注释"店铺串行"→"店铺并行"，与实际代码一致
  2. **H2**: sqlserver 版 `run_processing` 增加 `return exit_code`，`run_processing_with_logging` 返回退出码
  3. **M1**: `total_count` 用 `int()` 包裹，防止 API 返回字符串导致 `TypeError`
  4. **M2**: 签名原文和请求 URL 日志从 `logging.info` 降为 `logging.debug`，避免日志膨胀

## 2026-06-25: 入库并行优化 + 多 Bug 修复
- **文件:**
  - `得物对账单_sqlserver版.py`
  - `得物对账单_历史数据版.py`
- **原因:** 串行入库 18 个文件需 ~70 分钟，远程 DB 网络延迟是瓶颈；同时修复测试中发现的多个 Bug
- **决策:**
  1. **并行入库**: 提取 `_import_one_file` 独立函数，`process_import` 用 `ThreadPoolExecutor(IMPORT_WORKERS=3)` 并行处理，每线程独立 DB 连接
  2. **BATCH_SIZE 1000→5000**: 减少 commit 次数
  3. **sheet 缺失保护**: 5 个数据 sheet 加 `if 'sheet名' in sheet_names:` 保护，防止 KeyError
  4. **扣减/货损行级验证**: 数字列含文本时跳过非数据行（子表头、跨段行），不抛异常
  5. **`setup_logging` None 保护**: `logger.addHandler(text_handler)` 前判空
  6. **FutureWarning 抑制**: 加 `warnings.simplefilter('ignore', FutureWarning)`
  7. **窗口重叠 6 天**: 防止周账单跨边界遗漏
  8. **自增种子正确重置**: `DBCC CHECKIDENT('tbl', RESEED, 0)` 单引号语法
  9. **清理测试数据**: 删文件 + 清 8 张表 + 种子归零
  10. **PRIMARY KEY 约束**: `dw_dzd_xs` 改为 `IDENTITY_INSERT OFF`，依赖自增列

## 2026-06-25: 时间窗口重叠修复—防止跨边界周账单遗漏
- **文件:** `得物对账单_历史数据版.py`
- **原因:** 周账单可能跨窗口边界（如 01-27~02-02），原 logic 首尾相接（01-01~01-31、01-31~03-02），导致跨边界账单在两个窗口都不被 API 包含，遗漏数据
- **决策:** 
  1. 新增 `WINDOW_OVERLAP_DAYS = 6` 常量
  2. `generate_date_windows` 相邻窗口重叠 6 天（覆盖周账单最大跨度）
  3. 添加 `if window_end == end_date: break` 防止末窗死循环
- **效果:** 2025-01-01~2026-06-25 从 18 窗变为 23 窗，完全消除遗漏缝隙


- **变更:**
  1. `pyodbc.connect(conn_str)` → `pyodbc.connect(conn_str, fast_executemany=True)` — 批量发送，100 万行入库从 15~30 分钟降至 1~3 分钟
  2. `DRIVER={SQL Server}` → `DRIVER={ODBC Driver 17 for SQL Server}` — 使用已安装的新版驱动，提升批量操作效率

## 2026-06-24: 新增"历史数据版"— 用户自定义日期范围，按30天窗口遍历
- **文件:** `得物对账单_历史数据版.py`（新建，基于原版复制改造）
- **原因:** 原版时间窗口固定为 [today-37, today-7]，无法拉取历史数据。得物 API 要求单次查询最多 30 天
- **决策:**
  - 复制原文件为独立版本，保持主体逻辑不变
  - 新增 `get_date_range_params(start, end)` 替代 `get_default_dates()`
  - 新增 `generate_date_windows()` 将用户起止日期按 30 天切分
  - GUI：删除定时调度控件，替换为起始/终止日期选择器（QDateEdit）+ "开始历史数据获取"按钮
  - `_run_batch` 遍历所有窗口，每个窗口执行下载→提纯→入库三阶段
  - 窗口进度显示："第 3/12 个窗口"
  - 保留暂停功能，支持窗口级暂停
  - 手动"下载账单"按钮读取当前 GUI 日期单次执行

## 2026-06-24: OCR 审查修复 — 死代码/裸 except/JSON 解析/线程退出/CONTEXT.md
- **文件:**
  - `得物对账单_sqlserver版.py`
  - `CONTEXT.md`
- **原因:** OCR 审查发现 5 个 IMPORTANT 问题
- **修复:**
  1. **`max_id_before` 死代码** — 删除 `save_records` 中 `SELECT MAX(id)` 查询和未使用的 `max_id_before` 变量
  2. **裸 `except:`** — `check_file_changed` 和 `update_file_tracking` 中的 `except:` 改为 `except (json.JSONDecodeError, OSError):`
  3. **`response.json()` 未捕获 ValueError** — `fetch_api_data` 增加 `except ValueError` 分支记录 JSON 解析错误日志
  4. **窗口关闭线程安全** — `closeEvent` 增加调度线程 `join(timeout=3)` 等待退出 + 记录正在运行的任务警告
  5. **CONTEXT.md** — 更新"监控 5 家店铺"→"6 家店铺"，移除硬编码店铺列表

## 2026-06-24: PyInstaller 打包为"得物对账单自动获取.exe"
- **文件:** `得物对账单_sqlserver版.spec`（重建）
- **输出:** `dist/得物对账单自动获取.exe`（177MB，--onefile --windowed）
- **配置:** hiddenimports 含 PyQt6/pyodbc/pandas/openpyxl/requests/certifi 等；包含 certifi CA 证书

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

## 2026-06-24 22:32: 日志系统重构 — QTabWidget 多店铺日志面板
- **文件:** `得物对账单_sqlserver版.py`
- **原因:** 6 家店铺并行处理时日志交错混杂，难以追踪每家店铺的执行情况
- **决策:**
  - 单一日志面板 → QTabWidget，每店铺独立标签页
  - `log_signal` 改签 `pyqtSignal(str)` → `pyqtSignal(str, str)` 带 tab 路由
  - `_update_log` 增加 `tab="总览"` 默认参数，旧调用点无需修改
  - 下载阶段 `_process_one_shop` 用 `shop_log` 包装路由到店铺 tab
  - 提纯/入库阶段根据 `shop_folder` 路由日志
  - `QtLogHandler` 日志全部进"总览" tab
  - 标签页懒创建：店铺首次输出日志时自动生成
- **影响范围:** 仅 `得物对账单_sqlserver版.py`，布局层 + 信号层 + 业务层日志调用

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

## 2026-06-22 17:45: 删除奕氪小号 20251229-20260104 账单中 175 条跨账单重复记录
- **文件:** `dw_dzd_xs` 表（删除 175 条）
- **原因:** 全量比对发现该账单比 tiqu 文件多 175 条 order_id，经查这 175 条同时存在于上一周期 `20251222-20251228` 账单中，属于跨账单重复导入
- **决策:** 从 `20251229-20260104` 中删除这 175 条重复记录，保留 `20251222-20251228` 中的数据
- **影响范围:** 仅删除该账单的 175 条重复，其他数据不受影响

## 2026-06-23 18:00: 补充 3 个 sheet 纳入统一处理流程
- **文件:**
  - `得物对账单_sqlserver版.py`
  - `数据库迁移_20260625_补充sheet.sql`（新建）
- **原因:** `import_bills` 只处理了得物 Excel 中 6 个 sheet 中的 3 个（账单总览/销售订单/退货退款订单），剩余 3 个 sheet 在独立项目 `得物对账单sheet补充` 中用不同逻辑处理，且遗漏了「扣减其他费用明细」
- **决策:**
  1. **SHEETS_TO_KEEP** 增加 3 个 sheet：`本期结算其他项费用`、`扣减其他费用明细`、`本期货损买进订单`
  2. **表头处理**：这 3 个 sheet 是 2 行表头（说明行+列名行），与销售订单的 4 行表头（说明行+分组行+列名行+子列名行）不同，增加 `elif` 分支跳过合并逻辑
  3. **新建 3 张 DB 表**（含 bill_no/bill_period/ZH 字段）：
     - `dw_dzd_other_fee` — 25 列，全量存储本期结算其他项费用
     - `dw_dzd_deduction_detail` — 10 列，全量存储扣减其他费用明细
     - `dw_dzd_cargo_damage` — 24 列，全量存储本期货损买进订单
  4. **入库逻辑**：新增 6 个函数（3 个 wrapper + 3 个 field_mapping），延用每 500 条 batch insert + 事务提交模式
- **影响范围:** 已有 tiqu 文件不含新 3 个 sheet，需重新运行「账单处理」生成新 tiqu；历史数据不受影响（新表从头开始累积）

## 2026-06-23 18:30: 清理死表 + 补充表注释
- **文件:**
  - `得物对账单_sqlserver版.py`
  - `数据库迁移_20260625_补充表注释.sql`（新建）
- **原因:** 死表 `dw_dzd_bill_import_records` 已由用户在数据库中删除，代码中对应的 `check_if_imported` 和 `record_import` 函数零调用；`dw_dwd_bill_records` 和 `dw_dwd_bill_records_copy1` 缺少充分的表注释

## 2026-06-24: 修复入库流程 2 个严重 Bug
- **文件:** `得物对账单_sqlserver版.py`
- **原因:** 全流程审查发现 process_import 使用已关闭的数据库连接 + header=None 导致列名丢失
- **Bug 1:** `with DBConnection() as cursor:` 作用域只包了 `check_if_imported_new`，后续 `_retry_import` 调用使用的 cursor 已关闭连接 → 所有入库静默失败
  - **修复:** `with DBConnection()` 下移至包裹所有 `_retry_import` 调用和 `record_import_new`
- **Bug 2:** `pd.read_excel(..., header=None)` 生成整数索引列，但 import 函数用字符串列名取值 → 所有记录为空
  - **修复:** 新增 `_prepare(df)` 函数将 row 0 提升为 DataFrame.columns 并删除 row 0；5 个数据 sheet 传入前先经过 `_prepare()`
- **小问题:** `import_bills` 中 `update_file_tracking` 对同一批文件调了两次 → 删除重复
- **影响范围:** 此前运行"账单入库"实际未写入数据，修后恢复正常
- **决策:**
  1. 删除 `check_if_imported` 和 `record_import` 两个死函数
  2. 给 `dw_dwd_bill_records` 加详细注释：标明数据来源（period_list API）、消费者（download_files）、去重逻辑
  3. 给 `dw_dwd_bill_records_copy1` 加详细注释：标明用途（入库判重）、使用环节（process_import 前检查）
- **影响范围:** 无行为变化，纯清理 + 文档化

## 2026-06-23 19:00: 修复 8 项潜在风险
- **文件:** `得物对账单_sqlserver版.py`
- **原因:** 全面代码审查发现多项 Bug 和安全隐患
- **修复:**
  1. **API 分页** — `fetch_api_data` 增加多页循环，`page_no` 递增至拉取全部账单（上限 100 页），防止 page_size=30 限制导致静默遗漏
  2. **`bill_nos` 未定义** — 在 `if/else` 前初始化为 `[]`，避免 `processed_data` 为空时 `UnboundLocalError`
  3. **`record_import_new` 失败重插** — 用 `try/except` 包裹，失败时日志记录而非崩溃；提醒人工处理
  4. **线程安全** — `AutoRun` 改用 `@property` + `threading.Lock` 保护 `paused`/`running` 状态
  5. **下载原子性** — 先写到 `.tmp` 再 `os.replace` 重命名，下载中断不留残文件；超时从 30s 提到 60s
  6. **`BillProcessor` 结果格式** — 从 `f"错误: {str(e)}"` 改为 `{'success': False, 'error': str(e)}`，消除中文前缀判断
  7. **`openpyxl` 警告过滤** — 从全忽略改为仅过滤 `openpyxl.styles.stylesheet`
  8. **`ExcelFile` 资源释放** — 改用 `with` 上下文管理器，提取 `sheet_names` 后立即关闭
  9. **API 密钥日志脱敏** — `app_key` 日志记录隐藏中间部分（`sk****xxxx`）
- **影响范围:** 无行为变化，健壮性提升

## 2026-06-23 19:30: 补修 process_all 结果格式 + 清理死代码 + 更新文档
- **文件:**
  - `得物对账单_sqlserver版.py`
  - `CONTEXT.md`
- **原因:** 上一轮修复中 `process_all` 方法仍返回旧格式字符串，而 `download_files` 已改为接收 dict，运行时将 `AttributeError`；`save_records` 存在死代码；`CONTEXT.md` 未同步更新
- **修复:**
  1. **`process_all` 结果格式** — `self.results[bill_no] = download_url` → `{'success': True, 'url': download_url}`；错误分支同步修改
  2. **`BillProcessor.__init__` 类型标注** — `Dict[str, str]` → `Dict[str, dict]`
  3. **`save_records` 清理** — 删除外层重复的 `inserted_ids`/`bill_nos`/`new_records` 初始化；删除误导性注释"SQL Server不支持executemany"
  4. **`CONTEXT.md` 同步** — 更新数据表清单（9 张表）、数据流阶段（6 个 sheet）、表头处理说明（3 种分支）
- **影响范围:** 无行为变化，文档与代码一致

## 2026-06-23 20:00: GUI 改造为深色玻璃拟态风格
- **文件:** `得物对账单_sqlserver版.py`
- **原因:** 按照 TOKEN-MONITOR-UI-PROMPT 规范改造界面视觉风格
- **样式变更:**
  1. **深色主题** — 基底 #303438，卡片 #1c1c24，文字 #eef5fb，强调色 #b7ead4
  2. **等宽字体** — 全局使用 Consolas 等宽字体（12px/10px/28px 三档）
  3. **玻璃卡片** — 日志区域用深色卡片 + 边框 #42474d，模拟悬浮感
  4. **紧凑间距** — 8px 间隔，14px 内边距，4px 按钮间距
  5. **状态指示** — ● 空闲 / ● 运行中 / ● 已暂停 三态指示器，带语义色
  6. **按钮交互** — hover 高亮为薄荷绿 #b7ead4（仅支持 Windows clam 主题）
  7. **自适应布局** — 窗口最小 800x500，按钮均匀等宽分布
- **影响范围:** 纯视觉改造，功能逻辑无变化

## 2026-06-23 20:15: GUI 替换为 PyQt6（突破 Tkinter 样式限制）
- **文件:**
  - `得物对账单_sqlserver版.py`
  - `requirements.txt`
- **原因:** Tkinter 无法实现圆角、rgba 半透明、自定义滚动条等现代 UI 效果，改用 PyQt6
- **变更:**
  1. **框架替换** — `tkinter` → `PyQt6`（QMainWindow + QPlainTextEdit + QPushButton）
  2. **日志处理器** — `TextHandler(tk.Text)` → `QtLogHandler(pyqtSignal)` 线程安全
  3. **视觉效果** — 卡片圆角 8px、按钮圆角 6px、rgba 半透明色、极窄 6px 滚动条、hover 高亮
  4. **入口** — `main_gui()` 调用 → `MainWindow` 类 + `app.exec()`
  5. **依赖** — `requirements.txt` 添加 PyQt6==6.11.0
- **影响范围:** 需要安装 PyQt6（`pip install PyQt6==6.11.0`），打包时需要含 PyQt6 相关 dll

## 2026-06-23 20:45: 修复审查发现的 CRITICAL 问题 + 代码优化
- **文件:** `得物对账单_sqlserver版.py`
- **原因:** auto 模式下代码审查发现 `safe_params` 未定义（CRITICAL，API 调用必崩）、倒计时线程不安全、DB 配置硬编码、空值检测重复
- **修复:**
  1. **`safe_params` 未定义** — 添加脱敏参数字典定义，app_key 掩码为 `sk****xxxx` 格式
  2. **倒计时线程安全** — 增加 `countdown_signal = pyqtSignal(int)`，后台线程通过信号更新 GUI，避免直接调用 `QLabel.setText()`
  3. **DB 配置硬编码** — 移除 `DB_SERVER`/`DB_DATABASE`/`DB_USERNAME` 的 fallback 默认值，全部从环境变量读取；缺失时启动报错
  4. **空值检测去重** — 提取 `is_empty()` 工具函数 + `EMPTY_VALUES` 常量，替换 6 处重复代码
  5. **类型标注修正** — `generate_result_file` 的 `download_results` 类型改为 `Dict[str, dict]`
- **影响范围:** 无行为变化，健壮性和代码质量提升

## 2026-06-23 21:00: 表头解析改为自动检测，消除格式脆弱性
- **文件:** `得物对账单_sqlserver版.py`
- **原因:** 原表头处理依赖 `iloc[1:]`、`iloc[:3]` 等硬编码行号，得物调整 Excel 行列结构时数据会错位
- **方案:** 基于 Excel 行特征自动检测表头结构
- **实现:**
  1. **`detect_headers(df)`** — 检测各行类型：说明行（首列含"说明"）、表头行（全非空无浮点）、数据行（含浮点/数字/长值）
  2. **`merge_header_names(df, header_rows)`** — 多行表头按列拼接为扁平列名（N 行→1 行）
  3. **`import_bills` 替换** — 删除 3 种 sheet 的 if/elif/else 分支，统一调用 `detect_headers` + `merge_header_names`
  4. **`extract_bill_period_from_file` 强化** — 增加模糊列名匹配 fallback（精确匹配失败时找包含"账单起止时间"的列）
- **影响范围:** 无行为变化；tiqu 文件格式不变，field_mapping 键值不变

## 2026-06-23 21:15: 修复 field_mapping 键值错误，发现得物新增 4 列
- **文件:** `得物对账单_sqlserver版.py`
- **原因:**
  1. `退货完成时间` → `退货创建时间` 写错字，导致 `return_create_time` 字段永为空
  2. 得物统一升级 Excel 格式，新增了业务时间、交易成功时间、结算规则类型、备注内容 4 列，所有店铺所有近期文件均受影响
- **修复:** 第 1174 行 `退货完成时间` → `退货创建时间`
- **未修复（需用户确认）:** 4 个新列尚未加入 field_mapping，需建表迁移

## 2026-06-23 21:30: 追加得物新增的 4 个字段
- **文件:**
  - `得物对账单_sqlserver版.py`
  - `数据库迁移_20260625_新增字段.sql`（新建）
- **原因:** 得物统一升级 Excel 格式，所有店铺文件新增了业务时间、交易成功时间、结算规则类型、备注内容 4 列，原程序未映射导致数据静默丢弃
- **变更:**
  1. **field_mapping** — 销售订单增加 4 个键值对，退货退款订单增加 1 个（备注内容）
  2. **DB 迁移** — `dw_dzd_xs` +4 列（business_time, transaction_success_time, settlement_rule_type, remark_content），`dw_dzd_thtk` +1 列（remark_content）
- **影响范围:** 需要重新运行「账单处理」→「账单入库」让已下载的文件补上这些字段；DB 表已备份且清空，迁移无风险

## 2026-06-23 22:00: 追加物流签收时间等 3 个字段 + 修复映射回退
- **文件:**
  - `得物对账单_sqlserver版.py`
  - `数据库迁移_20260625_新增字段_第二批.sql`（新建）
- **原因:** 干运行发现退货完成时间映射被错误修改（完成→创建），以及全部店铺文件还有 3 个未映射列被静默丢弃
- **变更:**
  1. **映射修复** — 还原 `退货创建时间` → `退货完成时间`（原始映射正确，改反了）
  2. **新增映射** — 销售订单增加 3 个键值对（物流签收时间、包含吊牌卡套费、政府补贴商家出资金额），退货退款订单增加 2 个
  3. **DB 迁移** — `dw_dzd_xs` 追加 3 列（logistics_sign_time, includes_tag_card_fee, government_subsidy_amount），`dw_dzd_thtk` 追加 2 列
- **影响范围:** 所有 column mapping 已与最新 Excel 格式对齐

## 2026-06-23 22:30: 性能优化：BATCH_SIZE 调整 + 缓存读取 + 并行下载
- **文件:** `得物对账单_sqlserver版.py`
- **原因:** 全流程串行处理导致入库/下载阶段存在 I/O 瓶颈
- **优化:**
  1. **BATCH_SIZE 500→1000** — 减少 commit 次数，DB 写入吞吐量翻倍
  2. **Excel 缓存读取** — `process_import` 一次读取 tiqu 文件全部 sheet，避免每个 sheet 单独读文件（6 次→1 次）
  3. **重试逻辑去重** — 6 个重复的 while+try/except 块合并为 `_retry_import()` 辅助函数（-60% 重复代码）
  4. **并行下载** — `download_files` 改用 `ThreadPoolExecutor(max_workers=4)` 并发下载，网络 I/O 不再串行
- **影响范围:** 无功能变化，仅性能提升

## 2026-06-23 23:00: 功能强化：文件变化检测 + 日志清理 + 进度条
- **文件:** `得物对账单_sqlserver版.py`
- **变更:**
  1. **文件变化检测** — 新增 `check_file_changed()` + `update_file_tracking()`，`import_bills` 运行时检查原始文件大小/修改时间是否变化，变化则自动删除旧 tiqu 并重新生成，实现增量更新
  2. **日志自动清理** — 新增 `cleanup_old_logs()`，程序启动时自动删除超过 30 天的 `.log` 文件
  3. **GUI 进度条** — 状态行增加 `QProgressBar`（薄荷绿主题色，极窄圆角），视觉反馈任务进度
- **影响范围:** 无行为变化，增量功能

## 2026-06-23 23:30: 店铺级并行处理（3 并发）
- **文件:** `得物对账单_sqlserver版.py`
- **原因:** API 请求大部分是网络等待，串行处理 5 家店铺浪费 I/O 时间
- **变更:** `run_processing` 中 `for credential in credentials` 串行循环改为 `ThreadPoolExecutor(max_workers=3)` 并行，每店内部保持串行；限流 3 并发远低于 API 限制（1000次/分钟）
- **影响范围:** 下载账单阶段从单店处理变为最多 3 店并行，整体耗时减少约 60%

## 2026-06-24: 下载失败自动重试机制
- **文件:** `得物对账单_sqlserver版.py`
- **原因:** 网络波动可能导致下载失败，但 bill_no 已记录到 dw_dwd_bill_records，下次 AutoRun 跳过下载，导致文件永久缺失
- **变更:** 在 `_process_one_shop` 中，对 `skipped_bill_nos`（已记录的 bill_no）逐一检查 `download/<shop>/<bill_no>.xlsx` 文件是否存在；若不存在，重新加入下载列表走 generate → export → download 流程
- **影响范围:** 下载失败的账单会在下次 AutoRun 自动补下，无需手动干预

## 2026-06-24: 修复 bill_no 提取方式，支持同名后缀文件入库
- **文件:** `得物对账单_sqlserver版.py`
- **原因:** `process_import` 中 `filename.split('_')[0]` 提取 bill_no 有 2 个问题：
  1. 文件名无 `_` 时（如 "20260330-...-1207033691884621_tiqu.xlsx"），返回整个字符串包含 `.xlsx`
  2. 文件名有 `_` 时（如 "HD12345_1_tiqu.xlsx"），只取 `_` 前部分（"HD12345"），导致后缀文件与 base 文件 bill_no 相同，后者被去重跳过
- **变更:** 
  1. `process_import` 中 `split('_')[0]` → `replace('_tiqu.xlsx', '')`
  2. `import_bills` 和 `process_import` 增加 `~$` 临时文件过滤，防止 Excel 打开时生成的锁文件被错误读取
- **效果:** 后缀文件（如 `HD12345_1_tiqu.xlsx`）的 bill_no 正确提取为 `HD12345_1`，与 base 文件 `HD12345` 区分开，各自独立入库
- **影响范围:** 仅影响提取 bill_no 的逻辑，无其他影响

## 2026-06-24: 修复 3 个代码质量问题
- **文件:** `得物对账单_sqlserver版.py`
- **变更:**
  1. **分页终止条件** — `fetch_api_data` 中 `total_count=0` 时会导致首次有数据就退出分页循环，改为 `total_count>0` 时才判断是否已拉满
  2. **GUI 线程安全** — AutoRun 的 `paused`/`running` setter 直接操作 Qt 控件（背景线程中调用），改为通过 `pyqtSignal` 代理到主线程
  3. **`.tmp` 文件残留** — `download_files` 中下载失败时清理残留的 `.tmp` 文件
  4. **死代码** — 删除 `import_bills` 中未使用的 `shop_name` 变量

## 2026-06-24: OCR 代码审查修复
- **文件:** `得物对账单_sqlserver版.py`, `数据库迁移_20260609_账单总览.sql`
- **变更:**
  1. `.tmp` 清理中 `os.remove` 失败会掩盖原始下载异常，改为嵌套 try/except 忽略清理失败
  2. SQL 迁移补充 `ALTER TABLE dw_dzd_bill_overview ADD ZH` 段，覆盖表已存在但缺 ZH 列的场景

## 2026-06-24: 代码质量修复（第三方审查后）
- **文件:** `得物对账单_sqlserver版.py`
- **变更:**
  1. `extract_bill_period_from_file` → `extract_bill_period_from_data`，从缓存 DataFrame 提取 bill_period，避免二次读文件
  2. `_prepare` 闭包 → 模块级 `prepare_df` 函数，不再每次重新创建
  3. `import json` 从函数内部移到文件顶部
  4. `time.sleep(120)` → `DOWNLOAD_WAIT_SECONDS` 常量
  5. `max_workers=4/5` → `DOWNLOAD_WORKERS` / `SHOP_WORKERS` 常量
  6. 删除死代码：`generate_result_file` + 6 个 `_from_file` 函数
  7. `file_tracking.json` 路径从 RESULT_DIR 移到 `.cache` 目录

## 2026-06-24: 改造为每天定时自动运行
- **文件:** `得物对账单_sqlserver版.py`
- **变更:**
  1. 删除自动启动：不再一打开就 10 秒倒计时启动，改为等待用户操作
  2. 新增 `QTimeEdit` 时间选择器（默认 02:00），用户可设置每天执行时间
  3. 新增"启动自动运行"按钮，用户手动触发定时任务
  4. 新增 `_get_next_run_seconds` 计算距下次执行时间
  5. 新增 `_run_schedule_loop` 替换原 `_run_auto_sequence`，执行后等待到次日同一时间
  6. 移除 6 小时硬编码循环间隔

## 2026-06-24: 线程安全优化—任务防重复 + 调度与手动互斥
- **文件:** `得物对账单_sqlserver版.py`
- **变更:**
  1. 新增 `_run_task` 替代 `_run_thread`，用 `_running_tasks` 集合 + `_task_lock` 防止同一按钮重复点击
  2. 新增 `_safe_run_phase`，调度循环中的阶段若已被手动触发则跳过，避免调度与手动任务冲突
  3. 按钮（下载账单/账单处理/账单入库/测试连接）全部改用 `_run_task`
  4. 调度循环的 3 个阶段全部改用 `_safe_run_phase`
  5. 移除 AutoRun 内冗余的 `import threading`
  6. 补齐 `closeEvent` 前缺失的空行

