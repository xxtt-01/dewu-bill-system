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
