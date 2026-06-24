-- ============================================================
-- 得物对账单系统 - 数据库迁移脚本
-- 版本: 2026-06-09
-- 功能: 1. 创建 dw_dzd_bill_overview 表（含中文注释）
--       2. 给 dw_dzd_xs / dw_dzd_thtk 追加 bill_no, bill_period
-- ============================================================

-- ============================================================
-- 1. 创建账单总览表
-- ============================================================
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'dw_dzd_bill_overview')
BEGIN
    CREATE TABLE dw_dzd_bill_overview (
        id INT IDENTITY(1,1) PRIMARY KEY,
        bill_no VARCHAR(100),
        company_name VARCHAR(200),
        settlement_cycle VARCHAR(50),
        bill_period VARCHAR(50),
        update_time_desc VARCHAR(50),
        product_transaction_amount DECIMAL(18,2),
        platform_service_fee DECIMAL(18,2),
        distribution_service_fee DECIMAL(18,2),
        advance_payment_recovery DECIMAL(18,2),
        seller_subsidy_amount DECIMAL(18,2),
        adjustment_item DECIMAL(18,2),
        seller_return_shipping_fee VARCHAR(50),
        after_sales_service DECIMAL(18,2),
        trade_in_subsidy_amount DECIMAL(18,2),
        other_non_transaction_receivable DECIMAL(18,2),
        other_non_transaction_platform_fee DECIMAL(18,2),
        other_advance_payment_recovery DECIMAL(18,2),
        other_non_transaction_settlement DECIMAL(18,2),
        other_deductions DECIMAL(18,2),
        price_reduction_refund DECIMAL(18,2),
        price_reduction_subsidy DECIMAL(18,2),
        total_payable_amount DECIMAL(18,2),
        settlement_status VARCHAR(50),
        ZH VARCHAR(100),
        create_time DATETIME DEFAULT GETDATE()
    );
END
GO

-- 针对已存在 dw_dzd_bill_overview 表但缺少 ZH 列的数据库做补充
IF EXISTS (SELECT * FROM sys.tables WHERE name = 'dw_dzd_bill_overview')
   AND NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dw_dzd_bill_overview') AND name = 'ZH')
BEGIN
    ALTER TABLE dw_dzd_bill_overview ADD ZH VARCHAR(100);
    EXEC sp_addextendedproperty 'MS_Description', '店铺名称', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'ZH';
END
GO

-- ============================================================
-- 2. 表中文注释
-- ============================================================
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', NULL, NULL))
    EXEC sp_addextendedproperty 'MS_Description', '得物对账单 - 账单总览表', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview';
GO

-- ============================================================
-- 3. 字段中文注释
-- ============================================================

-- 账单编号
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'bill_no'))
    EXEC sp_addextendedproperty 'MS_Description', '账单编号', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'bill_no';
-- 公司名称
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'company_name'))
    EXEC sp_addextendedproperty 'MS_Description', '公司名称', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'company_name';
-- 结算周期
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'settlement_cycle'))
    EXEC sp_addextendedproperty 'MS_Description', '结算周期', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'settlement_cycle';
-- 账单起止时间
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'bill_period'))
    EXEC sp_addextendedproperty 'MS_Description', '账单起止时间', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'bill_period';
-- 对账单更新时间
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'update_time_desc'))
    EXEC sp_addextendedproperty 'MS_Description', '对账单更新时间', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'update_time_desc';
-- 本期商品交易金额
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'product_transaction_amount'))
    EXEC sp_addextendedproperty 'MS_Description', '本期商品交易金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'product_transaction_amount';
-- 本期交易类平台服务费金额
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'platform_service_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '本期交易类平台服务费金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'platform_service_fee';
-- 分销服务费
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'distribution_service_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '分销服务费', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'distribution_service_fee';
-- 平台预付款收回金额
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'advance_payment_recovery'))
    EXEC sp_addextendedproperty 'MS_Description', '平台预付款收回金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'advance_payment_recovery';
-- 卖家补贴金额
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'seller_subsidy_amount'))
    EXEC sp_addextendedproperty 'MS_Description', '卖家补贴金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'seller_subsidy_amount';
-- 调整项
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'adjustment_item'))
    EXEC sp_addextendedproperty 'MS_Description', '调整项', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'adjustment_item';
-- 卖家退运服务费
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'seller_return_shipping_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '卖家退运服务费', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'seller_return_shipping_fee';
-- 售后无忧
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'after_sales_service'))
    EXEC sp_addextendedproperty 'MS_Description', '售后无忧', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'after_sales_service';
-- 以旧换新补贴金额
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'trade_in_subsidy_amount'))
    EXEC sp_addextendedproperty 'MS_Description', '以旧换新补贴金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'trade_in_subsidy_amount';
-- 其他非交易类应收商品金额
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'other_non_transaction_receivable'))
    EXEC sp_addextendedproperty 'MS_Description', '其他非交易类应收商品金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'other_non_transaction_receivable';
-- 其他非交易类平台费用
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'other_non_transaction_platform_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '其他非交易类平台费用', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'other_non_transaction_platform_fee';
-- 平台预付款收回金额(其他非交易)
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'other_advance_payment_recovery'))
    EXEC sp_addextendedproperty 'MS_Description', '平台预付款收回金额(其他非交易)', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'other_advance_payment_recovery';
-- 其他非交易类应结金额
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'other_non_transaction_settlement'))
    EXEC sp_addextendedproperty 'MS_Description', '其他非交易类应结金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'other_non_transaction_settlement';
-- 扣减其他费用
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'other_deductions'))
    EXEC sp_addextendedproperty 'MS_Description', '扣减其他费用', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'other_deductions';
-- 售中降价(退款)
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'price_reduction_refund'))
    EXEC sp_addextendedproperty 'MS_Description', '售中降价(退款)', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'price_reduction_refund';
-- 售中降价(退津贴)
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'price_reduction_subsidy'))
    EXEC sp_addextendedproperty 'MS_Description', '售中降价(退津贴)', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'price_reduction_subsidy';
-- 应结总金额
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'total_payable_amount'))
    EXEC sp_addextendedproperty 'MS_Description', '应结总金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'total_payable_amount';
-- 结算状态
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'settlement_status'))
    EXEC sp_addextendedproperty 'MS_Description', '结算状态', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'settlement_status';
-- 店铺名称
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'ZH'))
    EXEC sp_addextendedproperty 'MS_Description', '店铺名称', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_bill_overview', 'COLUMN', 'ZH';
GO

-- ============================================================
-- 4. 给 dw_dzd_xs 追加字段
-- ============================================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dw_dzd_xs') AND name = 'bill_no')
BEGIN
    ALTER TABLE dw_dzd_xs ADD bill_no VARCHAR(100);
    EXEC sp_addextendedproperty 'MS_Description', '账单编号', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_xs', 'COLUMN', 'bill_no';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dw_dzd_xs') AND name = 'bill_period')
BEGIN
    ALTER TABLE dw_dzd_xs ADD bill_period VARCHAR(50);
    EXEC sp_addextendedproperty 'MS_Description', '账单起止时间', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_xs', 'COLUMN', 'bill_period';
END
GO

-- ============================================================
-- 5. 给 dw_dzd_thtk 追加字段
-- ============================================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dw_dzd_thtk') AND name = 'bill_no')
BEGIN
    ALTER TABLE dw_dzd_thtk ADD bill_no VARCHAR(100);
    EXEC sp_addextendedproperty 'MS_Description', '账单编号', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_thtk', 'COLUMN', 'bill_no';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dw_dzd_thtk') AND name = 'bill_period')
BEGIN
    ALTER TABLE dw_dzd_thtk ADD bill_period VARCHAR(50);
    EXEC sp_addextendedproperty 'MS_Description', '账单起止时间', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_thtk', 'COLUMN', 'bill_period';
END
GO

PRINT '数据库迁移完成！';
GO
