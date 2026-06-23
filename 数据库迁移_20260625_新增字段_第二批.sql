-- ============================================================
-- 得物对账单系统 - 数据库迁移脚本（第二批字段追加）
-- 版本: 2026-06-25
-- 功能: 追加物流签收时间、包含吊牌卡套费、政府补贴商家出资金额
-- ============================================================

-- ============================================================
-- 1. dw_dzd_xs 追加 3 个字段
-- ============================================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dw_dzd_xs') AND name = 'logistics_sign_time')
BEGIN
    ALTER TABLE dw_dzd_xs ADD logistics_sign_time DATETIME;
    EXEC sp_addextendedproperty 'MS_Description', '物流签收时间', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_xs', 'COLUMN', 'logistics_sign_time';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dw_dzd_xs') AND name = 'includes_tag_card_fee')
BEGIN
    ALTER TABLE dw_dzd_xs ADD includes_tag_card_fee DECIMAL(18,2);
    EXEC sp_addextendedproperty 'MS_Description', '包含吊牌卡套费', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_xs', 'COLUMN', 'includes_tag_card_fee';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dw_dzd_xs') AND name = 'government_subsidy_amount')
BEGIN
    ALTER TABLE dw_dzd_xs ADD government_subsidy_amount DECIMAL(18,2);
    EXEC sp_addextendedproperty 'MS_Description', '政府补贴商家出资金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_xs', 'COLUMN', 'government_subsidy_amount';
END
GO

-- ============================================================
-- 2. dw_dzd_thtk 追加 2 个字段
-- ============================================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dw_dzd_thtk') AND name = 'includes_tag_card_fee')
BEGIN
    ALTER TABLE dw_dzd_thtk ADD includes_tag_card_fee DECIMAL(18,2);
    EXEC sp_addextendedproperty 'MS_Description', '包含吊牌卡套费', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_thtk', 'COLUMN', 'includes_tag_card_fee';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dw_dzd_thtk') AND name = 'government_subsidy_amount')
BEGIN
    ALTER TABLE dw_dzd_thtk ADD government_subsidy_amount DECIMAL(18,2);
    EXEC sp_addextendedproperty 'MS_Description', '政府补贴商家出资金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_thtk', 'COLUMN', 'government_subsidy_amount';
END
GO

PRINT '字段追加完成！dw_dzd_xs(+3), dw_dzd_thtk(+2)';
GO
