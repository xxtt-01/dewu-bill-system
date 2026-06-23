-- ============================================================
-- 得物对账单系统 - 数据库迁移脚本
-- 版本: 2026-06-25
-- 功能: 给 dw_dzd_xs / dw_dzd_thtk 追加得物新增字段
--       业务时间、交易成功时间、结算规则类型、备注内容
-- ============================================================

-- ============================================================
-- 1. 给 dw_dzd_xs 追加字段
-- ============================================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dw_dzd_xs') AND name = 'business_time')
BEGIN
    ALTER TABLE dw_dzd_xs ADD business_time DATETIME;
    EXEC sp_addextendedproperty 'MS_Description', '业务时间', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_xs', 'COLUMN', 'business_time';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dw_dzd_xs') AND name = 'transaction_success_time')
BEGIN
    ALTER TABLE dw_dzd_xs ADD transaction_success_time DATETIME;
    EXEC sp_addextendedproperty 'MS_Description', '交易成功时间', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_xs', 'COLUMN', 'transaction_success_time';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dw_dzd_xs') AND name = 'settlement_rule_type')
BEGIN
    ALTER TABLE dw_dzd_xs ADD settlement_rule_type VARCHAR(100);
    EXEC sp_addextendedproperty 'MS_Description', '结算规则类型', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_xs', 'COLUMN', 'settlement_rule_type';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dw_dzd_xs') AND name = 'remark_content')
BEGIN
    ALTER TABLE dw_dzd_xs ADD remark_content NVARCHAR(500);
    EXEC sp_addextendedproperty 'MS_Description', '备注内容', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_xs', 'COLUMN', 'remark_content';
END
GO

-- ============================================================
-- 2. 给 dw_dzd_thtk 追加字段
-- ============================================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dw_dzd_thtk') AND name = 'remark_content')
BEGIN
    ALTER TABLE dw_dzd_thtk ADD remark_content NVARCHAR(500);
    EXEC sp_addextendedproperty 'MS_Description', '备注内容', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_thtk', 'COLUMN', 'remark_content';
END
GO

PRINT '字段追加完成！dw_dzd_xs(+4), dw_dzd_thtk(+1)';
GO
