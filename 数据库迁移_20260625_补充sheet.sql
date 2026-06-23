-- ============================================================
-- 得物对账单系统 - 数据库迁移脚本
-- 版本: 2026-06-25
-- 功能: 创建 3 个补充 sheet 数据表
--       本期结算其他项费用 → dw_dzd_other_fee
--       扣减其他费用明细   → dw_dzd_deduction_detail
--       本期货损买进订单   → dw_dzd_cargo_damage
-- ============================================================

-- ============================================================
-- 1. 创建本期结算其他项费用表
-- ============================================================
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'dw_dzd_other_fee')
BEGIN
    CREATE TABLE dw_dzd_other_fee (
        id INT IDENTITY(1,1) PRIMARY KEY,
        fee_type VARCHAR(100),
        order_id VARCHAR(100),
        original_order_id VARCHAR(100),
        order_type VARCHAR(100),
        occur_time DATETIME,
        product_amount DECIMAL(18,2),
        advance_payment_recovery DECIMAL(18,2),
        operation_service_fee DECIMAL(18,2),
        certification_direct_fee DECIMAL(18,2),
        transfer_fee DECIMAL(18,2),
        technical_service_fee DECIMAL(18,2),
        customer_service_fee DECIMAL(18,2),
        seller_subsidy_buyer_product DECIMAL(18,2),
        price_reduction_refund DECIMAL(18,2),
        price_reduction_subsidy DECIMAL(18,2),
        export_promotion_fee DECIMAL(18,2),
        seller_subsidy_installment_fee DECIMAL(18,2),
        other_compensation DECIMAL(18,2),
        settlement_amount DECIMAL(18,2),
        settlement_rate DECIMAL(18,6),
        bill_no VARCHAR(100),
        bill_period VARCHAR(50),
        ZH VARCHAR(100),
        create_time DATETIME DEFAULT GETDATE()
    );
END
GO

-- ============================================================
-- 2. 创建扣减其他费用明细表
-- ============================================================
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'dw_dzd_deduction_detail')
BEGIN
    CREATE TABLE dw_dzd_deduction_detail (
        id INT IDENTITY(1,1) PRIMARY KEY,
        fee_type VARCHAR(100),
        total_repayment DECIMAL(18,2),
        detail_item VARCHAR(200),
        repayment_amount DECIMAL(18,2),
        currency VARCHAR(10),
        bill_no VARCHAR(100),
        bill_period VARCHAR(50),
        ZH VARCHAR(100),
        create_time DATETIME DEFAULT GETDATE()
    );
END
GO

-- ============================================================
-- 3. 创建本期货损买进订单表
-- ============================================================
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'dw_dzd_cargo_damage')
BEGIN
    CREATE TABLE dw_dzd_cargo_damage (
        id INT IDENTITY(1,1) PRIMARY KEY,
        bill_no_from_sheet VARCHAR(100),
        purchase_order_id VARCHAR(100),
        original_order_id VARCHAR(100),
        order_type VARCHAR(100),
        occur_time DATETIME,
        product_amount DECIMAL(18,2),
        advance_payment_recovery DECIMAL(18,2),
        operation_service_fee DECIMAL(18,2),
        certification_direct_fee DECIMAL(18,2),
        technical_service_fee DECIMAL(18,2),
        platform_base_service_fee DECIMAL(18,2),
        base_service_fee_amount DECIMAL(18,2),
        performance_service_fee_amount DECIMAL(18,2),
        transfer_fee DECIMAL(18,2),
        after_sales_service_fee DECIMAL(18,2),
        consumer_shipping_subsidy DECIMAL(18,2),
        customer_service_fee DECIMAL(18,2),
        packaging_service_fee DECIMAL(18,2),
        settlement_amount DECIMAL(18,2),
        bill_no VARCHAR(100),
        bill_period VARCHAR(50),
        ZH VARCHAR(100),
        create_time DATETIME DEFAULT GETDATE()
    );
END
GO

-- ============================================================
-- 4. 表中文注释
-- ============================================================
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', NULL, NULL))
    EXEC sp_addextendedproperty 'MS_Description', '得物对账单 - 本期结算其他项费用', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee';
GO
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail', NULL, NULL))
    EXEC sp_addextendedproperty 'MS_Description', '得物对账单 - 扣减其他费用明细', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail';
GO
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', NULL, NULL))
    EXEC sp_addextendedproperty 'MS_Description', '得物对账单 - 本期货损买进订单', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage';
GO

-- ============================================================
-- 5. 字段中文注释
-- ============================================================

-- dw_dzd_other_fee
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'fee_type'))
    EXEC sp_addextendedproperty 'MS_Description', '费用类型', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'fee_type';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'order_id'))
    EXEC sp_addextendedproperty 'MS_Description', '订单号', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'order_id';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'original_order_id'))
    EXEC sp_addextendedproperty 'MS_Description', '原订单号', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'original_order_id';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'order_type'))
    EXEC sp_addextendedproperty 'MS_Description', '订单类型', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'order_type';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'occur_time'))
    EXEC sp_addextendedproperty 'MS_Description', '发生时间', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'occur_time';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'product_amount'))
    EXEC sp_addextendedproperty 'MS_Description', '商品金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'product_amount';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'advance_payment_recovery'))
    EXEC sp_addextendedproperty 'MS_Description', '平台预付款收回金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'advance_payment_recovery';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'operation_service_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '操作服务费', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'operation_service_fee';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'certification_direct_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '认证直发服务费', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'certification_direct_fee';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'transfer_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '转账手续费', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'transfer_fee';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'technical_service_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '技术服务费', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'technical_service_fee';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'customer_service_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '客服托管服务费', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'customer_service_fee';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'seller_subsidy_buyer_product'))
    EXEC sp_addextendedproperty 'MS_Description', '卖家补贴买家(商品)', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'seller_subsidy_buyer_product';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'price_reduction_refund'))
    EXEC sp_addextendedproperty 'MS_Description', '售中降价(退款)', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'price_reduction_refund';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'price_reduction_subsidy'))
    EXEC sp_addextendedproperty 'MS_Description', '售中降价(退津贴)', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'price_reduction_subsidy';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'export_promotion_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '出口推广服务费', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'export_promotion_fee';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'seller_subsidy_installment_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '卖家补贴买家(分期手续费)', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'seller_subsidy_installment_fee';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'other_compensation'))
    EXEC sp_addextendedproperty 'MS_Description', '其他赔付项', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'other_compensation';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'settlement_amount'))
    EXEC sp_addextendedproperty 'MS_Description', '结算金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'settlement_amount';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'settlement_rate'))
    EXEC sp_addextendedproperty 'MS_Description', '结算汇率', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'settlement_rate';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'bill_no'))
    EXEC sp_addextendedproperty 'MS_Description', '账单编号', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'bill_no';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'bill_period'))
    EXEC sp_addextendedproperty 'MS_Description', '账单起止时间', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'bill_period';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'ZH'))
    EXEC sp_addextendedproperty 'MS_Description', '店铺名称', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_other_fee', 'COLUMN', 'ZH';
GO

-- dw_dzd_deduction_detail
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail', 'COLUMN', 'fee_type'))
    EXEC sp_addextendedproperty 'MS_Description', '费用类型', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail', 'COLUMN', 'fee_type';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail', 'COLUMN', 'total_repayment'))
    EXEC sp_addextendedproperty 'MS_Description', '偿还总金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail', 'COLUMN', 'total_repayment';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail', 'COLUMN', 'detail_item'))
    EXEC sp_addextendedproperty 'MS_Description', '明细费用项', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail', 'COLUMN', 'detail_item';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail', 'COLUMN', 'repayment_amount'))
    EXEC sp_addextendedproperty 'MS_Description', '偿还金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail', 'COLUMN', 'repayment_amount';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail', 'COLUMN', 'currency'))
    EXEC sp_addextendedproperty 'MS_Description', '币种', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail', 'COLUMN', 'currency';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail', 'COLUMN', 'bill_no'))
    EXEC sp_addextendedproperty 'MS_Description', '账单编号', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail', 'COLUMN', 'bill_no';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail', 'COLUMN', 'bill_period'))
    EXEC sp_addextendedproperty 'MS_Description', '账单起止时间', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail', 'COLUMN', 'bill_period';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail', 'COLUMN', 'ZH'))
    EXEC sp_addextendedproperty 'MS_Description', '店铺名称', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_deduction_detail', 'COLUMN', 'ZH';
GO

-- dw_dzd_cargo_damage
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'bill_no_from_sheet'))
    EXEC sp_addextendedproperty 'MS_Description', '账单编号(来自sheet)', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'bill_no_from_sheet';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'purchase_order_id'))
    EXEC sp_addextendedproperty 'MS_Description', '买进订单号', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'purchase_order_id';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'original_order_id'))
    EXEC sp_addextendedproperty 'MS_Description', '原订单号', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'original_order_id';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'order_type'))
    EXEC sp_addextendedproperty 'MS_Description', '订单类型', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'order_type';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'occur_time'))
    EXEC sp_addextendedproperty 'MS_Description', '发生时间', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'occur_time';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'product_amount'))
    EXEC sp_addextendedproperty 'MS_Description', '商品金额（元）', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'product_amount';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'advance_payment_recovery'))
    EXEC sp_addextendedproperty 'MS_Description', '平台预付款收回金额（元）', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'advance_payment_recovery';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'operation_service_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '操作服务费（元）', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'operation_service_fee';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'certification_direct_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '认证直发服务费（元）', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'certification_direct_fee';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'technical_service_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '技术服务费（元）', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'technical_service_fee';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'platform_base_service_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '平台基础服务费（元）', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'platform_base_service_fee';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'base_service_fee_amount'))
    EXEC sp_addextendedproperty 'MS_Description', '其中:基础服务费金额（元）', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'base_service_fee_amount';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'performance_service_fee_amount'))
    EXEC sp_addextendedproperty 'MS_Description', '其中:履约服务费金额（元）', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'performance_service_fee_amount';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'transfer_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '转账手续费（元）', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'transfer_fee';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'after_sales_service_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '售后无忧服务费（元）', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'after_sales_service_fee';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'consumer_shipping_subsidy'))
    EXEC sp_addextendedproperty 'MS_Description', '消费者邮费补贴（元）', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'consumer_shipping_subsidy';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'customer_service_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '客服托管服务费（元）', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'customer_service_fee';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'packaging_service_fee'))
    EXEC sp_addextendedproperty 'MS_Description', '包装服务费（元）', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'packaging_service_fee';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'settlement_amount'))
    EXEC sp_addextendedproperty 'MS_Description', '结算金额（元）', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'settlement_amount';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'bill_no'))
    EXEC sp_addextendedproperty 'MS_Description', '账单编号', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'bill_no';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'bill_period'))
    EXEC sp_addextendedproperty 'MS_Description', '账单起止时间', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'bill_period';
IF NOT EXISTS (SELECT * FROM fn_listextendedproperty('MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'ZH'))
    EXEC sp_addextendedproperty 'MS_Description', '店铺名称', 'SCHEMA', 'dbo', 'TABLE', 'dw_dzd_cargo_damage', 'COLUMN', 'ZH';
GO

PRINT '数据库迁移完成！新增 3 张补充 sheet 表：dw_dzd_other_fee, dw_dzd_deduction_detail, dw_dzd_cargo_damage';
GO
