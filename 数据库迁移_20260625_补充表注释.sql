-- ============================================================
-- 得物对账单系统 - 补充表注释（覆盖更新）
-- 版本: 2026-06-25
-- 功能: 给 dw_dwd_bill_records / dw_dwd_bill_records_copy1
--       添加详细表级和字段级中文注释
-- ============================================================

-- ============================================================
-- 1. dw_dwd_bill_records — API 拉取的账单列表
-- ============================================================
-- 先删除已有注释再重新添加
BEGIN TRY
    EXEC sp_dropextendedproperty 'MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', NULL, NULL;
END TRY BEGIN CATCH END CATCH
EXEC sp_addextendedproperty 'MS_Description', '【API账单列表】从得物开放平台 period_list 接口获取的账单记录，每账单一行。按 bill_no+店铺 去重，用于判断哪些账单需要触发 generate→export 下载流程。数据来源：API；消费者：download_files()', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records';
GO

-- 字段注释（逐个删除重建）
BEGIN TRY EXEC sp_dropextendedproperty 'MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'bill_no'; END TRY BEGIN CATCH END CATCH
EXEC sp_addextendedproperty 'MS_Description', '账单编号（唯一标识），格式: 起止日期-QY-店铺ID', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'bill_no';

BEGIN TRY EXEC sp_dropextendedproperty 'MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'bill_start_time'; END TRY BEGIN CATCH END CATCH
EXEC sp_addextendedproperty 'MS_Description', '账单起始日期', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'bill_start_time';

BEGIN TRY EXEC sp_dropextendedproperty 'MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'bill_end_time'; END TRY BEGIN CATCH END CATCH
EXEC sp_addextendedproperty 'MS_Description', '账单结束日期', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'bill_end_time';

BEGIN TRY EXEC sp_dropextendedproperty 'MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'settle_amount'; END TRY BEGIN CATCH END CATCH
EXEC sp_addextendedproperty 'MS_Description', '应结金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'settle_amount';

BEGIN TRY EXEC sp_dropextendedproperty 'MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'order_receive_amount'; END TRY BEGIN CATCH END CATCH
EXEC sp_addextendedproperty 'MS_Description', '实收金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'order_receive_amount';

BEGIN TRY EXEC sp_dropextendedproperty 'MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'platform_service_fee'; END TRY BEGIN CATCH END CATCH
EXEC sp_addextendedproperty 'MS_Description', '平台服务费', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'platform_service_fee';

BEGIN TRY EXEC sp_dropextendedproperty 'MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'seller_subsidies_amount'; END TRY BEGIN CATCH END CATCH
EXEC sp_addextendedproperty 'MS_Description', '卖家补贴金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'seller_subsidies_amount';

BEGIN TRY EXEC sp_dropextendedproperty 'MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'refund_amount'; END TRY BEGIN CATCH END CATCH
EXEC sp_addextendedproperty 'MS_Description', '退款金额', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'refund_amount';

BEGIN TRY EXEC sp_dropextendedproperty 'MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'status'; END TRY BEGIN CATCH END CATCH
EXEC sp_addextendedproperty 'MS_Description', '账单状态', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'status';

BEGIN TRY EXEC sp_dropextendedproperty 'MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'name'; END TRY BEGIN CATCH END CATCH
EXEC sp_addextendedproperty 'MS_Description', '店铺名称（对应 dewu_app_credentials.cred_id）', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records', 'COLUMN', 'name';
GO

-- ============================================================
-- 2. dw_dwd_bill_records_copy1 — 已入库账单判重
-- ============================================================
BEGIN TRY
    EXEC sp_dropextendedproperty 'MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records_copy1', NULL, NULL;
END TRY BEGIN CATCH END CATCH
EXEC sp_addextendedproperty 'MS_Description', '【已入库账单判重】存储已成功导入明细表（dw_dzd_xs/dw_dzd_thtk 等）的 bill_no。process_import 入库前先查此表，已存在的 bill_no 跳过，防止重复写入。此表是最后一道防重闸门', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records_copy1';
GO

BEGIN TRY EXEC sp_dropextendedproperty 'MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records_copy1', 'COLUMN', 'bill_no'; END TRY BEGIN CATCH END CATCH
EXEC sp_addextendedproperty 'MS_Description', '已入库的账单编号（唯一，判重依据）', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records_copy1', 'COLUMN', 'bill_no';

BEGIN TRY EXEC sp_dropextendedproperty 'MS_Description', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records_copy1', 'COLUMN', 'name'; END TRY BEGIN CATCH END CATCH
EXEC sp_addextendedproperty 'MS_Description', '账单所属店铺名称', 'SCHEMA', 'dbo', 'TABLE', 'dw_dwd_bill_records_copy1', 'COLUMN', 'name';
GO

PRINT '表注释更新完成！';
GO
