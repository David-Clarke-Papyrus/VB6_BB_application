--
-- Script To Update dbo.tCategoryStatsMonthly Table In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.tCategoryStatsMonthly Table'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_StockValue_CostEx')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_StockValue_CostEx]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_StockValue_SPInc')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_StockValue_SPInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_StockQty')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_StockQty]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_SalesValue_RetailInc')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_SalesValue_RetailInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_SalesValue_CostEx')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_SalesValue_CostEx]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_SalesQty')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_SalesQty]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_Last12MonthStockValue')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_Last12MonthStockValue]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_Last12MonthSalesValue')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_Last12MonthSalesValue]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_QtyMonthsInStockTurnRange')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_QtyMonthsInStockTurnRange]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_ReturnsQty')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_ReturnsQty]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_ReturnsValue_CostEx')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_ReturnsValue_CostEx]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_ReturnsValue_RetailInc')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_ReturnsValue_RetailInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_OrdersOSQty')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_OrdersOSQty]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_OrdersOSValue_CostEx')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_OrdersOSValue_CostEx]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_OrdersOSValue_RetailInc')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_OrdersOSValue_RetailInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_OrdersPlacedQty')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_OrdersPlacedQty]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_OrdersPlacedValue_CostEx')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_OrdersPlacedValue_CostEx]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_OrdersPlacedValue_RetailInc')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_OrdersPlacedValue_RetailInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_DELLsReceivedQty')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_DELLsReceivedQty]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_DELLsReceivedValue_CostEx')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_DELLsReceivedValue_CostEx]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_DELLSReceivedValue_RetailInc')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_DELLSReceivedValue_RetailInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_MissingLastStockTake_RetailInc')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_MissingLastStockTake_RetailInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_MissingLastStockTake_CostEx')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_MissingLastStockTake_CostEx]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_MissingLastStockTake_qty')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_MissingLastStockTake_qty]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_Margin')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_Margin]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_Qty')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_Qty]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_QtyInTop50')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_QtyInTop50]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_SalesAsPercentOfTotalSales_RetailInc')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_SalesAsPercentOfTotalSales_RetailInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_SalesAsPercentOfTotalSOH_RetailInc')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_SalesAsPercentOfTotalSOH_RetailInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_StockAsPercentOfTotalSOH_CostEx')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_StockAsPercentOfTotalSOH_CostEx]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_ReturnsAsPercentDeliveries')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_ReturnsAsPercentDeliveries]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tCategoryStatsMonthly_CATS_ReturnsAsPercentSales')
      ALTER TABLE [dbo].[tCategoryStatsMonthly] DROP CONSTRAINT [DF_tCategoryStatsMonthly_CATS_ReturnsAsPercentSales]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE TABLE [dbo].[tmp_tCategoryStatsMonthly] (
   [CATS_ID] [int] IDENTITY (1, 1) NOT NULL,
   [CATS_CATEGORYID] [int] NOT NULL,
   [CATS_Month] [datetime] NOT NULL,
   [CATS_StockValue_CostEx] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_StockValue_CostEx] DEFAULT ((0)),
   [CATS_StockValue_SPInc] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_StockValue_SPInc] DEFAULT ((0)),
   [CATS_StockValue_RRPInc] [real] NULL,
   [CATS_StockQty] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_StockQty] DEFAULT ((0)),
   [CATS_SalesValue_RetailInc] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_SalesValue_RetailInc] DEFAULT ((0)),
   [CATS_SalesValue_CostEx] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_SalesValue_CostEx] DEFAULT ((0)),
   [CATS_SalesValue_RRPInc] [real] NULL,
   [CATS_SalesQty] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_SalesQty] DEFAULT ((0)),
   [CATS_Last12MonthStockValue] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_Last12MonthStockValue] DEFAULT ((0)),
   [CATS_Last12MonthSalesValue] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_Last12MonthSalesValue] DEFAULT ((0)),
   [CATS_QtyMonthsInStockTurnRange] [int] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_QtyMonthsInStockTurnRange] DEFAULT ((0)),
   [CATS_ReturnsQty] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_ReturnsQty] DEFAULT ((0)),
   [CATS_ReturnsValue_CostEx] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_ReturnsValue_CostEx] DEFAULT ((0)),
   [CATS_ReturnsValue_RetailInc] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_ReturnsValue_RetailInc] DEFAULT ((0)),
   [CATS_OrdersOSQty] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_OrdersOSQty] DEFAULT ((0)),
   [CATS_OrdersOSValue_CostEx] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_OrdersOSValue_CostEx] DEFAULT ((0)),
   [CATS_OrdersOSValue_RetailInc] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_OrdersOSValue_RetailInc] DEFAULT ((0)),
   [CATS_OrdersPlacedQty] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_OrdersPlacedQty] DEFAULT ((0)),
   [CATS_OrdersPlacedValue_CostEx] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_OrdersPlacedValue_CostEx] DEFAULT ((0)),
   [CATS_OrdersPlacedValue_RetailInc] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_OrdersPlacedValue_RetailInc] DEFAULT ((0)),
   [CATS_DELLsReceivedQty] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_DELLsReceivedQty] DEFAULT ((0)),
   [CATS_DELLsReceivedValue_CostEx] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_DELLsReceivedValue_CostEx] DEFAULT ((0)),
   [CATS_DELLSReceivedValue_RetailInc] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_DELLSReceivedValue_RetailInc] DEFAULT ((0)),
   [CATS_MissingLastStockTake_RetailInc] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_MissingLastStockTake_RetailInc] DEFAULT ((0)),
   [CATS_MissingLastStockTake_CostEx] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_MissingLastStockTake_CostEx] DEFAULT ((0)),
   [CATS_MissingLastStockTake_qty] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_MissingLastStockTake_qty] DEFAULT ((0)),
   [CATS_Margin] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_Margin] DEFAULT ((0)),
   [CATS_Qty] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_Qty] DEFAULT ((0)),
   [CATS_QtyInTop50] [int] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_QtyInTop50] DEFAULT ((0)),
   [CATS_SalesAsPercentOfTotalSales_RetailInc] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_SalesAsPercentOfTotalSales_RetailInc] DEFAULT ((0)),
   [CATS_SalesAsPercentOfTotalSOH_RetailInc] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_SalesAsPercentOfTotalSOH_RetailInc] DEFAULT ((0)),
   [CATS_StockAsPercentOfTotalSOH_CostEx] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_StockAsPercentOfTotalSOH_CostEx] DEFAULT ((0)),
   [CATS_StockAsPercentOfTotalSOH_RRPInc] [real] NULL,
   [CATS_ReturnsAsPercentDeliveries] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_ReturnsAsPercentDeliveries] DEFAULT ((0)),
   [CATS_ReturnsAsPercentSales] [real] NULL CONSTRAINT [DF_tCategoryStatsMonthly_CATS_ReturnsAsPercentSales] DEFAULT ((0))
)
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   SET IDENTITY_INSERT [dbo].[tmp_tCategoryStatsMonthly] ON
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   INSERT INTO [dbo].[tmp_tCategoryStatsMonthly] ([CATS_ID], [CATS_CATEGORYID], [CATS_Month], [CATS_StockValue_CostEx], [CATS_StockValue_SPInc], [CATS_StockValue_RRPInc], [CATS_StockQty], [CATS_SalesValue_RetailInc], [CATS_SalesValue_CostEx], [CATS_SalesValue_RRPInc], [CATS_SalesQty], [CATS_Last12MonthStockValue], [CATS_Last12MonthSalesValue], [CATS_QtyMonthsInStockTurnRange], [CATS_ReturnsQty], [CATS_ReturnsValue_CostEx], [CATS_ReturnsValue_RetailInc], [CATS_OrdersOSQty], [CATS_OrdersOSValue_CostEx], [CATS_OrdersOSValue_RetailInc], [CATS_OrdersPlacedQty], [CATS_OrdersPlacedValue_CostEx], [CATS_OrdersPlacedValue_RetailInc], [CATS_DELLsReceivedQty], [CATS_DELLsReceivedValue_CostEx], [CATS_DELLSReceivedValue_RetailInc], [CATS_MissingLastStockTake_RetailInc], [CATS_MissingLastStockTake_CostEx], [CATS_MissingLastStockTake_qty], [CATS_Margin], [CATS_Qty], [CATS_QtyInTop50], [CATS_SalesAsPercentOfTotalSales_RetailInc], [CATS_SalesAsPercentOfTotalSOH_RetailInc], [CATS_StockAsPercentOfTotalSOH_CostEx], [CATS_StockAsPercentOfTotalSOH_RRPInc], [CATS_ReturnsAsPercentDeliveries], [CATS_ReturnsAsPercentSales])
   SELECT [CATS_ID], [CATS_CATEGORYID], [CATS_Month], [CATS_StockValue_CostEx], [CATS_StockValue_SPInc], NULL, [CATS_StockQty], [CATS_SalesValue_RetailInc], [CATS_SalesValue_CostEx], NULL, [CATS_SalesQty], [CATS_Last12MonthStockValue], [CATS_Last12MonthSalesValue], [CATS_QtyMonthsInStockTurnRange], [CATS_ReturnsQty], [CATS_ReturnsValue_CostEx], [CATS_ReturnsValue_RetailInc], [CATS_OrdersOSQty], [CATS_OrdersOSValue_CostEx], [CATS_OrdersOSValue_RetailInc], [CATS_OrdersPlacedQty], [CATS_OrdersPlacedValue_CostEx], [CATS_OrdersPlacedValue_RetailInc], [CATS_DELLsReceivedQty], [CATS_DELLsReceivedValue_CostEx], [CATS_DELLSReceivedValue_RetailInc], [CATS_MissingLastStockTake_RetailInc], [CATS_MissingLastStockTake_CostEx], [CATS_MissingLastStockTake_qty], [CATS_Margin], [CATS_Qty], [CATS_QtyInTop50], [CATS_SalesAsPercentOfTotalSales_RetailInc], [CATS_SalesAsPercentOfTotalSOH_RetailInc], [CATS_StockAsPercentOfTotalSOH_CostEx], NULL, [CATS_ReturnsAsPercentDeliveries], [CATS_ReturnsAsPercentSales]
   FROM [dbo].[tCategoryStatsMonthly]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   SET IDENTITY_INSERT [dbo].[tmp_tCategoryStatsMonthly] OFF
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   DROP TABLE [dbo].[tCategoryStatsMonthly]
GO

sp_rename N'[dbo].[tmp_tCategoryStatsMonthly]', N'tCategoryStatsMonthly'

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   ALTER TABLE [dbo].[tCategoryStatsMonthly] ADD CONSTRAINT [PK_tCategoryStatsMonthly] PRIMARY KEY CLUSTERED ([CATS_ID])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.tCategoryStatsMonthly Table Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.tCategoryStatsMonthly Table'
END
GO

--
-- Script To Create dbo.tSummaryStatsMonthly Table In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.tSummaryStatsMonthly Table'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO

CREATE TABLE [dbo].[tSummaryStatsMonthly] (
   [SUMM_ID] [int] IDENTITY (1, 1) NOT NULL,
   [SUMM_Month] [datetime] NOT NULL,
   [SUMM_StockValue_CostEx] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_StockValue_CostEx] DEFAULT ((0)),
   [SUMM_StockValue_SPInc] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_StockValue_SPInc] DEFAULT ((0)),
   [SUMM_StockValue_RRPInc] [real] NULL,
   [SUMM_StockQty] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_StockQty] DEFAULT ((0)),
   [SUMM_SalesValue_RetailInc] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_SalesValue_RetailInc] DEFAULT ((0)),
   [SUMM_SalesValue_CostEx] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_SalesValue_CostEx] DEFAULT ((0)),
   [SUMM_SalesValue_RRPInc] [real] NULL,
   [SUMM_SalesQty] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_SalesQty] DEFAULT ((0)),
   [SUMM_Last12MonthStockValue] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_Last12MonthStockValue] DEFAULT ((0)),
   [SUMM_Last12MonthSalesValue] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_Last12MonthSalesValue] DEFAULT ((0)),
   [SUMM_QtyMonthsInStockTurnRange] [int] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_QtyMonthsInStockTurnRange] DEFAULT ((0)),
   [SUMM_ReturnsQty] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_ReturnsQty] DEFAULT ((0)),
   [SUMM_ReturnsValue_CostEx] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_ReturnsValue_CostEx] DEFAULT ((0)),
   [SUMM_ReturnsValue_RetailInc] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_ReturnsValue_RetailInc] DEFAULT ((0)),
   [SUMM_OrdersOSQty] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_OrdersOSQty] DEFAULT ((0)),
   [SUMM_OrdersOSValue_CostEx] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_OrdersOSValue_CostEx] DEFAULT ((0)),
   [SUMM_OrdersOSValue_RetailInc] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_OrdersOSValue_RetailInc] DEFAULT ((0)),
   [SUMM_OrdersPlacedQty] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_OrdersPlacedQty] DEFAULT ((0)),
   [SUMM_OrdersPlacedValue_CostEx] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_OrdersPlacedValue_CostEx] DEFAULT ((0)),
   [SUMM_OrdersPlacedValue_RetailInc] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_OrdersPlacedValue_RetailInc] DEFAULT ((0)),
   [SUMM_DELLsReceivedQty] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_DELLsReceivedQty] DEFAULT ((0)),
   [SUMM_DELLsReceivedValue_CostEx] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_DELLsReceivedValue_CostEx] DEFAULT ((0)),
   [SUMM_DELLSReceivedValue_RetailInc] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_DELLSReceivedValue_RetailInc] DEFAULT ((0)),
   [SUMM_MissingLastStockTake_RetailInc] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_MissingLastStockTake_RetailInc] DEFAULT ((0)),
   [SUMM_MissingLastStockTake_CostEx] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_MissingLastStockTake_CostEx] DEFAULT ((0)),
   [SUMM_MissingLastStockTake_qty] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_MissingLastStockTake_qty] DEFAULT ((0)),
   [SUMM_Margin] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_Margin] DEFAULT ((0)),
   [SUMM_Qty] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_Qty] DEFAULT ((0)),
   [SUMM_QtyInTop50] [int] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_QtyInTop50] DEFAULT ((0)),
   [SUMM_SalesAsPercentOfTotalSales_RetailInc] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_SalesAsPercentOfTotalSales_RetailInc] DEFAULT ((0)),
   [SUMM_SalesAsPercentOfTotalSOH_RetailInc] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_SalesAsPercentOfTotalSOH_RetailInc] DEFAULT ((0)),
   [SUMM_StockAsPercentOfTotalSOH_CostEx] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_StockAsPercentOfTotalSOH_CostEx] DEFAULT ((0)),
   [SUMM_StockAsPercentOfTotalSOH_RRPInc] [real] NULL,
   [SUMM_ReturnsAsPercentDeliveries] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_ReturnsAsPercentDeliveries] DEFAULT ((0)),
   [SUMM_ReturnsAsPercentSales] [real] NULL CONSTRAINT [DF_tSummaryStatsMonthly_SUMM_ReturnsAsPercentSales] DEFAULT ((0))
)
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   ALTER TABLE [dbo].[tSummaryStatsMonthly] ADD CONSTRAINT [PK_tSummaryStatsMonthly] PRIMARY KEY CLUSTERED ([SUMM_ID])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.tSummaryStatsMonthly Table Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.tSummaryStatsMonthly Table'
END
GO

--
-- Script To Update dbo.tSupplierStatsMonthly Table In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.tSupplierStatsMonthly Table'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_StockValue_CostEx')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_StockValue_CostEx]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_StockValue_SPInc')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_StockValue_SPInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_StockQty')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_StockQty]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_SalesValue_RetailInc')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_SalesValue_RetailInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_SalesValue_CostEx')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_SalesValue_CostEx]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_SalesQty')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_SalesQty]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_Last12MonthStockValue')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_Last12MonthStockValue]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_Last12MonthSalesValue')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_Last12MonthSalesValue]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_QtyMonthsInStockTurnRange')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_QtyMonthsInStockTurnRange]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_ReturnsQty')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_ReturnsQty]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_ReturnsValue_CostEx')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_ReturnsValue_CostEx]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_ReturnsValue_RetailInc')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_ReturnsValue_RetailInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_OrdersOSQty')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_OrdersOSQty]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_OrdersOSValue_CostEx')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_OrdersOSValue_CostEx]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_OrdersOSValue_RetailInc')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_OrdersOSValue_RetailInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_OrdersPlacedQty')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_OrdersPlacedQty]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_OrdersPlacedValue_CostEx')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_OrdersPlacedValue_CostEx]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_OrdersPlacedValue_RetailInc')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_OrdersPlacedValue_RetailInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_DELLsReceivedQty')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_DELLsReceivedQty]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_DELLsReceivedValue_CostEx')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_DELLsReceivedValue_CostEx]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_DELLSReceivedValue_RetailInc')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_DELLSReceivedValue_RetailInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_MissingLastStockTake_RetailInc')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_MissingLastStockTake_RetailInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_MissingLastStockTake_CostEx')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_MissingLastStockTake_CostEx]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_MissingLastStockTake_qty')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_MissingLastStockTake_qty]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_Margin')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_Margin]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_Qty')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_Qty]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_QtyInTop50')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_QtyInTop50]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_SalesAsPercentOfTotalSales_RetailInc')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_SalesAsPercentOfTotalSales_RetailInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_SalesAsPercentOfTotalSOH_RetailInc')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_SalesAsPercentOfTotalSOH_RetailInc]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_StockAsPercentOfTotalSOH_CostEx')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_StockAsPercentOfTotalSOH_CostEx]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_ReturnsAsPercentDeliveries')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_ReturnsAsPercentDeliveries]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tSupplierStatsMonthly_SUPPS_ReturnsAsPercentSales')
      ALTER TABLE [dbo].[tSupplierStatsMonthly] DROP CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_ReturnsAsPercentSales]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE TABLE [dbo].[tmp_tSupplierStatsMonthly] (
   [SUPPS_ID] [int] IDENTITY (1, 1) NOT NULL,
   [SUPPS_SUPPLIERID] [int] NOT NULL,
   [SUPPS_Month] [datetime] NOT NULL,
   [SUPPS_StockValue_CostEx] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_StockValue_CostEx] DEFAULT ((0)),
   [SUPPS_StockValue_SPInc] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_StockValue_SPInc] DEFAULT ((0)),
   [SUPPS_StockValue_RRPInc] [real] NULL,
   [SUPPS_StockQty] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_StockQty] DEFAULT ((0)),
   [SUPPS_SalesValue_RetailInc] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_SalesValue_RetailInc] DEFAULT ((0)),
   [SUPPS_SalesValue_CostEx] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_SalesValue_CostEx] DEFAULT ((0)),
   [SUPPS_SalesValue_RRPInc] [real] NULL,
   [SUPPS_SalesQty] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_SalesQty] DEFAULT ((0)),
   [SUPPS_Last12MonthStockValue] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_Last12MonthStockValue] DEFAULT ((0)),
   [SUPPS_Last12MonthSalesValue] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_Last12MonthSalesValue] DEFAULT ((0)),
   [SUPPS_QtyMonthsInStockTurnRange] [int] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_QtyMonthsInStockTurnRange] DEFAULT ((0)),
   [SUPPS_ReturnsQty] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_ReturnsQty] DEFAULT ((0)),
   [SUPPS_ReturnsValue_CostEx] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_ReturnsValue_CostEx] DEFAULT ((0)),
   [SUPPS_ReturnsValue_RetailInc] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_ReturnsValue_RetailInc] DEFAULT ((0)),
   [SUPPS_OrdersOSQty] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_OrdersOSQty] DEFAULT ((0)),
   [SUPPS_OrdersOSValue_CostEx] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_OrdersOSValue_CostEx] DEFAULT ((0)),
   [SUPPS_OrdersOSValue_RetailInc] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_OrdersOSValue_RetailInc] DEFAULT ((0)),
   [SUPPS_OrdersPlacedQty] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_OrdersPlacedQty] DEFAULT ((0)),
   [SUPPS_OrdersPlacedValue_CostEx] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_OrdersPlacedValue_CostEx] DEFAULT ((0)),
   [SUPPS_OrdersPlacedValue_RetailInc] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_OrdersPlacedValue_RetailInc] DEFAULT ((0)),
   [SUPPS_DELLsReceivedQty] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_DELLsReceivedQty] DEFAULT ((0)),
   [SUPPS_DELLsReceivedValue_CostEx] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_DELLsReceivedValue_CostEx] DEFAULT ((0)),
   [SUPPS_DELLSReceivedValue_RetailInc] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_DELLSReceivedValue_RetailInc] DEFAULT ((0)),
   [SUPPS_MissingLastStockTake_RetailInc] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_MissingLastStockTake_RetailInc] DEFAULT ((0)),
   [SUPPS_MissingLastStockTake_CostEx] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_MissingLastStockTake_CostEx] DEFAULT ((0)),
   [SUPPS_MissingLastStockTake_qty] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_MissingLastStockTake_qty] DEFAULT ((0)),
   [SUPPS_Margin] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_Margin] DEFAULT ((0)),
   [SUPPS_Qty] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_Qty] DEFAULT ((0)),
   [SUPPS_QtyInTop50] [int] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_QtyInTop50] DEFAULT ((0)),
   [SUPPS_SalesAsPercentOfTotalSales_RetailInc] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_SalesAsPercentOfTotalSales_RetailInc] DEFAULT ((0)),
   [SUPPS_SalesAsPercentOfTotalSOH_RetailInc] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_SalesAsPercentOfTotalSOH_RetailInc] DEFAULT ((0)),
   [SUPPS_StockAsPercentOfTotalSOH_CostEx] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_StockAsPercentOfTotalSOH_CostEx] DEFAULT ((0)),
   [SUPPS_StockAsPercentOfTotalSOH_RRPInc] [real] NULL,
   [SUPPS_ReturnsAsPercentDeliveries] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_ReturnsAsPercentDeliveries] DEFAULT ((0)),
   [SUPPS_ReturnsAsPercentSales] [real] NULL CONSTRAINT [DF_tSupplierStatsMonthly_SUPPS_ReturnsAsPercentSales] DEFAULT ((0))
)
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   SET IDENTITY_INSERT [dbo].[tmp_tSupplierStatsMonthly] ON
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   INSERT INTO [dbo].[tmp_tSupplierStatsMonthly] ([SUPPS_ID], [SUPPS_SUPPLIERID], [SUPPS_Month], [SUPPS_StockValue_CostEx], [SUPPS_StockValue_SPInc], [SUPPS_StockValue_RRPInc], [SUPPS_StockQty], [SUPPS_SalesValue_RetailInc], [SUPPS_SalesValue_CostEx], [SUPPS_SalesValue_RRPInc], [SUPPS_SalesQty], [SUPPS_Last12MonthStockValue], [SUPPS_Last12MonthSalesValue], [SUPPS_QtyMonthsInStockTurnRange], [SUPPS_ReturnsQty], [SUPPS_ReturnsValue_CostEx], [SUPPS_ReturnsValue_RetailInc], [SUPPS_OrdersOSQty], [SUPPS_OrdersOSValue_CostEx], [SUPPS_OrdersOSValue_RetailInc], [SUPPS_OrdersPlacedQty], [SUPPS_OrdersPlacedValue_CostEx], [SUPPS_OrdersPlacedValue_RetailInc], [SUPPS_DELLsReceivedQty], [SUPPS_DELLsReceivedValue_CostEx], [SUPPS_DELLSReceivedValue_RetailInc], [SUPPS_MissingLastStockTake_RetailInc], [SUPPS_MissingLastStockTake_CostEx], [SUPPS_MissingLastStockTake_qty], [SUPPS_Margin], [SUPPS_Qty], [SUPPS_QtyInTop50], [SUPPS_SalesAsPercentOfTotalSales_RetailInc], [SUPPS_SalesAsPercentOfTotalSOH_RetailInc], [SUPPS_StockAsPercentOfTotalSOH_CostEx], [SUPPS_StockAsPercentOfTotalSOH_RRPInc], [SUPPS_ReturnsAsPercentDeliveries], [SUPPS_ReturnsAsPercentSales])
   SELECT [SUPPS_ID], [SUPPS_SUPPLIERID], [SUPPS_Month], [SUPPS_StockValue_CostEx], [SUPPS_StockValue_SPInc], NULL, [SUPPS_StockQty], [SUPPS_SalesValue_RetailInc], [SUPPS_SalesValue_CostEx], NULL, [SUPPS_SalesQty], [SUPPS_Last12MonthStockValue], [SUPPS_Last12MonthSalesValue], [SUPPS_QtyMonthsInStockTurnRange], [SUPPS_ReturnsQty], [SUPPS_ReturnsValue_CostEx], [SUPPS_ReturnsValue_RetailInc], [SUPPS_OrdersOSQty], [SUPPS_OrdersOSValue_CostEx], [SUPPS_OrdersOSValue_RetailInc], [SUPPS_OrdersPlacedQty], [SUPPS_OrdersPlacedValue_CostEx], [SUPPS_OrdersPlacedValue_RetailInc], [SUPPS_DELLsReceivedQty], [SUPPS_DELLsReceivedValue_CostEx], [SUPPS_DELLSReceivedValue_RetailInc], [SUPPS_MissingLastStockTake_RetailInc], [SUPPS_MissingLastStockTake_CostEx], [SUPPS_MissingLastStockTake_qty], [SUPPS_Margin], [SUPPS_Qty], [SUPPS_QtyInTop50], [SUPPS_SalesAsPercentOfTotalSales_RetailInc], [SUPPS_SalesAsPercentOfTotalSOH_RetailInc], [SUPPS_StockAsPercentOfTotalSOH_CostEx], NULL, [SUPPS_ReturnsAsPercentDeliveries], [SUPPS_ReturnsAsPercentSales]
   FROM [dbo].[tSupplierStatsMonthly]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   SET IDENTITY_INSERT [dbo].[tmp_tSupplierStatsMonthly] OFF
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   DROP TABLE [dbo].[tSupplierStatsMonthly]
GO

sp_rename N'[dbo].[tmp_tSupplierStatsMonthly]', N'tSupplierStatsMonthly'

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   ALTER TABLE [dbo].[tSupplierStatsMonthly] ADD CONSTRAINT [PK_tSupplierStatsMonthly] PRIMARY KEY CLUSTERED ([SUPPS_ID])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.tSupplierStatsMonthly Table Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.tSupplierStatsMonthly Table'
END
GO

--
-- Script To Create dbo.CalcExtAddVAT2 Function In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.CalcExtAddVAT2 Function'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE Function [dbo].[CalcExtAddVAT2] (@Qty INTEGER,@PRICE REAL,@DISCOUNT REAL=0 ,@VATRATE NUMERIC(8,4)=0,@CURRDIVISOR INT = 1) 
    RETURNS REAL as 
BEGIN 
DECLARE @RET REAL

	SELECT @RET =((CAST( ISNULL(@QTY,0) AS REAL) * (ISNULL(@PRICE,0) * (CAST(100.00000 - ISNULL(@DISCOUNT,0) AS REAL)/100)) )
		 *   (CAST(100.00000+ISNULL(@VATRATE,0) AS REAL)/100))/@CURRDIVISOR
	RETURN @RET
END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.CalcExtAddVAT2 Function Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.CalcExtAddVAT2 Function'
END
GO

--
-- Script To Create dbo.FormatQtyOH_1 Function In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.FormatQtyOH_1 Function'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE FUNCTION [dbo].[FormatQtyOH_1](@Amt as INT)

RETURNS VARCHAR(20)
AS
BEGIN
	-- Declare the return variable here
DECLARE @RESULT VARCHAR(20)
	If @AMT <=0
		SELECT @RESULT = ''out of stock''
	Else
	If @Amt <  11
		SELECT @Result = ''in stock - low''
	Else
		SELECT @RESULT = CAST(@Amt as VARCHAR(10))
	-- Return the result of the function
	RETURN @RESULT

END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.FormatQtyOH_1 Function Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.FormatQtyOH_1 Function'
END
GO

--
-- Script To Create dbo.ufn_parsefind Function In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.ufn_parsefind Function'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('/*

Copyright © 2009 - John Burnette -- All Rights Reserved

*/

CREATE FUNCTION dbo.ufn_parsefind
(
@EntString varchar(max),
@Delimiter varchar(10),
@Occurrence bigint
)
RETURNS varchar(max)

AS

BEGIN

    DECLARE @CurString varchar(max)
    DECLARE @Pos bigint
    DECLARE @Loop bigint

    -- REQUIRE DELIMITER AT END OF STRING
    IF right(@EntString,1)<>@Delimiter
    SET @EntString = @EntString + @Delimiter

    -- ESTABLISH CORRECT SYNTAX FOR DELIMITER IN PATINDEX FUNCTION
    SET @Delimiter = ''%'' + @Delimiter + ''%''

    SET @Loop = 1
    SET @Pos = patindex(@Delimiter, @EntString)

    -- LOOP THROUGH IF DELIMTERS FOUND
    IF @Pos = 0
    BEGIN
        SET @CurString = Null
 END
    ELSE
    BEGIN
        WHILE @Loop <= @Occurrence and @Pos <> 0 
        BEGIN
            SET @Pos = patindex(@Delimiter, @EntString)
            SET @CurString = left(@EntString,@Pos-1)
            SET @EntString = right(@EntString,len(@EntString)-len(@CurString)-1)
            SET @Loop = @Loop + 1
        END 
    END

    -- DEFAULT A NULL FOR BLANK VALUES
    IF isnull(@CurString,'''')='''' or len(@CurString)<1
    SET @CurString = NULL
    
    -- RETURN VALUE
    RETURN @CurString

END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.ufn_parsefind Function Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.ufn_parsefind Function'
END
GO

--
-- Script To Update dbo.vCategoryPerformance View In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.vCategoryPerformance View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER VIEW dbo.vCategoryPerformance
AS
SELECT     TOP (100) PERCENT dbo.tCategoryStatsMonthly.CATS_Month, dbo.tCategoryStatsMonthly.CATS_CATEGORYID AS CategoryID, 
                      dbo.tCategoryStatsMonthly.CATS_StockValue_SPInc, dbo.tCategoryStatsMonthly.CATS_StockValue_CostEx, 
                      dbo.tCategoryStatsMonthly.CATS_StockValue_RRPInc, dbo.tCategoryStatsMonthly.CATS_SalesValue_RetailInc, 
                      dbo.tCategoryStatsMonthly.CATS_Margin, dbo.tCategoryStatsMonthly.CATS_Qty, dbo.tCategoryStatsMonthly.CATS_QtyInTop50, 
                      dbo.tCategoryStatsMonthly.CATS_SalesAsPercentOfTotalSales_RetailInc, dbo.tCategoryStatsMonthly.CATS_SalesAsPercentOfTotalSOH_RetailInc, 
                      dbo.tCategoryStatsMonthly.CATS_StockAsPercentOfTotalSOH_CostEx, dbo.tCategoryStatsMonthly.CATS_StockAsPercentOfTotalSOH_RRPInc, 
                      dbo.tCategoryStatsMonthly.CATS_ReturnsValue_RetailInc, dbo.tCategoryStatsMonthly.CATS_ReturnsValue_CostEx, 
                      dbo.tCategoryStatsMonthly.CATS_ReturnsAsPercentDeliveries, dbo.tCategoryStatsMonthly.CATS_ReturnsAsPercentSales, 
                      dbo.tCategoryStatsMonthly.CATS_OrdersPlacedQty, dbo.tCategoryStatsMonthly.CATS_OrdersPlacedValue_RetailInc, 
                      dbo.tCategoryStatsMonthly.CATS_OrdersPlacedValue_CostEx, dbo.tCategoryStatsMonthly.CATS_OrdersOSValue_RetailInc, 
                      dbo.tCategoryStatsMonthly.CATS_OrdersOSValue_CostEx, dbo.tCategoryStatsMonthly.CATS_MissingLastStockTake_RetailInc, 
                      dbo.tCategoryStatsMonthly.CATS_MissingLastStockTake_CostEx, dbo.tCategoryStatsMonthly.CATS_MissingLastStockTake_qty, 
                      dbo.tCategoryStatsMonthly.CATS_DELLsReceivedQty, dbo.tCategoryStatsMonthly.CATS_DELLsReceivedValue_CostEx, 
                      dbo.tCategoryStatsMonthly.CATS_DELLSReceivedValue_RetailInc, ISNULL(dbo.tDict.DICT_Desc, ''Unknown category'') AS CategoryName, 
                      dbo.tCategoryStatsMonthly.CATS_Last12MonthSalesValue, dbo.tCategoryStatsMonthly.CATS_Last12MonthStockValue, 
                      CASE WHEN ISNULL(CATS_StockValue_SPInc, 0) = 0 OR
                      ISNULL(CATS_QtyMonthsInStockTurnRange, 0) = 0 OR
                      CATS_Last12MonthStockValue = 0 THEN 0 ELSE CATS_Last12MonthSalesValue * (12 / CATS_QtyMonthsInStockTurnRange) 
                      / (CATS_Last12MonthStockValue / CATS_QtyMonthsInStockTurnRange) END AS StockTurn, 
                      dbo.tCategoryStatsMonthly.CATS_QtyMonthsInStockTurnRange
FROM         dbo.tCategoryStatsMonthly LEFT OUTER JOIN
                      dbo.tDict ON dbo.tCategoryStatsMonthly.CATS_CATEGORYID = dbo.tDict.DICT_ID')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vCategoryPerformance View Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.vCategoryPerformance View'
END
GO

--
-- Script To Update dbo.vCategoryPerformance_Pivot View In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.vCategoryPerformance_Pivot View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER VIEW dbo.vCategoryPerformance_Pivot
AS
SELECT     TOP (100) PERCENT CategoryName, CategoryID, h12, h11, h10, h09, h08, h07, h06, h05, h04, h03, h02, h01, m12_1, m11_1, m10_1, m09_1, m08_1, 
                      m07_1, m06_1, m05_1, m04_1, m03_1, m02_1, m01_1, m12_2, m11_2, m10_2, m09_2, m08_2, m07_2, m06_2, m05_2, m04_2, m03_2, m02_2, 
                      m01_2, m12_3, m11_3, m10_3, m09_3, m08_3, m07_3, m06_3, m05_3, m04_3, m03_3, m02_3, m01_3, m12_4, m11_4, m10_4, m09_4, m08_4, 
                      m07_4, m06_4, m05_4, m04_4, m03_4, m02_4, m01_4, m12_5, m11_5, m10_5, m09_5, m08_5, m07_5, m06_5, m05_5, m04_5, m03_5, m02_5, 
                      m01_5, m12_6, m11_6, m10_6, m09_6, m08_6, m07_6, m06_6, m05_6, m04_6, m03_6, m02_6, m01_6, m12_7, m11_7, m10_7, m09_7, m08_7, 
                      m07_7, m06_7, m05_7, m04_7, m03_7, m02_7, m01_7, m12_8, m11_8, m10_8, m09_8, m08_8, m07_8, m06_8, m05_8, m04_8, m03_8, m02_8, 
                      m01_8, m12_9, m11_9, m10_9, m09_9, m08_9, m07_9, m06_9, m05_9, m04_9, m03_9, m02_9, m01_9, m12_10, m11_10, m10_10, m09_10, m08_10, 
                      m07_10, m06_10, m05_10, m04_10, m03_10, m02_10, m01_10, m12_11, m11_11, m10_11, m09_11, m08_11, m07_11, m06_11, m05_11, m04_11, 
                      m03_11, m02_11, m01_11, m12_12, m11_12, m10_12, m09_12, m08_12, m07_12, m06_12, m05_12, m04_12, m03_12, m02_12, m01_12, m12_13, 
                      m11_13, m10_13, m09_13, m08_13, m07_13, m06_13, m05_13, m04_13, m03_13, m02_13, m01_13, m12_14, m11_14, m10_14, m09_14, m08_14, 
                      m07_14, m06_14, m05_14, m04_14, m03_14, m02_14, m01_14
FROM         (SELECT     TOP (100) PERCENT CAST(CategoryName AS VARCHAR(100)) AS CategoryName, CAST(CategoryID AS INT) AS CategoryID, DATEADD(m, - 0, 
                                              dbo.StartOfMonth(GETDATE())) AS h12, DATEADD(m, - 1, dbo.StartOfMonth(GETDATE())) AS h11, DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GETDATE())) AS h10, DATEADD(m, - 3, dbo.StartOfMonth(GETDATE())) AS h09, DATEADD(m, - 4, 
                                              dbo.StartOfMonth(GETDATE())) AS h08, DATEADD(m, - 5, dbo.StartOfMonth(GETDATE())) AS h07, DATEADD(m, - 6, 
                                              dbo.StartOfMonth(GETDATE())) AS h06, DATEADD(m, - 7, dbo.StartOfMonth(GETDATE())) AS h05, DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GETDATE())) AS h04, DATEADD(m, - 9, dbo.StartOfMonth(GETDATE())) AS h03, DATEADD(m, - 10, 
                                              dbo.StartOfMonth(GETDATE())) AS h02, DATEADD(m, - 11, dbo.StartOfMonth(GETDATE())) AS h01, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_QtyInTop50, 0) ELSE 0 END) 
                                              AS m12_1, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_QtyInTop50, 0) 
                                              ELSE 0 END) AS m11_1, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_QtyInTop50, 0) ELSE 0 END) AS m10_1, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 3, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_QtyInTop50, 0) ELSE 0 END) AS m09_1, SUM(CASE X.CATS_MONTH WHEN DATEADD(m,
                                               - 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_QtyInTop50, 0) ELSE 0 END) AS m08_1, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_QtyInTop50, 0) ELSE 0 END) 
                                              AS m07_1, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_QtyInTop50, 0) 
                                              ELSE 0 END) AS m06_1, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_QtyInTop50, 0) ELSE 0 END) AS m05_1, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_QtyInTop50, 0) ELSE 0 END) AS m04_1, SUM(CASE X.CATS_MONTH WHEN DATEADD(m,
                                               - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_QtyInTop50, 0) ELSE 0 END) AS m03_1, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_QtyInTop50, 0) ELSE 0 END) 
                                              AS m02_1, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 11, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_QtyInTop50, 0) 
                                              ELSE 0 END) AS m01_1, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_Salesvalue_RetailInc, 0) ELSE 0 END) AS m12_2, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 1, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_Salesvalue_RetailInc, 0) ELSE 0 END) AS m11_2, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_Salesvalue_RetailInc, 0) 
                                              ELSE 0 END) AS m10_2, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_Salesvalue_RetailInc, 0) ELSE 0 END) AS m09_2, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 4, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_Salesvalue_RetailInc, 0) ELSE 0 END) AS m08_2, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_Salesvalue_RetailInc, 0) 
                                              ELSE 0 END) AS m07_2, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_Salesvalue_RetailInc, 0) ELSE 0 END) AS m06_2, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 7, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_Salesvalue_RetailInc, 0) ELSE 0 END) AS m05_2, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_Salesvalue_RetailInc, 0) 
                                              ELSE 0 END) AS m04_2, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_Salesvalue_RetailInc, 0) ELSE 0 END) AS m03_2, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 10, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_Salesvalue_RetailInc, 0) ELSE 0 END) AS m02_2, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 11, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_Salesvalue_RetailInc, 0) 
                                              ELSE 0 END) AS m01_2, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m12_3, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, 
                                              - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m11_3, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m10_3, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, 
                                              - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m09_3, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m08_3, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, 
                                              - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m07_3, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m06_3, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, 
                                              - 7, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m05_3, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m04_3, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, 
                                              - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m03_3, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m02_3, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, 
                                              - 11, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m01_3, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m12_4, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, 
                                              - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m11_4, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m10_4, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, 
                                              - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m09_4, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m08_4, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, 
                                              - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m07_4, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m06_4, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, 
                                              - 7, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m05_4, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m04_4, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, 
                                              - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m03_4, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m02_4, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, 
                                              - 11, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m01_4, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_StockValue_RRPINC, 0) 
                                              ELSE 0 END) AS m12_5, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_StockValue_CostEx, 0) ELSE 0 END) AS m11_5, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_StockValue_CostEx, 0) ELSE 0 END) AS m10_5, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_StockValue_CostEx, 0) ELSE 0 END)
                                               AS m09_5, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_StockValue_CostEx, 0) 
                                              ELSE 0 END) AS m08_5, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_StockValue_CostEx, 0) ELSE 0 END) AS m07_5, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 6, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_StockValue_CostEx, 0) ELSE 0 END) AS m06_5, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_StockValue_CostEx, 0) ELSE 0 END)
                                               AS m05_5, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_StockValue_CostEx, 0) 
                                              ELSE 0 END) AS m04_5, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_StockValue_CostEx, 0) ELSE 0 END) AS m03_5, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 10, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_StockValue_CostEx, 0) ELSE 0 END) AS m02_5, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 11, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_StockValue_CostEx, 0) 
                                              ELSE 0 END) AS m01_5, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m12_6, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 1, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m11_6, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m10_6, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 3, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m09_6, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m08_6, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m07_6, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m06_6, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 7, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m05_6, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m04_6, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 9, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m03_6, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.CATS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m02_6, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 11,
                                               dbo.StartOfMonth(GetDate())) THEN ISNULL(X.CATS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m01_6, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m12_7, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m11_7, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m10_7, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m09_7, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m08_7, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m07_7, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m06_7, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m05_7, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m04_7, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m03_7, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m02_7, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 11, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m01_7, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsValue_RetailInc, 0) 
                                              ELSE 0 END) AS m12_8, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m11_8, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m10_8, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsValue_RetailInc, 0) 
                                              ELSE 0 END) AS m09_8, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m08_8, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m07_8, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsValue_RetailInc, 0) 
                                              ELSE 0 END) AS m06_8, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m05_8, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m04_8, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsValue_RetailInc, 0) 
                                              ELSE 0 END) AS m03_8, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m02_8, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m01_8, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsAsPercentDeliveries, 0) 
                                              ELSE 0 END) AS m12_9, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m11_9, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m10_9, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsAsPercentDeliveries, 0) 
                                              ELSE 0 END) AS m09_9, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m08_9, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m07_9, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsAsPercentDeliveries, 0) 
                                              ELSE 0 END) AS m06_9, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m05_9, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m04_9, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsAsPercentDeliveries, 0) 
                                              ELSE 0 END) AS m03_9, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m02_9, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m01_9, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsAsPercentSales, 0) 
                                              ELSE 0 END) AS m12_10, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_ReturnsAsPercentSales, 0) ELSE 0 END) AS m11_10, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsAsPercentSales, 0) ELSE 0 END) AS m10_10, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsAsPercentSales, 0) 
                                              ELSE 0 END) AS m09_10, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_ReturnsAsPercentSales, 0) ELSE 0 END) AS m08_10, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsAsPercentSales, 0) ELSE 0 END) AS m07_10, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsAsPercentSales, 0) 
                                              ELSE 0 END) AS m06_10, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_ReturnsAsPercentSales, 0) ELSE 0 END) AS m05_10, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsAsPercentSales, 0) ELSE 0 END) AS m04_10, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsAsPercentSales, 0) 
                                              ELSE 0 END) AS m03_10, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_ReturnsAsPercentSales, 0) ELSE 0 END) AS m02_10, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_ReturnsAsPercentSales, 0) ELSE 0 END) AS m01_10, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_Margin, 0) ELSE 0 END) AS m12_11, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_Margin, 0) ELSE 0 END) AS m11_11, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_Margin, 0) ELSE 0 END) AS m10_11, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_Margin, 0) ELSE 0 END) AS m09_11, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_Margin, 0) ELSE 0 END) AS m08_11, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_Margin, 0) ELSE 0 END) AS m07_11, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_Margin, 0) ELSE 0 END) AS m06_11, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_Margin, 0) ELSE 0 END) AS m05_11, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_Margin, 0) ELSE 0 END) AS m04_11, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_Margin, 0) ELSE 0 END) AS m03_11, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_Margin, 0) ELSE 0 END) AS m02_11, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 11, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_Margin, 0) ELSE 0 END) AS m01_11, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_OrdersPlacedValue_RetailInc, 0) 
                                              ELSE 0 END) AS m12_12, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_OrdersPlacedValue_RetailInc, 0) ELSE 0 END) AS m11_12, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_OrdersPlacedValue_RetailInc, 0) ELSE 0 END) AS m10_12, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_OrdersPlacedValue_RetailInc, 0) 
                                              ELSE 0 END) AS m09_12, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_OrdersPlacedValue_RetailInc, 0) ELSE 0 END) AS m08_12, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_OrdersPlacedValue_RetailInc, 0) ELSE 0 END) AS m07_12, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_OrdersPlacedValue_RetailInc, 0) 
                                              ELSE 0 END) AS m06_12, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_OrdersPlacedValue_RetailInc, 0) ELSE 0 END) AS m05_12, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_OrdersPlacedValue_RetailInc, 0) ELSE 0 END) AS m04_12, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_OrdersPlacedValue_RetailInc, 0) 
                                              ELSE 0 END) AS m03_12, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_OrdersPlacedValue_RetailInc, 0) ELSE 0 END) AS m02_12, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_OrdersPlacedValue_RetailInc, 0) ELSE 0 END) AS m01_12, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_OrdersOSValue_CostEx, 0) 
                                              ELSE 0 END) AS m12_13, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m11_13, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m10_13, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_OrdersOSValue_CostEx, 0) 
                                              ELSE 0 END) AS m09_13, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m08_13, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m07_13, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_OrdersOSValue_CostEx, 0) 
                                              ELSE 0 END) AS m06_13, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m05_13, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m04_13, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_OrdersOSValue_CostEx, 0) 
                                              ELSE 0 END) AS m03_13, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m02_13, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m01_13, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_MissingLastStockTake_RetailInc, 0) 
                                              ELSE 0 END) AS m12_14, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m11_14, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m10_14, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_MissingLastStockTake_RetailInc, 0) 
                                              ELSE 0 END) AS m09_14, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m08_14, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m07_14, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_MissingLastStockTake_RetailInc, 0) 
                                              ELSE 0 END) AS m06_14, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m05_14, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m04_14, 
                                              SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_MissingLastStockTake_RetailInc, 0) 
                                              ELSE 0 END) AS m03_14, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(CATS_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m02_14, SUM(CASE X.CATS_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(CATS_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m01_14
                       FROM          dbo.vCategoryPerformance AS X
                       GROUP BY CategoryName, CategoryID
                       ORDER BY CategoryName) AS Z
ORDER BY CategoryName')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vCategoryPerformance_Pivot View Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.vCategoryPerformance_Pivot View'
END
GO

--
-- Script To Update dbo.vPerformance_2 View In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.vPerformance_2 View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER VIEW dbo.vPerformance_2
AS
SELECT     TOP (100) PERCENT dbo.GetMonthYear(dbo.vSalesSummary.Dte) AS Mth, dbo.vSalesSummary.SUPPID, dbo.vSalesSummary.PID, 
                      dbo.vSalesSummary.MainSectionID
FROM         dbo.vSalesSummary LEFT OUTER JOIN
                      dbo.tTP AS tTP_1 ON dbo.vSalesSummary.SUPPID = tTP_1.TP_ID
GROUP BY dbo.GetMonthYear(dbo.vSalesSummary.Dte), dbo.vSalesSummary.PID, dbo.vSalesSummary.SUPPID, dbo.vSalesSummary.MainSectionID')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vPerformance_2 View Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.vPerformance_2 View'
END
GO

--
-- Script To Update dbo.vPerformance_StockByCategory View In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.vPerformance_StockByCategory View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER VIEW dbo.vPerformance_StockByCategory
AS
SELECT     dbo.tCategoryStatsMonthly.CATS_StockValue_CostEx, dbo.tCategoryStatsMonthly.CATS_StockValue_SPInc, dbo.tCategoryStatsMonthly.CATS_Month, 
                      dbo.tCategoryStatsMonthly.CATS_CATEGORYID, dbo.tCategoryStatsMonthly.CATS_StockQty, dbo.tCategoryStatsMonthly.CATS_ReturnsQty, 
                      dbo.tCategoryStatsMonthly.CATS_ReturnsValue_CostEx, dbo.tCategoryStatsMonthly.CATS_ReturnsValue_RetailInc, 
                      dbo.tCategoryStatsMonthly.CATS_OrdersOSQty, dbo.tCategoryStatsMonthly.CATS_OrdersOSValue_CostEx, 
                      dbo.tCategoryStatsMonthly.CATS_OrdersOSValue_RetailInc, dbo.tCategoryStatsMonthly.CATS_OrdersPlacedQty, 
                      dbo.tCategoryStatsMonthly.CATS_OrdersPlacedValue_CostEx, dbo.tCategoryStatsMonthly.CATS_OrdersPlacedValue_RetailInc, 
                      dbo.tCategoryStatsMonthly.CATS_DELLsReceivedQty, dbo.tCategoryStatsMonthly.CATS_DELLsReceivedValue_CostEx, 
                      dbo.tCategoryStatsMonthly.CATS_DELLSReceivedValue_RetailInc, dbo.tCategoryStatsMonthly.CATS_MissingLastStockTake_RetailInc, 
                      dbo.tCategoryStatsMonthly.CATS_MissingLastStockTake_CostEx, dbo.tCategoryStatsMonthly.CATS_MissingLastStockTake_qty, 
                      dbo.tCategoryStatsMonthly.CATS_StockValue_RRPInc
FROM         dbo.tCurrency INNER JOIN
                      dbo.tConfiguration ON dbo.tCurrency.CURR_ID = dbo.tConfiguration.CF_DefaultCurrID CROSS JOIN
                      dbo.tCategoryStatsMonthly
GROUP BY dbo.tCategoryStatsMonthly.CATS_StockValue_CostEx, dbo.tCategoryStatsMonthly.CATS_StockValue_SPInc, dbo.tCategoryStatsMonthly.CATS_Month, 
                      dbo.tCategoryStatsMonthly.CATS_CATEGORYID, dbo.tCategoryStatsMonthly.CATS_StockQty, dbo.tCategoryStatsMonthly.CATS_ReturnsQty, 
                      dbo.tCategoryStatsMonthly.CATS_ReturnsValue_CostEx, dbo.tCategoryStatsMonthly.CATS_ReturnsValue_RetailInc, 
                      dbo.tCategoryStatsMonthly.CATS_OrdersOSQty, dbo.tCategoryStatsMonthly.CATS_OrdersOSValue_CostEx, 
                      dbo.tCategoryStatsMonthly.CATS_OrdersOSValue_RetailInc, dbo.tCategoryStatsMonthly.CATS_OrdersPlacedQty, 
                      dbo.tCategoryStatsMonthly.CATS_OrdersPlacedValue_CostEx, dbo.tCategoryStatsMonthly.CATS_OrdersPlacedValue_RetailInc, 
                      dbo.tCategoryStatsMonthly.CATS_DELLsReceivedQty, dbo.tCategoryStatsMonthly.CATS_DELLsReceivedValue_CostEx, 
                      dbo.tCategoryStatsMonthly.CATS_DELLSReceivedValue_RetailInc, dbo.tCategoryStatsMonthly.CATS_MissingLastStockTake_RetailInc, 
                      dbo.tCategoryStatsMonthly.CATS_MissingLastStockTake_CostEx, dbo.tCategoryStatsMonthly.CATS_MissingLastStockTake_qty, 
                      dbo.tCategoryStatsMonthly.CATS_StockValue_RRPInc')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vPerformance_StockByCategory View Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.vPerformance_StockByCategory View'
END
GO

--
-- Script To Update dbo.vPerformance_StockBySupplier View In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.vPerformance_StockBySupplier View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER VIEW dbo.vPerformance_StockBySupplier
AS
SELECT     dbo.tSupplierStatsMonthly.SUPPS_StockValue_CostEx, dbo.tSupplierStatsMonthly.SUPPS_StockValue_SPInc, 
                      dbo.tSupplierStatsMonthly.SUPPS_StockValue_RRPInc, dbo.tSupplierStatsMonthly.SUPPS_Month, dbo.tSupplierStatsMonthly.SUPPS_SUPPLIERID, 
                      dbo.tSupplierStatsMonthly.SUPPS_StockQty, dbo.tSupplierStatsMonthly.SUPPS_ReturnsQty, dbo.tSupplierStatsMonthly.SUPPS_ReturnsValue_CostEx, 
                      dbo.tSupplierStatsMonthly.SUPPS_ReturnsValue_RetailInc, dbo.tSupplierStatsMonthly.SUPPS_OrdersOSQty, 
                      dbo.tSupplierStatsMonthly.SUPPS_OrdersOSValue_CostEx, dbo.tSupplierStatsMonthly.SUPPS_OrdersOSValue_RetailInc, 
                      dbo.tSupplierStatsMonthly.SUPPS_OrdersPlacedQty, dbo.tSupplierStatsMonthly.SUPPS_OrdersPlacedValue_CostEx, 
                      dbo.tSupplierStatsMonthly.SUPPS_OrdersPlacedValue_RetailInc, dbo.tSupplierStatsMonthly.SUPPS_DELLsReceivedQty, 
                      dbo.tSupplierStatsMonthly.SUPPS_DELLsReceivedValue_CostEx, dbo.tSupplierStatsMonthly.SUPPS_DELLSReceivedValue_RetailInc, 
                      dbo.tSupplierStatsMonthly.SUPPS_MissingLastStockTake_RetailInc, dbo.tSupplierStatsMonthly.SUPPS_MissingLastStockTake_CostEx, 
                      dbo.tSupplierStatsMonthly.SUPPS_MissingLastStockTake_qty
FROM         dbo.tCurrency INNER JOIN
                      dbo.tConfiguration ON dbo.tCurrency.CURR_ID = dbo.tConfiguration.CF_DefaultCurrID CROSS JOIN
                      dbo.tSupplierStatsMonthly
GROUP BY dbo.tSupplierStatsMonthly.SUPPS_StockValue_CostEx, dbo.tSupplierStatsMonthly.SUPPS_StockValue_SPInc, 
                      dbo.tSupplierStatsMonthly.SUPPS_Month, dbo.tSupplierStatsMonthly.SUPPS_SUPPLIERID, dbo.tSupplierStatsMonthly.SUPPS_StockQty, 
                      dbo.tSupplierStatsMonthly.SUPPS_ReturnsQty, dbo.tSupplierStatsMonthly.SUPPS_ReturnsValue_CostEx, 
                      dbo.tSupplierStatsMonthly.SUPPS_ReturnsValue_RetailInc, dbo.tSupplierStatsMonthly.SUPPS_OrdersOSQty, 
                      dbo.tSupplierStatsMonthly.SUPPS_OrdersOSValue_CostEx, dbo.tSupplierStatsMonthly.SUPPS_OrdersOSValue_RetailInc, 
                      dbo.tSupplierStatsMonthly.SUPPS_OrdersPlacedQty, dbo.tSupplierStatsMonthly.SUPPS_OrdersPlacedValue_CostEx, 
                      dbo.tSupplierStatsMonthly.SUPPS_OrdersPlacedValue_RetailInc, dbo.tSupplierStatsMonthly.SUPPS_DELLsReceivedQty, 
                      dbo.tSupplierStatsMonthly.SUPPS_DELLsReceivedValue_CostEx, dbo.tSupplierStatsMonthly.SUPPS_DELLSReceivedValue_RetailInc, 
                      dbo.tSupplierStatsMonthly.SUPPS_MissingLastStockTake_RetailInc, dbo.tSupplierStatsMonthly.SUPPS_MissingLastStockTake_CostEx, 
                      dbo.tSupplierStatsMonthly.SUPPS_MissingLastStockTake_qty, dbo.tSupplierStatsMonthly.SUPPS_StockValue_RRPInc')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vPerformance_StockBySupplier View Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.vPerformance_StockBySupplier View'
END
GO

--
-- Script To Update dbo.vPerformance_StockValue View In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.vPerformance_StockValue View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER VIEW dbo.vPerformance_StockValue
AS
SELECT     SUM(CAST(CAST(dbo.tProduct.P_QtyOnHand * dbo.tProduct.P_SP AS REAL) / dbo.tCurrency.CURR_Divisor AS REAL)) AS ValueAtSP, 
                      SUM(CAST(CAST(dbo.tProduct.P_QtyOnHand * dbo.tProduct.P_Cost AS REAL) / dbo.tCurrency.CURR_Divisor AS REAL)) AS ValueAtCost, 
                      CAST(CAST(dbo.tProduct.P_QtyOnHand * dbo.tProduct.P_RRP AS REAL) / dbo.tCurrency.CURR_Divisor AS REAL) AS ValueAtRRP, 
                      SUM(dbo.tProduct.P_QtyOnHand) AS QTYOH, ISNULL(dbo.vPublisherDistributorPerPID.TP_ID, 0) AS SUPPID, 
                      ISNULL(dbo.vMainSections.PSEC_SEC_ID, 0) AS CATEGORYID
FROM         dbo.vMainSections RIGHT OUTER JOIN
                      dbo.tProduct ON dbo.vMainSections.PSEC_P_ID = dbo.tProduct.P_ID LEFT OUTER JOIN
                      dbo.vPublisherDistributorPerPID ON dbo.tProduct.P_ID = dbo.vPublisherDistributorPerPID.P_ID CROSS JOIN
                      dbo.tCurrency INNER JOIN
                      dbo.tConfiguration ON dbo.tCurrency.CURR_ID = dbo.tConfiguration.CF_DefaultCurrID
WHERE     (dbo.tProduct.P_ProductType IN (''G'', ''B''))
GROUP BY ISNULL(dbo.vPublisherDistributorPerPID.TP_ID, 0), ISNULL(dbo.vMainSections.PSEC_SEC_ID, 0), 
                      CAST(CAST(dbo.tProduct.P_QtyOnHand * dbo.tProduct.P_RRP AS REAL) / dbo.tCurrency.CURR_Divisor AS REAL)')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vPerformance_StockValue View Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.vPerformance_StockValue View'
END
GO

--
-- Script To Create dbo.vSummaryPerformance View In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.vSummaryPerformance View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.vSummaryPerformance
AS
SELECT     TOP (100) PERCENT SUMM_ID, SUMM_Month, SUMM_StockValue_SPInc, SUMM_StockValue_CostEx, SUMM_StockValue_RRPInc, 
                      SUMM_SalesValue_RetailInc, SUMM_Margin, SUMM_Qty, SUMM_QtyInTop50, SUMM_SalesAsPercentOfTotalSales_RetailInc, 
                      SUMM_SalesAsPercentOfTotalSOH_RetailInc, SUMM_StockAsPercentOfTotalSOH_CostEx, SUMM_StockAsPercentOfTotalSOH_RRPInc, 
                      SUMM_ReturnsValue_RetailInc, SUMM_ReturnsValue_CostEx, SUMM_ReturnsAsPercentDeliveries, SUMM_ReturnsAsPercentSales, 
                      SUMM_OrdersPlacedQty, SUMM_OrdersPlacedValue_RetailInc, SUMM_OrdersPlacedValue_CostEx, SUMM_OrdersOSValue_RetailInc, 
                      SUMM_OrdersOSValue_CostEx, SUMM_MissingLastStockTake_RetailInc, SUMM_MissingLastStockTake_CostEx, SUMM_MissingLastStockTake_qty, 
                      SUMM_DELLsReceivedQty, SUMM_DELLsReceivedValue_CostEx, SUMM_DELLSReceivedValue_RetailInc, SUMM_Last12MonthSalesValue, 
                      SUMM_Last12MonthStockValue, CASE WHEN ISNULL(SUMM_StockValue_SPInc, 0) = 0 OR
                      ISNULL(SUMM_QtyMonthsInStockTurnRange, 0) = 0 OR
                      SUMM_Last12MonthStockValue = 0 THEN 0 ELSE SUMM_Last12MonthSalesValue * (12 / SUMM_QtyMonthsInStockTurnRange) 
                      / (SUMM_Last12MonthStockValue / SUMM_QtyMonthsInStockTurnRange) END AS StockTurn, SUMM_QtyMonthsInStockTurnRange
FROM         dbo.tSummaryStatsMonthly')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vSummaryPerformance View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.vSummaryPerformance View'
END
GO

--
-- Script To Update dbo.vSummaryPerformance_Pivot View In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.vSummaryPerformance_Pivot View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER VIEW dbo.vSummaryPerformance_Pivot
AS
SELECT     TOP (100) PERCENT h12, h11, h10, h09, h08, h07, h06, h05, h04, h03, h02, h01, m12_1, m11_1, m10_1, m09_1, m08_1, m07_1, m06_1, m05_1, 
                      m04_1, m03_1, m02_1, m01_1, m12_2, m11_2, m10_2, m09_2, m08_2, m07_2, m06_2, m05_2, m04_2, m03_2, m02_2, m01_2, m12_3, m11_3, 
                      m10_3, m09_3, m08_3, m07_3, m06_3, m05_3, m04_3, m03_3, m02_3, m01_3, m12_4, m11_4, m10_4, m09_4, m08_4, m07_4, m06_4, m05_4, 
                      m04_4, m03_4, m02_4, m01_4, m12_5, m11_5, m10_5, m09_5, m08_5, m07_5, m06_5, m05_5, m04_5, m03_5, m02_5, m01_5, m12_6, m11_6, 
                      m10_6, m09_6, m08_6, m07_6, m06_6, m05_6, m04_6, m03_6, m02_6, m01_6, m12_7, m11_7, m10_7, m09_7, m08_7, m07_7, m06_7, m05_7, 
                      m04_7, m03_7, m02_7, m01_7, m12_8, m11_8, m10_8, m09_8, m08_8, m07_8, m06_8, m05_8, m04_8, m03_8, m02_8, m01_8, m12_9, m11_9, 
                      m10_9, m09_9, m08_9, m07_9, m06_9, m05_9, m04_9, m03_9, m02_9, m01_9, m12_10, m11_10, m10_10, m09_10, m08_10, m07_10, m06_10, 
                      m05_10, m04_10, m03_10, m02_10, m01_10, m12_11, m11_11, m10_11, m09_11, m08_11, m07_11, m06_11, m05_11, m04_11, m03_11, m02_11, 
                      m01_11, m12_12, m11_12, m10_12, m09_12, m08_12, m07_12, m06_12, m05_12, m04_12, m03_12, m02_12, m01_12, m12_13, m11_13, m10_13, 
                      m09_13, m08_13, m07_13, m06_13, m05_13, m04_13, m03_13, m02_13, m01_13, m12_14, m11_14, m10_14, m09_14, m08_14, m07_14, m06_14, 
                      m05_14, m04_14, m03_14, m02_14, m01_14
FROM         (SELECT     TOP (100) PERCENT DATEADD(m, - 0, dbo.StartOfMonth(GETDATE())) AS h12, DATEADD(m, - 1, dbo.StartOfMonth(GETDATE())) AS h11, 
                                              DATEADD(m, - 2, dbo.StartOfMonth(GETDATE())) AS h10, DATEADD(m, - 3, dbo.StartOfMonth(GETDATE())) AS h09, DATEADD(m, - 4, 
                                              dbo.StartOfMonth(GETDATE())) AS h08, DATEADD(m, - 5, dbo.StartOfMonth(GETDATE())) AS h07, DATEADD(m, - 6, 
                                              dbo.StartOfMonth(GETDATE())) AS h06, DATEADD(m, - 7, dbo.StartOfMonth(GETDATE())) AS h05, DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GETDATE())) AS h04, DATEADD(m, - 9, dbo.StartOfMonth(GETDATE())) AS h03, DATEADD(m, - 10, 
                                              dbo.StartOfMonth(GETDATE())) AS h02, DATEADD(m, - 11, dbo.StartOfMonth(GETDATE())) AS h01, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_QtyInTop50, 0) ELSE 0 END) 
                                              AS m12_1, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_QtyInTop50, 0) 
                                              ELSE 0 END) AS m11_1, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_QtyInTop50, 0) ELSE 0 END) AS m10_1, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 3, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_QtyInTop50, 0) ELSE 0 END) AS m09_1, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_QtyInTop50, 0) ELSE 0 END) 
                                              AS m08_1, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_QtyInTop50, 0) 
                                              ELSE 0 END) AS m07_1, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_QtyInTop50, 0) ELSE 0 END) AS m06_1, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 7, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_QtyInTop50, 0) ELSE 0 END) AS m05_1, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_QtyInTop50, 0) ELSE 0 END) 
                                              AS m04_1, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_QtyInTop50, 0) 
                                              ELSE 0 END) AS m03_1, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_QtyInTop50, 0) ELSE 0 END) AS m02_1, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_QtyInTop50, 0) ELSE 0 END) AS m01_1, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_Salesvalue_RetailInc, 0) 
                                              ELSE 0 END) AS m12_2, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_Salesvalue_RetailInc, 0) ELSE 0 END) AS m11_2, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_Salesvalue_RetailInc, 0) ELSE 0 END) AS m10_2, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_Salesvalue_RetailInc, 0) 
                                              ELSE 0 END) AS m09_2, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_Salesvalue_RetailInc, 0) ELSE 0 END) AS m08_2, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_Salesvalue_RetailInc, 0) ELSE 0 END) AS m07_2, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_Salesvalue_RetailInc, 0) 
                                              ELSE 0 END) AS m06_2, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_Salesvalue_RetailInc, 0) ELSE 0 END) AS m05_2, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_Salesvalue_RetailInc, 0) ELSE 0 END) AS m04_2, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_Salesvalue_RetailInc, 0) 
                                              ELSE 0 END) AS m03_2, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_Salesvalue_RetailInc, 0) ELSE 0 END) AS m02_2, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_Salesvalue_RetailInc, 0) ELSE 0 END) AS m01_2, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m12_3, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m11_3, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m10_3, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m09_3, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m08_3, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m07_3, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m06_3, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m05_3, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m04_3, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m03_3, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m02_3, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 11, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m01_3, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m12_4, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, 
                                              - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m11_4, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m10_4, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, 
                                              - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m09_4, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m08_4, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, 
                                              - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m07_4, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m06_4, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, 
                                              - 7, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m05_4, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m04_4, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, 
                                              - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m03_4, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m02_4, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, 
                                              - 11, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m01_4, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_StockValue_CostEx, 0) 
                                              ELSE 0 END) AS m12_5, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_StockValue_CostEx, 0) ELSE 0 END) AS m11_5, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_StockValue_CostEx, 0) ELSE 0 END) AS m10_5, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_StockValue_CostEx, 0) 
                                              ELSE 0 END) AS m09_5, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_StockValue_CostEx, 0) ELSE 0 END) AS m08_5, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_StockValue_CostEx, 0) ELSE 0 END) AS m07_5, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_StockValue_CostEx, 0) 
                                              ELSE 0 END) AS m06_5, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_StockValue_CostEx, 0) ELSE 0 END) AS m05_5, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_StockValue_CostEx, 0) ELSE 0 END) AS m04_5, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_StockValue_CostEx, 0) 
                                              ELSE 0 END) AS m03_5, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_StockValue_CostEx, 0) ELSE 0 END) AS m02_5, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_StockValue_CostEx, 0) ELSE 0 END) AS m01_5, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_StockAsPercentOfTotalSOH_CostEx, 0) ELSE 0 END) AS m12_6, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, 
                                              - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_StockAsPercentOfTotalSOH_CostEx, 0) ELSE 0 END) AS m11_6, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_StockAsPercentOfTotalSOH_CostEx, 0) ELSE 0 END) AS m10_6, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, 
                                              - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_StockAsPercentOfTotalSOH_CostEx, 0) ELSE 0 END) AS m09_6, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_StockAsPercentOfTotalSOH_CostEx, 0) ELSE 0 END) AS m08_6, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, 
                                              - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_StockAsPercentOfTotalSOH_CostEx, 0) ELSE 0 END) AS m07_6, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_StockAsPercentOfTotalSOH_CostEx, 0) ELSE 0 END) AS m06_6, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, 
                                              - 7, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_StockAsPercentOfTotalSOH_CostEx, 0) ELSE 0 END) AS m05_6, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_StockAsPercentOfTotalSOH_CostEx, 0) ELSE 0 END) AS m04_6, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, 
                                              - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_StockAsPercentOfTotalSOH_CostEx, 0) ELSE 0 END) AS m03_6, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUMM_StockAsPercentOfTotalSOH_CostEx, 0) ELSE 0 END) AS m02_6, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, 
                                              - 11, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUMM_StockAsPercentOfTotalSOH_CostEx, 0) ELSE 0 END) AS m01_6, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m12_7, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m11_7, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m10_7, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m09_7, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m08_7, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m07_7, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m06_7, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m05_7, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m04_7, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m03_7, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m02_7, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 11, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m01_7, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsValue_RetailInc, 0) 
                                              ELSE 0 END) AS m12_8, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m11_8, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m10_8, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsValue_RetailInc, 0) 
                                              ELSE 0 END) AS m09_8, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m08_8, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m07_8, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsValue_RetailInc, 0) 
                                              ELSE 0 END) AS m06_8, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m05_8, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m04_8, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsValue_RetailInc, 0) 
                                              ELSE 0 END) AS m03_8, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m02_8, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m01_8, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsAsPercentDeliveries, 0) 
                                              ELSE 0 END) AS m12_9, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m11_9, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m10_9, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsAsPercentDeliveries, 0) 
                                              ELSE 0 END) AS m09_9, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m08_9, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m07_9, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsAsPercentDeliveries, 0) 
                                              ELSE 0 END) AS m06_9, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m05_9, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m04_9, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsAsPercentDeliveries, 0) 
                                              ELSE 0 END) AS m03_9, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m02_9, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m01_9, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsAsPercentSales, 0) 
                                              ELSE 0 END) AS m12_10, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_ReturnsAsPercentSales, 0) ELSE 0 END) AS m11_10, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsAsPercentSales, 0) ELSE 0 END) AS m10_10, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsAsPercentSales, 0) 
                                              ELSE 0 END) AS m09_10, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_ReturnsAsPercentSales, 0) ELSE 0 END) AS m08_10, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsAsPercentSales, 0) ELSE 0 END) AS m07_10, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsAsPercentSales, 0) 
                                              ELSE 0 END) AS m06_10, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_ReturnsAsPercentSales, 0) ELSE 0 END) AS m05_10, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsAsPercentSales, 0) ELSE 0 END) AS m04_10, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsAsPercentSales, 0) 
                                              ELSE 0 END) AS m03_10, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_ReturnsAsPercentSales, 0) ELSE 0 END) AS m02_10, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_ReturnsAsPercentSales, 0) ELSE 0 END) AS m01_10, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_Margin, 0) ELSE 0 END) 
                                              AS m12_11, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_Margin, 0) 
                                              ELSE 0 END) AS m11_11, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_Margin, 
                                              0) ELSE 0 END) AS m10_11, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_Margin, 0) ELSE 0 END) AS m09_11, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 4, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_Margin, 0) ELSE 0 END) AS m08_11, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, 
                                              - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_Margin, 0) ELSE 0 END) AS m07_11, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_Margin, 0) ELSE 0 END) 
                                              AS m06_11, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_Margin, 0) 
                                              ELSE 0 END) AS m05_11, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_Margin, 
                                              0) ELSE 0 END) AS m04_11, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_Margin, 0) ELSE 0 END) AS m03_11, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 10, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_Margin, 0) ELSE 0 END) AS m02_11, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, 
                                              - 11, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_Margin, 0) ELSE 0 END) AS m01_11, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_OrdersPlacedValue_CostEx, 0) 
                                              ELSE 0 END) AS m12_12, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_OrdersPlacedValue_CostEx, 0) ELSE 0 END) AS m11_12, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_OrdersPlacedValue_CostEx, 0) ELSE 0 END) AS m10_12, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_OrdersPlacedValue_CostEx, 0) 
                                              ELSE 0 END) AS m09_12, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_OrdersPlacedValue_CostEx, 0) ELSE 0 END) AS m08_12, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_OrdersPlacedValue_CostEx, 0) ELSE 0 END) AS m07_12, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_OrdersPlacedValue_CostEx, 0) 
                                              ELSE 0 END) AS m06_12, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_OrdersPlacedValue_CostEx, 0) ELSE 0 END) AS m05_12, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_OrdersPlacedValue_CostEx, 0) ELSE 0 END) AS m04_12, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_OrdersPlacedValue_CostEx, 0) 
                                              ELSE 0 END) AS m03_12, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_OrdersPlacedValue_CostEx, 0) ELSE 0 END) AS m02_12, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_OrdersPlacedValue_CostEx, 0) ELSE 0 END) AS m01_12, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_OrdersOSValue_CostEx, 0) 
                                              ELSE 0 END) AS m12_13, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m11_13, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m10_13, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_OrdersOSValue_CostEx, 0) 
                                              ELSE 0 END) AS m09_13, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m08_13, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m07_13, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_OrdersOSValue_CostEx, 0) 
                                              ELSE 0 END) AS m06_13, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m05_13, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m04_13, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_OrdersOSValue_CostEx, 0) 
                                              ELSE 0 END) AS m03_13, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m02_13, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m01_13, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_MissingLastStockTake_RetailInc, 
                                              0) ELSE 0 END) AS m12_14, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m11_14, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m10_14, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_MissingLastStockTake_RetailInc, 
                                              0) ELSE 0 END) AS m09_14, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m08_14, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m07_14, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_MissingLastStockTake_RetailInc, 
                                              0) ELSE 0 END) AS m06_14, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m05_14, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m04_14, 
                                              SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_MissingLastStockTake_RetailInc, 
                                              0) ELSE 0 END) AS m03_14, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUMM_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m02_14, SUM(CASE X.SUMM_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUMM_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m01_14
                       FROM          dbo.vSummaryPerformance AS X) AS Z')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vSummaryPerformance_Pivot View Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.vSummaryPerformance_Pivot View'
END
GO

--
-- Script To Update dbo.vSupplierPerformance View In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.vSupplierPerformance View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER VIEW dbo.vSupplierPerformance
AS
SELECT     TOP (100) PERCENT dbo.tSupplierStatsMonthly.SUPPS_Month, dbo.tSupplierStatsMonthly.SUPPS_SUPPLIERID AS SUPPLID, 
                      dbo.tSupplierStatsMonthly.SUPPS_StockValue_SPInc, dbo.tSupplierStatsMonthly.SUPPS_StockValue_CostEx, 
                      dbo.tSupplierStatsMonthly.SUPPS_StockValue_RRPInc, dbo.tSupplierStatsMonthly.SUPPS_SalesValue_RetailInc, 
                      dbo.tSupplierStatsMonthly.SUPPS_Margin, dbo.tSupplierStatsMonthly.SUPPS_Qty, dbo.tSupplierStatsMonthly.SUPPS_QtyInTop50, 
                      dbo.tSupplierStatsMonthly.SUPPS_SalesAsPercentOfTotalSales_RetailInc, dbo.tSupplierStatsMonthly.SUPPS_SalesAsPercentOfTotalSOH_RetailInc, 
                      dbo.tSupplierStatsMonthly.SUPPS_StockAsPercentOfTotalSOH_CostEx, dbo.tSupplierStatsMonthly.SUPPS_StockAsPercentOfTotalSOH_RRPInc, 
                      dbo.tSupplierStatsMonthly.SUPPS_ReturnsValue_RetailInc, dbo.tSupplierStatsMonthly.SUPPS_ReturnsValue_CostEx, 
                      dbo.tSupplierStatsMonthly.SUPPS_ReturnsAsPercentDeliveries, dbo.tSupplierStatsMonthly.SUPPS_ReturnsAsPercentSales, 
                      dbo.tSupplierStatsMonthly.SUPPS_OrdersPlacedQty, dbo.tSupplierStatsMonthly.SUPPS_OrdersPlacedValue_RetailInc, 
                      dbo.tSupplierStatsMonthly.SUPPS_OrdersPlacedValue_CostEx, dbo.tSupplierStatsMonthly.SUPPS_OrdersOSValue_RetailInc, 
                      dbo.tSupplierStatsMonthly.SUPPS_OrdersOSValue_CostEx, dbo.tSupplierStatsMonthly.SUPPS_MissingLastStockTake_RetailInc, 
                      dbo.tSupplierStatsMonthly.SUPPS_MissingLastStockTake_CostEx, dbo.tSupplierStatsMonthly.SUPPS_MissingLastStockTake_qty, 
                      dbo.tSupplierStatsMonthly.SUPPS_DELLsReceivedQty, dbo.tSupplierStatsMonthly.SUPPS_DELLsReceivedValue_CostEx, 
                      dbo.tSupplierStatsMonthly.SUPPS_DELLSReceivedValue_RetailInc, dbo.tTP.TP_Name AS SupplierName, 
                      dbo.tSupplierStatsMonthly.SUPPS_Last12MonthSalesValue, dbo.tSupplierStatsMonthly.SUPPS_Last12MonthStockValue, 
                      CASE WHEN ISNULL(SUPPS_StockValue_SPInc, 0) = 0 OR
                      ISNULL(SUPPS_QtyMonthsInStockTurnRange, 0) = 0 OR
                      SUPPS_Last12MonthStockValue = 0 THEN 0 ELSE SUPPS_Last12MonthSalesValue * (12 / SUPPS_QtyMonthsInStockTurnRange) 
                      / (SUPPS_Last12MonthStockValue / SUPPS_QtyMonthsInStockTurnRange) END AS StockTurn, 
                      dbo.tSupplierStatsMonthly.SUPPS_QtyMonthsInStockTurnRange
FROM         dbo.tSupplierStatsMonthly LEFT OUTER JOIN
                      dbo.tTP ON dbo.tSupplierStatsMonthly.SUPPS_SUPPLIERID = dbo.tTP.TP_ID')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vSupplierPerformance View Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.vSupplierPerformance View'
END
GO

--
-- Script To Update dbo.vSupplierPerformance_Pivot View In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.vSupplierPerformance_Pivot View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER VIEW dbo.vSupplierPerformance_Pivot
AS
SELECT     TOP (100) PERCENT SupplierName, SUPPLID, h12, h11, h10, h09, h08, h07, h06, h05, h04, h03, h02, h01, m12_1, m11_1, m10_1, m09_1, m08_1, 
                      m07_1, m06_1, m05_1, m04_1, m03_1, m02_1, m01_1, m12_2, m11_2, m10_2, m09_2, m08_2, m07_2, m06_2, m05_2, m04_2, m03_2, m02_2, 
                      m01_2, m12_3, m11_3, m10_3, m09_3, m08_3, m07_3, m06_3, m05_3, m04_3, m03_3, m02_3, m01_3, m12_4, m11_4, m10_4, m09_4, m08_4, 
                      m07_4, m06_4, m05_4, m04_4, m03_4, m02_4, m01_4, m12_5, m11_5, m10_5, m09_5, m08_5, m07_5, m06_5, m05_5, m04_5, m03_5, m02_5, 
                      m01_5, m12_6, m11_6, m10_6, m09_6, m08_6, m07_6, m06_6, m05_6, m04_6, m03_6, m02_6, m01_6, m12_7, m11_7, m10_7, m09_7, m08_7, 
                      m07_7, m06_7, m05_7, m04_7, m03_7, m02_7, m01_7, m12_8, m11_8, m10_8, m09_8, m08_8, m07_8, m06_8, m05_8, m04_8, m03_8, m02_8, 
                      m01_8, m12_9, m11_9, m10_9, m09_9, m08_9, m07_9, m06_9, m05_9, m04_9, m03_9, m02_9, m01_9, m12_10, m11_10, m10_10, m09_10, m08_10, 
                      m07_10, m06_10, m05_10, m04_10, m03_10, m02_10, m01_10, m12_11, m11_11, m10_11, m09_11, m08_11, m07_11, m06_11, m05_11, m04_11, 
                      m03_11, m02_11, m01_11, m12_12, m11_12, m10_12, m09_12, m08_12, m07_12, m06_12, m05_12, m04_12, m03_12, m02_12, m01_12, m12_13, 
                      m11_13, m10_13, m09_13, m08_13, m07_13, m06_13, m05_13, m04_13, m03_13, m02_13, m01_13, m12_14, m11_14, m10_14, m09_14, m08_14, 
                      m07_14, m06_14, m05_14, m04_14, m03_14, m02_14, m01_14
FROM         (SELECT     TOP (100) PERCENT CAST(SupplierName AS VARCHAR(100)) AS SupplierName, CAST(SUPPLID AS INT) AS SUPPLID, DATEADD(m, - 0, 
                                              dbo.StartOfMonth(GETDATE())) AS h12, DATEADD(m, - 1, dbo.StartOfMonth(GETDATE())) AS h11, DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GETDATE())) AS h10, DATEADD(m, - 3, dbo.StartOfMonth(GETDATE())) AS h09, DATEADD(m, - 4, 
                                              dbo.StartOfMonth(GETDATE())) AS h08, DATEADD(m, - 5, dbo.StartOfMonth(GETDATE())) AS h07, DATEADD(m, - 6, 
                                              dbo.StartOfMonth(GETDATE())) AS h06, DATEADD(m, - 7, dbo.StartOfMonth(GETDATE())) AS h05, DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GETDATE())) AS h04, DATEADD(m, - 9, dbo.StartOfMonth(GETDATE())) AS h03, DATEADD(m, - 10, 
                                              dbo.StartOfMonth(GETDATE())) AS h02, DATEADD(m, - 11, dbo.StartOfMonth(GETDATE())) AS h01, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_QtyInTop50, 0) ELSE 0 END) 
                                              AS m12_1, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_QtyInTop50, 0) 
                                              ELSE 0 END) AS m11_1, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_QtyInTop50, 0) ELSE 0 END) AS m10_1, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 3, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_QtyInTop50, 0) ELSE 0 END) AS m09_1, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_QtyInTop50, 0) ELSE 0 END) 
                                              AS m08_1, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_QtyInTop50, 0) 
                                              ELSE 0 END) AS m07_1, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_QtyInTop50, 0) ELSE 0 END) AS m06_1, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 7, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_QtyInTop50, 0) ELSE 0 END) AS m05_1, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_QtyInTop50, 0) ELSE 0 END) 
                                              AS m04_1, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_QtyInTop50, 0) 
                                              ELSE 0 END) AS m03_1, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_QtyInTop50, 0) ELSE 0 END) AS m02_1, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_QtyInTop50, 0) ELSE 0 END) AS m01_1, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_Salesvalue_RetailInc, 0) 
                                              ELSE 0 END) AS m12_2, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_Salesvalue_RetailInc, 0) ELSE 0 END) AS m11_2, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_Salesvalue_RetailInc, 0) ELSE 0 END) AS m10_2, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_Salesvalue_RetailInc, 0) 
                                              ELSE 0 END) AS m09_2, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_Salesvalue_RetailInc, 0) ELSE 0 END) AS m08_2, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_Salesvalue_RetailInc, 0) ELSE 0 END) AS m07_2, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_Salesvalue_RetailInc, 0) 
                                              ELSE 0 END) AS m06_2, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_Salesvalue_RetailInc, 0) ELSE 0 END) AS m05_2, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_Salesvalue_RetailInc, 0) ELSE 0 END) AS m04_2, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_Salesvalue_RetailInc, 0) 
                                              ELSE 0 END) AS m03_2, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_Salesvalue_RetailInc, 0) ELSE 0 END) AS m02_2, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_Salesvalue_RetailInc, 0) ELSE 0 END) AS m01_2, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m12_3, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m11_3, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m10_3, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m09_3, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m08_3, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m07_3, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m06_3, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m05_3, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m04_3, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m03_3, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m02_3, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 11, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSales_RetailInc, 0) ELSE 0 END) AS m01_3, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m12_4, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m11_4, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m10_4, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m09_4, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m08_4, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m07_4, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m06_4, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m05_4, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m04_4, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m03_4, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m02_4, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 11, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_SalesAsPercentOfTotalSOH_RetailInc, 0) ELSE 0 END) AS m01_4, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_StockValue_RRPINC, 0) 
                                              ELSE 0 END) AS m12_5, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_StockValue_CostEx, 0) ELSE 0 END) AS m11_5, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_StockValue_CostEx, 0) ELSE 0 END) AS m10_5, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_StockValue_CostEx, 0) 
                                              ELSE 0 END) AS m09_5, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_StockValue_CostEx, 0) ELSE 0 END) AS m08_5, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_StockValue_CostEx, 0) ELSE 0 END) AS m07_5, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_StockValue_CostEx, 0) 
                                              ELSE 0 END) AS m06_5, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_StockValue_CostEx, 0) ELSE 0 END) AS m05_5, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_StockValue_CostEx, 0) ELSE 0 END) AS m04_5, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_StockValue_CostEx, 0) 
                                              ELSE 0 END) AS m03_5, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_StockValue_CostEx, 0) ELSE 0 END) AS m02_5, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_StockValue_CostEx, 0) ELSE 0 END) AS m01_5, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m12_6, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, 
                                              - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m11_6, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m10_6, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, 
                                              - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m09_6, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m08_6, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, 
                                              - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m07_6, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m06_6, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, 
                                              - 7, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m05_6, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m04_6, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, 
                                              - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m03_6, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(X.SUPPS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m02_6, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, 
                                              - 11, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.SUPPS_StockAsPercentOfTotalSOH_RRPInc, 0) ELSE 0 END) AS m01_6, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m12_7, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m11_7, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m10_7, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m09_7, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m08_7, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m07_7, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m06_7, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m05_7, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m04_7, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m03_7, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m02_7, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 11, dbo.StartOfMonth(GetDate())) THEN ISNULL(StockTurn, 0) ELSE 0 END) AS m01_7, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsValue_RetailInc, 0) 
                                              ELSE 0 END) AS m12_8, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m11_8, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m10_8, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsValue_RetailInc, 0) 
                                              ELSE 0 END) AS m09_8, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m08_8, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m07_8, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsValue_RetailInc, 0) 
                                              ELSE 0 END) AS m06_8, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m05_8, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m04_8, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsValue_RetailInc, 0) 
                                              ELSE 0 END) AS m03_8, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m02_8, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsValue_RetailInc, 0) ELSE 0 END) AS m01_8, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsAsPercentDeliveries, 0) 
                                              ELSE 0 END) AS m12_9, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m11_9, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m10_9, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsAsPercentDeliveries, 0) 
                                              ELSE 0 END) AS m09_9, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m08_9, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m07_9, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsAsPercentDeliveries, 0) 
                                              ELSE 0 END) AS m06_9, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m05_9, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m04_9, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsAsPercentDeliveries, 0) 
                                              ELSE 0 END) AS m03_9, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m02_9, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsAsPercentDeliveries, 0) ELSE 0 END) AS m01_9, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsAsPercentSales, 0) 
                                              ELSE 0 END) AS m12_10, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_ReturnsAsPercentSales, 0) ELSE 0 END) AS m11_10, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsAsPercentSales, 0) ELSE 0 END) AS m10_10, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsAsPercentSales, 0) 
                                              ELSE 0 END) AS m09_10, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_ReturnsAsPercentSales, 0) ELSE 0 END) AS m08_10, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsAsPercentSales, 0) ELSE 0 END) AS m07_10, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsAsPercentSales, 0) 
                                              ELSE 0 END) AS m06_10, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_ReturnsAsPercentSales, 0) ELSE 0 END) AS m05_10, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsAsPercentSales, 0) ELSE 0 END) AS m04_10, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsAsPercentSales, 0) 
                                              ELSE 0 END) AS m03_10, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_ReturnsAsPercentSales, 0) ELSE 0 END) AS m02_10, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_ReturnsAsPercentSales, 0) ELSE 0 END) AS m01_10, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_Margin, 0) ELSE 0 END) 
                                              AS m12_11, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_Margin, 0) 
                                              ELSE 0 END) AS m11_11, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_Margin, 
                                              0) ELSE 0 END) AS m10_11, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_Margin, 0) ELSE 0 END) AS m09_11, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 4, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_Margin, 0) ELSE 0 END) AS m08_11, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, 
                                              - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_Margin, 0) ELSE 0 END) AS m07_11, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_Margin, 0) ELSE 0 END) 
                                              AS m06_11, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_Margin, 0) 
                                              ELSE 0 END) AS m05_11, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 8, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_Margin, 
                                              0) ELSE 0 END) AS m04_11, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_Margin, 0) ELSE 0 END) AS m03_11, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 10, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_Margin, 0) ELSE 0 END) AS m02_11, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, 
                                              - 11, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_Margin, 0) ELSE 0 END) AS m01_11, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_OrdersPlacedValue_RetailInc, 0) 
                                              ELSE 0 END) AS m12_12, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_OrdersPlacedValue_RetailInc, 0) ELSE 0 END) AS m11_12, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_OrdersPlacedValue_RetailInc, 0) ELSE 0 END) AS m10_12, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_OrdersPlacedValue_RetailInc, 0) 
                                              ELSE 0 END) AS m09_12, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_OrdersPlacedValue_RetailInc, 0) ELSE 0 END) AS m08_12, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_OrdersPlacedValue_RetailInc, 0) ELSE 0 END) AS m07_12, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_OrdersPlacedValue_RetailInc, 0) 
                                              ELSE 0 END) AS m06_12, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_OrdersPlacedValue_RetailInc, 0) ELSE 0 END) AS m05_12, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_OrdersPlacedValue_RetailInc, 0) ELSE 0 END) AS m04_12, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_OrdersPlacedValue_RetailInc, 0) 
                                              ELSE 0 END) AS m03_12, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_OrdersPlacedValue_RetailInc, 0) ELSE 0 END) AS m02_12, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_OrdersPlacedValue_RetailInc, 0) ELSE 0 END) AS m01_12, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_OrdersOSValue_CostEx, 0) 
                                              ELSE 0 END) AS m12_13, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m11_13, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m10_13, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_OrdersOSValue_CostEx, 0) 
                                              ELSE 0 END) AS m09_13, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m08_13, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m07_13, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_OrdersOSValue_CostEx, 0) 
                                              ELSE 0 END) AS m06_13, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m05_13, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m04_13, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_OrdersOSValue_CostEx, 0) 
                                              ELSE 0 END) AS m03_13, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m02_13, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 11, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_OrdersOSValue_CostEx, 0) ELSE 0 END) AS m01_13, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_MissingLastStockTake_RetailInc, 
                                              0) ELSE 0 END) AS m12_14, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m11_14, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 2, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m10_14, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_MissingLastStockTake_RetailInc, 
                                              0) ELSE 0 END) AS m09_14, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m08_14, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 5, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m07_14, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_MissingLastStockTake_RetailInc, 
                                              0) ELSE 0 END) AS m06_14, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 7, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m05_14, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 8, 
                                              dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m04_14, 
                                              SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 9, dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_MissingLastStockTake_RetailInc, 
                                              0) ELSE 0 END) AS m03_14, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 10, dbo.StartOfMonth(GetDate())) 
                                              THEN ISNULL(SUPPS_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m02_14, SUM(CASE X.SUPPS_MONTH WHEN DATEADD(m, - 11,
                                               dbo.StartOfMonth(GetDate())) THEN ISNULL(SUPPS_MissingLastStockTake_RetailInc, 0) ELSE 0 END) AS m01_14
                       FROM          dbo.vSupplierPerformance AS X
                       GROUP BY SupplierName, SUPPLID
                       ORDER BY SupplierName) AS Z
ORDER BY SupplierName')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vSupplierPerformance_Pivot View Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.vSupplierPerformance_Pivot View'
END
GO

--
-- Script To Update dbo.CreateStatsSet Procedure In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.CreateStatsSet Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER PROCEDURE dbo.CreateStatsSet
AS
BEGIN
	SET NOCOUNT ON;
DECLARE @CHECKRecord INT
--DECLARE @DATEOFFIRSTTRANSACTIONINSYSTEM DATETIME
DECLARE @MTH DATETIME
DECLARE @SUPPID INT
DECLARE @TOTALMONTHSALES_RetailInc REAL
DECLARE @TOTALSOH_RetailInc REAL
DECLARE @TOTALSOH_CostEx REAL
DECLARE @TOTALSDEL_CostEx REAL
DECLARE @CNT INT
DECLARE @MthKey DATETIME
DECLARE @DATE DATETIME
--THis procedure is always run after the start of the new month (usually between midnight and start of business on the first day)
--so we should get the the previous month


BEGIN TRY

		SELECT @DATE = dbo.StartOfMonth(CF_UPDATEWINDOWEND) FROM tConfiguration
		DELETE FROM tSupplierStatsMonthly WHERE SUPPS_MONTH < @DATE
		DELETE FROM tCategoryStatsMonthly WHERE CATS_MONTH < @DATE
		DELETE FROM tSummaryStatsMonthly WHERE SUMM_MONTH < @DATE

		SELECT @DATE =GETDATE()
--			SELECT @DATEOFFIRSTTRANSACTIONINSYSTEM = MIN(TR_PROCESSINGDATE) FROM tTR
		SELECT @MTH = dbo.GetMonthYear(@Date)
		TRUNCATE TABLE tTop50
		EXEC dbo.GetPerformance_TopFifty  @MTH

--First we create or update the tSupplierStatsMonthly table
------------------------------------------------------------
			--make sure we don''t duplicate the record
			SELECT @CHECKRecord = COUNT(SUPPS_SUPPLIERID) FROM tSupplierStatsMonthly WHERE SUPPS_MONTH = dbo.StartOfMonth(@DATE) 
--			DELETE FROM tSupplierStatsMonthly WHERE SUPPS_MONTH = dbo.StartOfMonth(@DATE)
			--Insert record with stock value info
			IF @CheckRecord = 0
			BEGIN
				INSERT INTO tSupplierStatsMonthly (SUPPS_MONTH,SUPPS_SUPPLIERID,SUPPS_StockValue_CostEx,SUPPS_StockValue_SPInc,SUPPS_StockValue_RRPInc,SUPPS_StockQty) 
				SELECT dbo.StartOfMonth(@DATE),
						ISNULL(SUPPID,0),
						SUM(ISNULL(ValueAtCost,0)),
						SUM(ISNULL(ValueAtSP,0)),
						SUM(ISNULL(ValueAtRRP,0)),
						SUM(ISNULL(QTYOH,0)) 
					FROM vPerformance_StockValue GROUP BY SUPPID
			END
			ELSE
			BEGIN
				UPDATE tSupplierStatsMonthly 
				SET SUPPS_StockValue_CostEx = SVCE,
				SUPPS_StockValue_SPInc = SVSP,
				SUPPS_StockValue_RRPInc = RetP,
				SUPPS_StockQty = Q
					FROM tSupplierStatsMonthly a JOIN 
				(SELECT ISNULL(SUPPID,0) as SUPPID,
						SUM(ISNULL(ValueAtCost,0)) as SVCE,
						SUM(ISNULL(ValueAtSP,0)) as SVSP,
						SUM(ISNULL(ValueAtRRP,0)) as RetP,
						SUM(ISNULL(QTYOH,0)) as Q FROM vPerformance_StockValue GROUP BY SUPPID) b
				ON a.SUPPS_SUPPLIERID = b.SUPPID AND a.SUPPS_MONTH = dbo.StartOfMonth(@DATE)
				WHERE SUPPS_MONTH = dbo.StartOfMonth(@DATE)
			END
			--Update the inserted record with sales information  NOTE:TotalRetail is the price charged after discount i.e the sales value
			UPDATE tSupplierStatsMonthly SET SUPPS_SalesValue_RetailInc = b.TR,SUPPS_SalesValue_CostEx = b.TC,SUPPS_SalesQty = b.Q 
				FROM tSupplierStatsMonthly a JOIN 
				(SELECT SUPPID,SUM(ISNULL(TOTALRETAIL,0)) TR ,SUM(ISNULL(TOTALCOSTExVAT,0)) TC,SUM(ISNULL(QTY,0)) Q 
						FROM vSalesSummary 
						WHERE dte BETWEEN dbo.StartOfMonth(@DATE) and dbo.EndOfMonth(@DATE) GROUP BY SUPPID) b 
				ON a.SUPPS_SUPPLIERID = b.SUPPID AND a.SUPPS_MONTH = dbo.StartOfMonth(@DATE)
				WHERE SUPPS_MONTH = dbo.StartOfMonth(@DATE)

			--Update the inserted record with return values
			UPDATE tSupplierStatsMonthly SET 
				SUPPS_ReturnsValue_RetailInc = b.VRI,
				SUPPS_ReturnsValue_CostEx = b.vCE,
				SUPPS_ReturnsQty = b.QT 
				FROM tSupplierStatsMonthly a JOIN 
				(SELECT SUPPID,SUM(ISNULL(Value_RetailInc,0)) VRI ,SUM(ISNULL(Value_CostEx,0)) vCE,SUM(ISNULL(QtyReturned,0)) QT
						FROM vReturnedStock 
						WHERE dte BETWEEN dbo.StartOfMonth(@DATE) and dbo.EndOfMonth(@DATE) GROUP BY SUPPID) b 
				ON a.SUPPS_SUPPLIERID = b.SUPPID AND a.SUPPS_MONTH = dbo.StartOfMonth(@DATE)
				WHERE SUPPS_MONTH = dbo.StartOfMonth(@DATE)
		--update orders placed in period
			UPDATE tSupplierStatsMonthly SET 
				SUPPS_OrdersPlacedQty = b.Qty,
				SUPPS_OrdersPlacedValue_CostEx = b.vCE,
				SUPPS_OrdersPlacedValue_RetailInc = b.RI 
				FROM tSupplierStatsMonthly a JOIN 
				(SELECT SUPPID,
						SUM(ISNULL(Qty,0)) Qty ,
						SUM(ISNULL(CostExVAT,0)) vCE,
						SUM(ISNULL(Val,0)) RI
						FROM vPOLsPlacedInPeriod
						WHERE dte BETWEEN dbo.StartOfMonth(@DATE) and dbo.EndOfMonth(@DATE) GROUP BY SUPPID) b 
				ON a.SUPPS_SUPPLIERID = b.SUPPID AND a.SUPPS_MONTH = dbo.StartOfMonth(@DATE)
				WHERE SUPPS_MONTH = dbo.StartOfMonth(@DATE)
		--update deliveries received in period
			UPDATE tSupplierStatsMonthly SET 
				SUPPS_DELLsReceivedQty = b.Qty,
				SUPPS_DELLsReceivedValue_CostEx = b.vCE,
				SUPPS_DELLSReceivedValue_RetailInc = b.RI 
				FROM tSupplierStatsMonthly a JOIN 
				(SELECT SUPPID,
						SUM(ISNULL(Qty,0)) Qty ,
						SUM(ISNULL(value_CostEx,0)) vCE,
						SUM(ISNULL(Value_RetailInc,0)) RI
						FROM vDELLSinPeriod
						WHERE dte BETWEEN dbo.StartOfMonth(@DATE) and dbo.EndOfMonth(@DATE) GROUP BY SUPPID) b 
				ON a.SUPPS_SUPPLIERID = b.SUPPID AND a.SUPPS_MONTH = dbo.StartOfMonth(@DATE)
				WHERE SUPPS_MONTH = dbo.StartOfMonth(@DATE)

--Update Orders outstanding for current month   - THIS IS DIFFERENT IN THAT IT IS A SNAPSHOT OF ORDERS OS AT THIS POINT
			UPDATE tSupplierStatsMonthly SET 
				SUPPS_OrdersOSQty = b.Qty,
				SUPPS_OrdersOSValue_CostEx = b.vCE,
				SUPPS_OrdersOSValue_RetailInc = b.RI 
				FROM tSupplierStatsMonthly a JOIN 
				(SELECT SUPPID,
						SUM(ISNULL(QtyofItemsOS,0)) Qty ,
						SUM(ISNULL(ValueOfOrdersOS_CostExVat,0)) vCE,
						SUM(ISNULL(ValueOfOrdersOS_RetailInc,0)) RI
						FROM vPOLsOSPerSUPPID_Summary
						GROUP BY SUPPID) b 
				ON a.SUPPS_SUPPLIERID = b.SUPPID 


			UPDATE tSupplierStatsMonthly 
				SET SUPPS_Last12MonthStockValue = c.TwelveMonthStockValue_RRPINC,
					SUPPS_Last12MonthSalesValue = b.TR,
					SUPPS_QtyMonthsInStockTurnRange =
						DATEDIFF(mm, dbo.MaxDate(DATEADD(mm,-12,dbo.StartOfMonth(@DATE)),dbo.StartOfMonth(CF_UPDATEWINDOWEND)),dbo.StartOfMonth(@DATE)) + 1
					FROM tSupplierStatsMonthly a  CROSS JOIN tConfiguration	
				LEFT JOIN 
					(SELECT SUPPID,SUM(ISNULL(TOTALRETAIL,0)) TR ,SUM(ISNULL(QTY,0)) Q 
						FROM vSalesSummary LEFT JOIN tTP  ON SUPPID = TP_ID  CROSS JOIN tConfiguration
						WHERE dte BETWEEN  dbo.MaxDate(DATEADD(mm,-12,dbo.StartOfMonth(@DATE)),dbo.StartOfMonth(CF_UPDATEWINDOWEND)) and dbo.EndOfMonth(@DATE)
						GROUP BY SUPPID) b ON a.SUPPS_SUPPLIERID = b.SUPPID AND a.SUPPS_MONTH = dbo.StartOfMonth(@DATE) 
				LEFT JOIN
					(SELECT SUPPS_SUPPLIERID,SUM(SUPPS_StockValue_RRPInc) as TwelveMonthStockValue_RRPINC 
						FROM tSupplierStatsMonthly  LEFT JOIN tTP  ON SUPPS_SupplierID = TP_ID CROSS JOIN tConfiguration
						WHERE SUPPS_MONTH  BETWEEN dbo.MaxDate(DATEADD(mm,-12,dbo.StartOfMonth(@DATE)),dbo.StartOfMonth(CF_UPDATEWINDOWEND)) and dbo.EndOfMonth(@DATE)
						GROUP BY SUPPS_SupplierID) c ON a.SUPPS_SUPPLIERID = c.SUPPS_SUPPLIERID AND a.SUPPS_MONTH = dbo.StartOfMonth(@DATE)
				LEFT JOIN 
					tTP d 
					ON a.SUPPS_SupplierID = d.TP_ID 
					WHERE SUPPS_MONTH = dbo.StartOfMonth(@DATE)

		SELECT @TOTALMONTHSALES_RetailInc =  ISNULL(RetailInc,0) FROM vPerformance_4 WHERE mth = @MTH

		SELECT	@TOTALSOH_RetailInc =  SUM(ISNULL(SUPPS_StockValue_RRPInc,0)),
				@TOTALSOH_CostEx = SUM(ISNULL(SUPPS_StockValue_CostEx,0)),
				@TOTALSDEL_CostEx = SUM(ISNULL(SUPPS_DELLsReceivedValue_CostEx,0))
		FROM vPerformance_StockBySupplier WHERE SUPPS_Month = @MTH


		--Get number of titles in top 50 per supplier

		UPDATE tSupplierStatsMonthly SET SUPPS_QtyInTop50 = b.cnt
			FROM tSupplierStatsMonthly a JOIN (SELECT COUNT(a.PID) cnt, a.SUPPID SUPPID
						FROM vPerformance_2 a JOIN tTop50 b ON a.PID = b.PID WHERE a.mth = b.Mth AND a.SuppID = b.TPID GROUP BY SUPPID) b
			ON a.SUPPS_SUPPLIERID = b.SUPPID 
			WHERE a.SUPPS_Month = @MTH

		--Update percent of sales
		UPDATE tSupplierStatsMonthly SET SUPPS_SalesAsPercentOfTotalSales_RetailInc = 
				CASE WHEN ISNULL(@TOTALMONTHSALES_RetailInc,0) > 0 THEN (ISNULL(SUPPS_SalesValue_RetailInc,0)*100/@TOTALMONTHSALES_RetailInc) ELSE 0 END ,

				SUPPS_SalesAsPercentOfTotalSOH_RetailInc = 
				CASE WHEN ISNULL(@TOTALSOH_RetailInc,0) > 0 THEN (ISNULL(SUPPS_SalesValue_RetailInc,0)*100/@TOTALSOH_RetailInc) ELSE 999 END,

				SUPPS_StockAsPercentOfTotalSOH_CostEx = 
				CASE WHEN ISNULL(@TOTALSOH_CostEx,0) > 0 then (ISNULL(SUPPS_StockValue_CostEx,0)*100/@TOTALSOH_CostEx) ELSE 0 END,

				SUPPS_StockAsPercentOfTotalSOH_RRPInc = 
				CASE WHEN ISNULL(@TOTALSOH_RetailInc,0) > 0 then (ISNULL(SUPPS_StockValue_RRPInc,0)*100/@TOTALSOH_RetailInc) ELSE 0 END,

				SUPPS_ReturnsAsPercentDeliveries = 
				CASE WHEN ISNULL(SUPPS_DELLsReceivedValue_CostEx,0) > 0 THEN (ISNULL(SUPPS_ReturnsValue_CostEx,0)*100/SUPPS_DELLsReceivedValue_CostEx) ELSE 0 END,

				SUPPS_ReturnsAsPercentSales = 
				CASE WHEN ISNULL(SUPPS_SalesValue_RetailInc,0) > 0 THEN (ISNULL(SUPPS_ReturnsValue_RetailInc,0)*100/SUPPS_SalesValue_RetailInc) ELSE 0 END,

				SUPPS_Margin = ((ISNULL(SUPPS_Salesvalue_RetailInc,0)/1.14)- (SUPPS_SalesValue_CostEx))
			WHERE SUPPS_MONTH = @MTH


--Then we create or update the tCategoryStatsMonthly table
------------------------------------------------------------
			--make sure we don''t duplicate the record
			SELECT @CHECKRecord = COUNT(CATS_CATEGORYID) FROM tCategoryStatsMonthly WHERE CATS_MONTH = dbo.StartOfMonth(@DATE) 
			--Insert record with stock value info
			IF @CheckRecord = 0
			BEGIN
				INSERT INTO tCategoryStatsMonthly (CATS_MONTH,CATS_CATEGORYID,CATS_StockValue_CostEx,CATS_StockValue_SPInc,CATS_StockValue_RRPInc,CATS_StockQty) 
				SELECT dbo.StartOfMonth(@DATE),
						ISNULL(CATEGORYID,0),
						SUM(ISNULL(ValueAtCost,0)),
						SUM(ISNULL(ValueAtSP,0)),
						SUM(ISNULL(ValueAtRRP,0)),
						SUM(ISNULL(QTYOH,0)) 
					FROM vPerformance_StockValue GROUP BY CATEGORYID
			END
			ELSE
			BEGIN
				UPDATE tCategoryStatsMonthly 
				SET CATS_StockValue_CostEx = SVCE,
				CATS_StockValue_SPInc = SVSP,
				CATS_StockValue_RRPInc = RetP,
				CATS_StockQty = Q
					FROM tCategoryStatsMonthly a JOIN 
				(SELECT ISNULL(CATEGORYID,0) as CATEGORYID,
						SUM(ISNULL(ValueAtCost,0)) as SVCE,
						SUM(ISNULL(ValueAtSP,0)) as SVSP,
						SUM(ISNULL(ValueAtRRP,0)) as RetP,
						SUM(ISNULL(QTYOH,0)) as Q FROM vPerformance_StockValue GROUP BY CATEGORYID) b
				ON a.CATS_CATEGORYID = b.CATEGORYID AND a.CATS_MONTH = dbo.StartOfMonth(@DATE)
				WHERE CATS_MONTH = dbo.StartOfMonth(@DATE)
			END
			--Update the inserted record with sales information  NOTE:TotalRetail is the price charged after discount i.e the sales value
			UPDATE tCategoryStatsMonthly SET CATS_SalesValue_RetailInc = b.TR,CATS_SalesValue_CostEx = b.TC,CATS_SalesQty = b.Q 
				FROM tCategoryStatsMonthly a JOIN 
				(SELECT MainSectionID,SUM(ISNULL(TOTALRETAIL,0)) TR ,SUM(ISNULL(TOTALCOSTExVAT,0)) TC,SUM(ISNULL(QTY,0)) Q 
						FROM vSalesSummary 
						WHERE dte BETWEEN dbo.StartOfMonth(@DATE) and dbo.EndOfMonth(@DATE) GROUP BY MainSectionID) b 
				ON a.CATS_CATEGORYID = b.MainSectionID AND a.CATS_MONTH = dbo.StartOfMonth(@DATE)
				WHERE CATS_MONTH = dbo.StartOfMonth(@DATE)
--
			--Update the inserted record with return values
			UPDATE tCategoryStatsMonthly SET 
				CATS_ReturnsValue_RetailInc = b.VRI,
				CATS_ReturnsValue_CostEx = b.vCE,
				CATS_ReturnsQty = b.QT 
				FROM tCategoryStatsMonthly a JOIN 
				(SELECT CATEGORYID,SUM(ISNULL(Value_RetailInc,0)) VRI ,SUM(ISNULL(Value_CostEx,0)) vCE,SUM(ISNULL(QtyReturned,0)) QT
						FROM vReturnedStock 
						WHERE dte BETWEEN dbo.StartOfMonth(@DATE) and dbo.EndOfMonth(@DATE) GROUP BY CATEGORYID) b 
				ON a.CATS_CATEGORYID = b.CATEGORYID AND a.CATS_MONTH = dbo.StartOfMonth(@DATE)
				WHERE CATS_MONTH = dbo.StartOfMonth(@DATE)
		--update orders placed in period
			UPDATE tCategoryStatsMonthly SET 
				CATS_OrdersPlacedQty = b.Qty,
				CATS_OrdersPlacedValue_CostEx = b.vCE,
				CATS_OrdersPlacedValue_RetailInc = b.RI 
				FROM tCategoryStatsMonthly a JOIN 
				(SELECT CATEGORYID,
						SUM(ISNULL(Qty,0)) Qty ,
						SUM(ISNULL(CostExVAT,0)) vCE,
						SUM(ISNULL(Val,0)) RI
						FROM vPOLsPlacedInPeriod
						WHERE dte BETWEEN dbo.StartOfMonth(@DATE) and dbo.EndOfMonth(@DATE) GROUP BY CATEGORYID) b 
				ON a.CATS_CATEGORYID = b.CATEGORYID AND a.CATS_MONTH = dbo.StartOfMonth(@DATE)
				WHERE CATS_MONTH = dbo.StartOfMonth(@DATE)
		--update deliveries received in period
			UPDATE tCategoryStatsMonthly SET 
				CATS_DELLsReceivedQty = b.Qty,
				CATS_DELLsReceivedValue_CostEx = b.vCE,
				CATS_DELLSReceivedValue_RetailInc = b.RI 
				FROM tCategoryStatsMonthly a JOIN 
				(SELECT CATEGORYID,
						SUM(ISNULL(Qty,0)) Qty ,
						SUM(ISNULL(value_CostEx,0)) vCE,
						SUM(ISNULL(Value_RetailInc,0)) RI
						FROM vDELLSinPeriod
						WHERE dte BETWEEN dbo.StartOfMonth(@DATE) and dbo.EndOfMonth(@DATE) GROUP BY CATEGORYID) b 
				ON a.CATS_CATEGORYID = b.CATEGORYID AND a.CATS_MONTH = dbo.StartOfMonth(@DATE)
				WHERE CATS_MONTH = dbo.StartOfMonth(@DATE)
--
--Update Orders outstanding for current month   - THIS IS DIFFERENT IN THAT IT IS A SNAPSHOT OF ORDERS OS AT THIS POINT
			UPDATE tCategoryStatsMonthly SET 
				CATS_OrdersOSQty = b.Qty,
				CATS_OrdersOSValue_CostEx = b.vCE,
				CATS_OrdersOSValue_RetailInc = b.RI 
				FROM tCategoryStatsMonthly a JOIN 
				(SELECT CATEGORYID,
						SUM(ISNULL(QtyofItemsOS,0)) Qty ,
						SUM(ISNULL(ValueOfOrdersOS_CostExVat,0)) vCE,
						SUM(ISNULL(ValueOfOrdersOS_RetailInc,0)) RI
						FROM vPOLsOSPerSUPPID_Summary
						GROUP BY CATEGORYID) b 
				ON a.CATS_CATEGORYID = b.CATEGORYID 
--
			--Update rolling 12 month average stock values (for calculating stock turn)
			UPDATE tCategoryStatsMonthly 
				SET CATS_Last12MonthStockValue = c.TwelveMonthStockValue_RRPINC ,
					CATS_Last12MonthSalesValue = b.TR,
					CATS_QtyMonthsInStockTurnRange =
						DATEDIFF(mm, dbo.MaxDate(DATEADD(mm,-12,dbo.StartOfMonth(@DATE)),dbo.StartOfMonth(CF_UPDATEWINDOWEND)),dbo.StartOfMonth(@DATE)) + 1
						FROM tCategoryStatsMonthly a   CROSS JOIN tConfiguration	

				LEFT JOIN -- to get sales value
					(SELECT MainSectionID,SUM(ISNULL(TOTALRETAIL,0)) TR ,SUM(ISNULL(QTY,0)) Q 
						FROM vSalesSummary  LEFT JOIN tTP  ON SUPPID = TP_ID  CROSS JOIN tConfiguration	
						WHERE dte BETWEEN dbo.MaxDate(DATEADD(mm,-12,dbo.StartOfMonth(@DATE)),dbo.StartOfMonth(CF_UPDATEWINDOWEND)) and dbo.EndOfMonth(@DATE) + 1
						GROUP BY MainSectionID) b 
						ON a.CATS_CATEGORYID = b.MainSectionID AND CATS_MONTH = dbo.StartOfMonth(@DATE) 

				LEFT JOIN   --to get stock value
					(SELECT CATS_CATEGORYID,SUM(CATS_StockValue_RRPInc) as TwelveMonthStockValue_RRPINC 
						FROM tCategoryStatsMonthly    CROSS JOIN tConfiguration	
						WHERE CATS_MONTH  BETWEEN dbo.MaxDate(DATEADD(mm,-12,dbo.StartOfMonth(@DATE)),dbo.StartOfMonth(CF_UPDATEWINDOWEND))	and dbo.EndOfMonth(@DATE) + 1
						GROUP BY CATS_CATEGORYID) c
						ON a.CATS_CATEGORYID = c.CATS_CATEGORYID AND CATS_MONTH = dbo.StartOfMonth(@DATE)

					WHERE CATS_MONTH = dbo.StartOfMonth(@DATE)

		SELECT @TOTALMONTHSALES_RetailInc =  ISNULL(RetailInc,0) FROM vPerformance_4 WHERE mth = @MTH

		SELECT	@TOTALSOH_RetailInc =  SUM(ISNULL(CATS_StockValue_RRPInc,0)),
				@TOTALSOH_CostEx = SUM(ISNULL(CATS_StockValue_CostEx,0)),
				@TOTALSDEL_CostEx = SUM(ISNULL(CATS_DELLsReceivedValue_CostEx,0))
		FROM vPerformance_StockByCategory WHERE CATS_Month = @MTH

		--Get number of titles in top 50 per supplier

		UPDATE tCategoryStatsMonthly SET CATS_QtyInTop50 = b.cnt
			FROM tCategoryStatsMonthly a JOIN (SELECT COUNT(a.PID) cnt, a.MainSectionID CATEGORYID
						FROM vPerformance_2 a JOIN tTop50 b ON a.PID = b.PID WHERE a.mth = b.Mth AND a.MainSectionID = b.MaincategoryID GROUP BY MainSectionID) b
			ON a.CATS_CATEGORYID = b.CATEGORYID 
			WHERE a.CATS_Month = @MTH


				
		--Update percent of sales
		UPDATE tCategoryStatsMonthly SET CATS_SalesAsPercentOfTotalSales_RetailInc = 
				CASE WHEN ISNULL(@TOTALMONTHSALES_RetailInc,0) > 0 THEN (ISNULL(CATS_SalesValue_RetailInc,0)*100/@TOTALMONTHSALES_RetailInc) ELSE 0 END ,

				CATS_SalesAsPercentOfTotalSOH_RetailInc = 
				CASE WHEN ISNULL(@TOTALSOH_RetailInc,0) > 0 THEN (ISNULL(CATS_SalesValue_RetailInc,0)*100/@TOTALSOH_RetailInc) ELSE 0 END,

				CATS_StockAsPercentOfTotalSOH_CostEx = 
				CASE WHEN ISNULL(@TOTALSOH_CostEx,0) > 0 then (ISNULL(CATS_StockValue_CostEx,0)*100/@TOTALSOH_CostEx) ELSE 0 END,

				CATS_StockAsPercentOfTotalSOH_RRPInc = 
				CASE WHEN ISNULL(@TOTALSOH_RetailInc,0) > 0 then (ISNULL(CATS_StockValue_RRPInc,0)*100/@TOTALSOH_RetailInc) ELSE 0 END,

				CATS_ReturnsAsPercentDeliveries = 
				CASE WHEN ISNULL(CATS_DELLsReceivedValue_CostEx,0) > 0 THEN (ISNULL(CATS_ReturnsValue_CostEx,0)*100/CATS_DELLsReceivedValue_CostEx) ELSE 0 END,

				CATS_ReturnsAsPercentSales = 
				CASE WHEN ISNULL(CATS_SalesValue_RetailInc,0) > 0 THEN (ISNULL(CATS_ReturnsValue_RetailInc,0)*100/CATS_SalesValue_RetailInc) ELSE 0 END,

				CATS_Margin = ((ISNULL(CATS_Salesvalue_RetailInc,0)/1.14)- (CATS_SalesValue_CostEx))
			WHERE CATS_MONTH = @MTH


---------------Summary Data-------------------------------------------------------------------------------------
--Then we create or update the tSummaryStatsMonthly table
------------------------------------------------------------
			--make sure we don''t duplicate the record
			SELECT @CHECKRecord = COUNT(SUMM_ID) FROM tSummaryStatsMonthly WHERE SUMM_MONTH = dbo.StartOfMonth(@DATE) 
			--Insert record with stock value info
			IF @CheckRecord = 0
			BEGIN
				INSERT INTO tSummaryStatsMonthly (SUMM_MONTH,SUMM_StockValue_CostEx,SUMM_StockValue_SPInc,SUMM_StockValue_RRPInc,SUMM_StockQty) 
				SELECT dbo.StartOfMonth(@DATE),
						SUM(ISNULL(ValueAtCost,0)),
						SUM(ISNULL(ValueAtSP,0)),
						SUM(ISNULL(ValueAtRRP,0)),
						SUM(ISNULL(QTYOH,0)) 
					FROM vPerformance_StockValue 
			END
			ELSE
			BEGIN
				UPDATE tSummaryStatsMonthly 
				SET SUMM_StockValue_CostEx = SVCE,
				SUMM_StockValue_SPInc = SVSP,
				SUMM_StockValue_RRPInc = RetP,
				SUMM_StockQty = Q
					FROM tSummaryStatsMonthly a JOIN 
				(SELECT SUM(ISNULL(ValueAtCost,0)) as SVCE,
						SUM(ISNULL(ValueAtSP,0)) as SVSP,
						SUM(ISNULL(ValueAtRRP,0)) as RetP,
						SUM(ISNULL(QTYOH,0)) as Q FROM vPerformance_StockValue) b
				ON a.SUMM_MONTH = dbo.StartOfMonth(@DATE)
				WHERE SUMM_MONTH = dbo.StartOfMonth(@DATE)
			END
			--Update the inserted record with sales information  NOTE:TotalRetail is the price charged after discount i.e the sales value
			UPDATE tSummaryStatsMonthly SET SUMM_SalesValue_RetailInc = b.TR,SUMM_SalesValue_CostEx = b.TC,SUMM_SalesQty = b.Q 
				FROM tSummaryStatsMonthly a JOIN 
				(SELECT SUM(ISNULL(TOTALRETAIL,0)) TR ,SUM(ISNULL(TOTALCOSTExVAT,0)) TC,SUM(ISNULL(QTY,0)) Q 
						FROM vSalesSummary 
						WHERE dte BETWEEN dbo.StartOfMonth(@DATE) and dbo.EndOfMonth(@DATE)) b 
				ON  a.SUMM_MONTH = dbo.StartOfMonth(@DATE)
				WHERE SUMM_MONTH = dbo.StartOfMonth(@DATE)
--
			--Update the inserted record with return values
			UPDATE tSummaryStatsMonthly SET 
				SUMM_ReturnsValue_RetailInc = b.VRI,
				SUMM_ReturnsValue_CostEx = b.vCE,
				SUMM_ReturnsQty = b.QT 
				FROM tSummaryStatsMonthly a JOIN 
				(SELECT SUM(ISNULL(Value_RetailInc,0)) VRI ,SUM(ISNULL(Value_CostEx,0)) vCE,SUM(ISNULL(QtyReturned,0)) QT
						FROM vReturnedStock 
						WHERE dte BETWEEN dbo.StartOfMonth(@DATE) and dbo.EndOfMonth(@DATE)) b 
				ON a.SUMM_MONTH = dbo.StartOfMonth(@DATE)
				WHERE SUMM_MONTH = dbo.StartOfMonth(@DATE)
		--update orders placed in period
			UPDATE tSummaryStatsMonthly SET 
				SUMM_OrdersPlacedQty = b.Qty,
				SUMM_OrdersPlacedValue_CostEx = b.vCE,
				SUMM_OrdersPlacedValue_RetailInc = b.RI 
				FROM tSummaryStatsMonthly a JOIN 
				(SELECT SUM(ISNULL(Qty,0)) Qty ,
						SUM(ISNULL(CostExVAT,0)) vCE,
						SUM(ISNULL(Val,0)) RI
						FROM vPOLsPlacedInPeriod
						WHERE dte BETWEEN dbo.StartOfMonth(@DATE) and dbo.EndOfMonth(@DATE)) b 
				ON a.SUMM_MONTH = dbo.StartOfMonth(@DATE)
				WHERE SUMM_MONTH = dbo.StartOfMonth(@DATE)
		--update deliveries received in period
			UPDATE tSummaryStatsMonthly SET 
				SUMM_DELLsReceivedQty = b.Qty,
				SUMM_DELLsReceivedValue_CostEx = b.vCE,
				SUMM_DELLSReceivedValue_RetailInc = b.RI 
				FROM tSummaryStatsMonthly a JOIN 
				(SELECT SUM(ISNULL(Qty,0)) Qty ,
						SUM(ISNULL(value_CostEx,0)) vCE,
						SUM(ISNULL(Value_RetailInc,0)) RI
						FROM vDELLSinPeriod
						WHERE dte BETWEEN dbo.StartOfMonth(@DATE) and dbo.EndOfMonth(@DATE)) b 
				ON a.SUMM_MONTH = dbo.StartOfMonth(@DATE)
				WHERE SUMM_MONTH = dbo.StartOfMonth(@DATE)
--
--Update Orders outstanding for current month   - THIS IS DIFFERENT IN THAT IT IS A SNAPSHOT OF ORDERS OS AT THIS POINT
			UPDATE tSummaryStatsMonthly SET 
				SUMM_OrdersOSQty = b.Qty,
				SUMM_OrdersOSValue_CostEx = b.vCE,
				SUMM_OrdersOSValue_RetailInc = b.RI 
				FROM tSummaryStatsMonthly a JOIN 
				(SELECT SUM(ISNULL(QtyofItemsOS,0)) Qty ,
						SUM(ISNULL(ValueOfOrdersOS_CostExVat,0)) vCE,
						SUM(ISNULL(ValueOfOrdersOS_RetailInc,0)) RI
						FROM vPOLsOSPerSUPPID_Summary) b ON SUMM_MONTH = @DATE
--
			--Update rolling 12 month average stock values (for calculating stock turn)
			UPDATE tSummaryStatsMonthly 
				SET SUMM_Last12MonthStockValue = c.TwelveMonthStockValue_RRPINC ,
					SUMM_Last12MonthSalesValue = b.TR,
					SUMM_QtyMonthsInStockTurnRange =
						DATEDIFF(mm, dbo.MaxDate(DATEADD(mm,-12,dbo.StartOfMonth(@DATE)),dbo.StartOfMonth(CF_UPDATEWINDOWEND)),dbo.StartOfMonth(@DATE)) + 1
						FROM tSummaryStatsMonthly a   CROSS JOIN tConfiguration	

				LEFT JOIN -- to get sales value
					(SELECT SUM(ISNULL(TOTALRETAIL,0)) TR ,SUM(ISNULL(QTY,0)) Q 
						FROM vSalesSummary  LEFT JOIN tTP  ON SUPPID = TP_ID  CROSS JOIN tConfiguration	
						WHERE dte BETWEEN dbo.MaxDate(DATEADD(mm,-12,dbo.StartOfMonth(@DATE)),dbo.StartOfMonth(CF_UPDATEWINDOWEND)) and dbo.EndOfMonth(@DATE) + 1
						) b 
						ON SUMM_MONTH = dbo.StartOfMonth(@DATE) 

				LEFT JOIN   --to get stock value
					(SELECT SUM(SUMM_StockValue_RRPInc) as TwelveMonthStockValue_RRPINC 
						FROM tSummaryStatsMonthly    CROSS JOIN tConfiguration	
						WHERE SUMM_MONTH  BETWEEN dbo.MaxDate(DATEADD(mm,-12,dbo.StartOfMonth(@DATE)),dbo.StartOfMonth(CF_UPDATEWINDOWEND))	and dbo.EndOfMonth(@DATE) + 1
						) c
						ON  SUMM_MONTH = dbo.StartOfMonth(@DATE)

					WHERE SUMM_MONTH = dbo.StartOfMonth(@DATE)

		SELECT @TOTALMONTHSALES_RetailInc =  ISNULL(RetailInc,0) FROM vPerformance_4 WHERE mth = @MTH

		SELECT	@TOTALSOH_RetailInc =  SUM(ISNULL(ValueAtRRP,0)),
				@TOTALSOH_CostEx = SUM(ISNULL(ValueAtCost,0))--,
				--@TOTALSDEL_CostEx = SUM(ISNULL(SUMM_DELLsReceivedValue_CostEx,0))
		FROM vPerformance_StockValue

		--Get number of titles in top 50 per supplier

		UPDATE tSummaryStatsMonthly SET SUMM_QtyInTop50 = b.cnt
			FROM tSummaryStatsMonthly a JOIN (SELECT COUNT(a.PID) cnt
						FROM vPerformance_2 a JOIN tTop50 b ON a.PID = b.PID WHERE a.mth = b.Mth) b
			ON a.SUMM_Month = @MTH

		--Update percent of sales
		UPDATE tSummaryStatsMonthly SET SUMM_SalesAsPercentOfTotalSales_RetailInc = 
				CASE WHEN ISNULL(@TOTALMONTHSALES_RetailInc,0) > 0 THEN (ISNULL(SUMM_SalesValue_RetailInc,0)*100/@TOTALMONTHSALES_RetailInc) ELSE 0 END ,

				SUMM_SalesAsPercentOfTotalSOH_RetailInc = 
				CASE WHEN ISNULL(@TOTALSOH_RetailInc,0) > 0 THEN (ISNULL(SUMM_SalesValue_RetailInc,0)*100/@TOTALSOH_RetailInc) ELSE 0 END,

				SUMM_StockAsPercentOfTotalSOH_CostEx = 
				CASE WHEN ISNULL(@TOTALSOH_CostEx,0) > 0 then (ISNULL(SUMM_StockValue_CostEx,0)*100/@TOTALSOH_CostEx) ELSE 0 END,

				SUMM_StockAsPercentOfTotalSOH_RRPInc = 
				CASE WHEN ISNULL(@TOTALSOH_RetailInc,0) > 0 then (ISNULL(SUMM_StockValue_RRPInc,0)*100/@TOTALSOH_RetailInc) ELSE 0 END,

				SUMM_ReturnsAsPercentDeliveries = 
				CASE WHEN ISNULL(SUMM_DELLsReceivedValue_CostEx,0) > 0 THEN (ISNULL(SUMM_ReturnsValue_CostEx,0)*100/SUMM_DELLsReceivedValue_CostEx) ELSE 0 END,

				SUMM_ReturnsAsPercentSales = 
				CASE WHEN ISNULL(SUMM_SalesValue_RetailInc,0) > 0 THEN (ISNULL(SUMM_ReturnsValue_RetailInc,0)*100/SUMM_SalesValue_RetailInc) ELSE 0 END,

				SUMM_Margin = ((ISNULL(SUMM_Salesvalue_RetailInc,0)/1.14)- (SUMM_SalesValue_CostEx))
			WHERE SUMM_MONTH = @MTH

END TRY
BEGIN CATCH

DECLARE @ErrorString NVARCHAR(MAX)

		IF ERROR_PROCEDURE() <> OBJECT_NAME(@@PROCID)
			SELECT @ErrorString = ERROR_MESSAGE();
		ELSE
			SELECT @ErrorString = ''Error Message: '' + ERROR_MESSAGE() + '', Error Procedure: '' + 
						RTRIM(CONVERT(CHAR,ERROR_PROCEDURE())) + '', Error Line: '' + RTRIM(CONVERT(Char,ERROR_LINE ()));
		RAISERROR (@ErrorString, 16,1)
END CATCH

END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.CreateStatsSet Procedure Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.CreateStatsSet Procedure'
END
GO

--
-- Script To Update dbo.SaveLog Procedure In 5.198.70.15\PBKSINSTANCE2.PBKS
-- Generated Monday, May 31, 2010, at 08:51 AM
--
-- Please backup 5.198.70.15\PBKSINSTANCE2.PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.SaveLog Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER PROCEDURE [dbo].[SaveLog] (@Msg VARCHAR(MAX) = '''',
@PROCEDURENAME VARCHAR(MAX) = ''NULL'',
@XMLDATA XML = NULL,
@STRINGDATA VARCHAR(MAX) = ''null'')
AS
BEGIN
DECLARE @ErrorString NVARCHAR(4000)

	BEGIN TRY
		SELECT @MSG = @MSG + ''/'' + @STRINGDATA
		INSERT INTO _tSBLog (SBL_Msg,SBL_Proc,SBL_XMLData) VALUES (@Msg,@PROCEDURENAME, @XMLDATA)
	END TRY
	BEGIN CATCH
		IF ERROR_PROCEDURE() <> OBJECT_NAME(@@PROCID)
			SELECT @ErrorString = ERROR_MESSAGE();
		ELSE
			SELECT @ErrorString = ''Error Message: '' + ERROR_MESSAGE() + '', Error Procedure: '' + 
						ERROR_PROCEDURE() + '', Error Line: '' + CAST(ERROR_LINE() as VARCHAR(10));
		
		RAISERROR (@ErrorString, 16,1)
	END CATCH
END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.SaveLog Procedure Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.SaveLog Procedure'
END
GO