--
-- Script To Update dbo.tBudget Table In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.tBudget Table'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tBudget_B_ID')
      ALTER TABLE [dbo].[tBudget] DROP CONSTRAINT [DF_tBudget_B_ID]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE TABLE [dbo].[tmp_tBudget] (
   [B_ID] [uniqueidentifier] ROWGUIDCOL NOT NULL CONSTRAINT [DF_tBudget_B_ID] DEFAULT (newid()),
   [B_BudgetMonth] [datetime] NOT NULL,
   [B_DeliveriesBudget] [numeric] (18, 2) NULL,
   [B_ReturnsBudget] [numeric] (18, 2) NULL,
   [B_RetailValueReturnsIssued] [numeric] (18, 2) NULL,
   [B_RetailValueReturnsInProcess] [numeric] (18, 2) NULL,
   [B_OrdersAtRetailValueIssued] [numeric] (18, 2) NULL,
   [B_OrdersAtRetailValueInProcess] [numeric] (18, 2) NULL,
   [B_OrdersAgainstBudget] [numeric] (18, 2) NULL,
   [B_RetailValueReceivedIssued] [numeric] (18, 2) NULL,
   [B_RetailValueReceivedInProcess] [numeric] (18, 2) NULL,
   [B_DeliveriesAgainstBudget] [numeric] (18, 2) NULL,
   [B_DeliveriesAgainstBudget_FourMonthAverage] [numeric] (18, 2) NULL
)
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   INSERT INTO [dbo].[tmp_tBudget] ([B_ID], [B_BudgetMonth], [B_DeliveriesBudget], [B_ReturnsBudget], [B_RetailValueReturnsIssued], [B_RetailValueReturnsInProcess], [B_OrdersAtRetailValueIssued], [B_OrdersAtRetailValueInProcess], [B_OrdersAgainstBudget], [B_RetailValueReceivedIssued], [B_RetailValueReceivedInProcess], [B_DeliveriesAgainstBudget], [B_DeliveriesAgainstBudget_FourMonthAverage])
   SELECT [B_ID], [B_BudgetMonth], NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL
   FROM [dbo].[tBudget]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   DROP TABLE [dbo].[tBudget]
GO

sp_rename N'[dbo].[tmp_tBudget]', N'tBudget'

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   ALTER TABLE [dbo].[tBudget] ADD CONSTRAINT [PK_tBudget] PRIMARY KEY CLUSTERED ([B_ID])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.tBudget Table Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.tBudget Table'
END
GO

--
-- Script To Update dbo.tDEL Table In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.tDEL Table'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE INDEX [SupplierInvoiceDate] ON [dbo].[tDEL] ([DEL_SupplierInvoiceDate]) WITH (ALLOW_PAGE_LOCKS = OFF)
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.tDEL Table Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.tDEL Table'
END
GO

--
-- Script To Update dbo.tExportToAccountingLog Table In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.tExportToAccountingLog Table'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE TABLE [dbo].[tmp_tExportToAccountingLog] (
   [RowID] [int] IDENTITY (1, 1) NOT NULL,
   [FKEY] [int] NULL,
   [Period] [int] NULL,
   [TransactionNominalDate] [datetime] NULL,
   [GDC] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [Acno] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [Reference] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [Description] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [Amount] [numeric] (15, 2) NULL,
   [TaxType] [int] NULL,
   [TaxAmount] [numeric] (15, 2) NULL,
   [Openitem] [char] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [Costcode] [char] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [ContraAccount] [char] (9) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [ExchangeRate] [numeric] (9, 6) NULL,
   [BankExchangeRate] [numeric] (9, 6) NULL,
   [BatchID] [int] NULL,
   [DiscountTax] [numeric] (12, 2) NULL,
   [DiscountAmount] [numeric] (12, 2) NULL,
   [HomeAmount] [numeric] (12, 2) NULL,
   [TRGLOBlobalID] [uniqueidentifier] NULL,
   [SignedDate] [datetime] NULL
)
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   SET IDENTITY_INSERT [dbo].[tmp_tExportToAccountingLog] ON
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   INSERT INTO [dbo].[tmp_tExportToAccountingLog] ([RowID], [FKEY], [Period], [TransactionNominalDate], [GDC], [Acno], [Reference], [Description], [Amount], [TaxType], [TaxAmount], [Openitem], [Costcode], [ContraAccount], [ExchangeRate], [BankExchangeRate], [BatchID], [DiscountTax], [DiscountAmount], [HomeAmount], [TRGLOBlobalID], [SignedDate])
   SELECT [RowID], NULL, [Period], [TransactionNominalDate], [GDC], [Acno], [Reference], [Description], [Amount], [TaxType], [TaxAmount], [Openitem], [Costcode], [ContraAccount], [ExchangeRate], [BankExchangeRate], [BatchID], [DiscountTax], [DiscountAmount], [HomeAmount], [TRGLOBlobalID], [SignedDate]
   FROM [dbo].[tExportToAccountingLog]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   SET IDENTITY_INSERT [dbo].[tmp_tExportToAccountingLog] OFF
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   DROP TABLE [dbo].[tExportToAccountingLog]
GO

sp_rename N'[dbo].[tmp_tExportToAccountingLog]', N'tExportToAccountingLog'

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   ALTER TABLE [dbo].[tExportToAccountingLog] ADD CONSTRAINT [PK_tExportToAccountingLog] PRIMARY KEY CLUSTERED ([RowID])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.tExportToAccountingLog Table Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.tExportToAccountingLog Table'
END
GO

--
-- Script To Update dbo.tExportToAccountingMaster Table In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.tExportToAccountingMaster Table'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   ALTER TABLE [dbo].[tExportToAccountingMaster] ADD CONSTRAINT [PK_tExportToAccountingMaster] PRIMARY KEY CLUSTERED ([RowID])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.tExportToAccountingMaster Table Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.tExportToAccountingMaster Table'
END
GO

--
-- Script To Update dbo.tProduct Table In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.tProduct Table'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'FK_tProduct_tPT')
      ALTER TABLE [dbo].[tProduct] DROP CONSTRAINT [FK_tProduct_tPT]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF__tProduct__P_ID__6A30C649')
      ALTER TABLE [dbo].[tProduct] DROP CONSTRAINT [DF__tProduct__P_ID__6A30C649]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF__tProduct__P_Date__6B24EA82')
      ALTER TABLE [dbo].[tProduct] DROP CONSTRAINT [DF__tProduct__P_Date__6B24EA82]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tProduct_P_ExcludeFromSales')
      ALTER TABLE [dbo].[tProduct] DROP CONSTRAINT [DF_tProduct_P_ExcludeFromSales]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE TABLE [dbo].[tmp_tProduct] (
   [P_ID] [uniqueidentifier] ROWGUIDCOL NOT NULL CONSTRAINT [DF__tProduct__P_ID__6A30C649] DEFAULT (newid()),
   [P_Image] [varbinary] (max) NULL,
   [P_Title] [varchar] (900) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_Article] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_Subtitle] [varchar] (3000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_EAN] [varchar] (13) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_ProductType_ID] [int] NULL,
   [P_RRP] [int] NULL,
   [P_SP] [numeric] (16, 2) NULL,
   [P_Cost] [real] NULL,
   [P_Special] [int] NULL,
   [P_USPrice] [int] NULL,
   [P_UKPrice] [int] NULL,
   [P_EUPrice] [int] NULL,
   [P_ForeignOrderedCURRID] [int] NULL,
   [P_ForeignOrderedPrice] [int] NULL,
   [P_Seriestitle] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_MainAuthor] [varchar] (900) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_Publisher] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_Edition] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_BindingCode] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_PubDate] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_PubPlace] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_FlagText] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_Comment] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_Note] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_Summary] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_Description] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_SupplierID] [int] NULL,
   [P_BFSupplierCode] [varchar] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_DealID] [int] NULL,
   [P_LastDateOrdered] [datetime] NULL,
   [P_LastPriceOrdered] [int] NULL,
   [P_LastQtyFirmOrdered] [int] NULL,
   [P_LastQtySSOrdered] [int] NULL,
   [P_LastQtyDelivered] [int] NULL,
   [P_LastPriceDelivered] [int] NULL,
   [P_LastDateDelivered] [datetime] NULL,
   [P_DateLastStockTake] [datetime] NULL,
   [P_LastDateSold] [datetime] NULL,
   [P_LastQtySold] [int] NULL,
   [P_LastPriceSold] [int] NULL,
   [P_QtyLastStockTake] [int] NULL,
   [P_LastStockTakeID] [int] NULL,
   [P_CatalogHeadingID] [int] NULL,
   [P_StckAgeQty6Mnths] [int] NULL,
   [P_StckAgeQty12Mnths] [int] NULL,
   [P_StckAgeQty18Mnths] [int] NULL,
   [P_StckAgeQty18MnthsPlus] [int] NULL,
   [P_StckAgeDate] [datetime] NULL,
   [P_Obsolete] [bit] NULL,
   [P_OOS] [bit] NULL,
   [P_DateRecordAdded] [datetime] NULL CONSTRAINT [DF__tProduct__P_Date__6B24EA82] DEFAULT (getdate()),
   [P_LastCopySerial] [smallint] NULL,
   [OLDPID] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_QtyCopiesOnHand] [int] NULL,
   [P_QtyOnHand] [int] NULL,
   [P_QtyReserved] [int] NULL,
   [P_QtyExpectedBack] [int] NULL,
   [P_QtyOnAppro] [int] NULL,
   [P_QtyOnOrder] [int] NULL,
   [P_QtyOnOrder_UnIssued] [int] NULL,
   [P_QtyOnBackorder] [int] NULL,
   [P_QtyTotalSold] [int] NULL,
   [P_SpecialVAT] [bit] NULL,
   [P_VATRate] [numeric] (8, 2) NULL,
   [P_BottomOfDocument] [bit] NULL,
   [P_Section] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_BIC] [varchar] (51) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_DateLastModified] [datetime] NULL,
   [P_CatHead_ID] [int] NULL,
   [P_Seesafe] [int] NULL,
   [P_Status] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_SkipBFWash] [bit] NULL,
   [P_DefaultDeliveryDays] [int] NULL,
   [P_LastApproToTPID] [int] NULL,
   [P_ReturnAvailability] [smallint] NULL,
   [P_ProductType] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_Code] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_LoyaltyRate] [smallint] NULL,
   [P_SalesCurrentMonth] [int] NULL,
   [P_SalesYTD] [int] NULL,
   [P_NDA] [bit] NULL,
   [P_ExcludeFromSales] [bit] NULL CONSTRAINT [DF_tProduct_P_ExcludeFromSales] DEFAULT ((0)),
   [P_QtyOnHand_PreST] [int] NULL,
   [P_CostLastStockTake] [real] NULL,
   [P_LastForeignPrice] [int] NULL,
   [P_LastFCID] [int] NULL,
   [P_LastFCFactor] [numeric] (9, 5) NULL,
   [P_ReRunQty] [real] NULL,
   [P_ReRunAvgCost] [real] NULL,
   [P_MasterCategory] [int] NULL,
   [P_Weight] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_Length] [real] NULL,
   [P_Width] [real] NULL,
   [P_SystemStatus] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_MultibuyCode] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [P_Core] [bit] NULL,
   [P_TinyImage] [varbinary] (max) NULL
)
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   INSERT INTO [dbo].[tmp_tProduct] ([P_ID], [P_Image], [P_Title], [P_Article], [P_Subtitle], [P_EAN], [P_ProductType_ID], [P_RRP], [P_SP], [P_Cost], [P_Special], [P_USPrice], [P_UKPrice], [P_EUPrice], [P_ForeignOrderedCURRID], [P_ForeignOrderedPrice], [P_Seriestitle], [P_MainAuthor], [P_Publisher], [P_Edition], [P_BindingCode], [P_PubDate], [P_PubPlace], [P_FlagText], [P_Comment], [P_Note], [P_Summary], [P_Description], [P_SupplierID], [P_BFSupplierCode], [P_DealID], [P_LastDateOrdered], [P_LastPriceOrdered], [P_LastQtyFirmOrdered], [P_LastQtySSOrdered], [P_LastQtyDelivered], [P_LastPriceDelivered], [P_LastDateDelivered], [P_DateLastStockTake], [P_LastDateSold], [P_LastQtySold], [P_LastPriceSold], [P_QtyLastStockTake], [P_LastStockTakeID], [P_CatalogHeadingID], [P_StckAgeQty6Mnths], [P_StckAgeQty12Mnths], [P_StckAgeQty18Mnths], [P_StckAgeQty18MnthsPlus], [P_StckAgeDate], [P_Obsolete], [P_OOS], [P_DateRecordAdded], [P_LastCopySerial], [OLDPID], [P_QtyCopiesOnHand], [P_QtyOnHand], [P_QtyReserved], [P_QtyExpectedBack], [P_QtyOnAppro], [P_QtyOnOrder], [P_QtyOnOrder_UnIssued], [P_QtyOnBackorder], [P_QtyTotalSold], [P_SpecialVAT], [P_VATRate], [P_BottomOfDocument], [P_Section], [P_BIC], [P_DateLastModified], [P_CatHead_ID], [P_Seesafe], [P_Status], [P_SkipBFWash], [P_DefaultDeliveryDays], [P_LastApproToTPID], [P_ReturnAvailability], [P_ProductType], [P_Code], [P_LoyaltyRate], [P_SalesCurrentMonth], [P_SalesYTD], [P_NDA], [P_ExcludeFromSales], [P_QtyOnHand_PreST], [P_CostLastStockTake], [P_LastForeignPrice], [P_LastFCID], [P_LastFCFactor], [P_ReRunQty], [P_ReRunAvgCost], [P_MasterCategory], [P_Weight], [P_Length], [P_Width], [P_SystemStatus], [P_MultibuyCode], [P_Core], [P_TinyImage])
   SELECT [P_ID], [P_Image], [P_Title], [P_Article], [P_Subtitle], [P_EAN], [P_ProductType_ID], [P_RRP], [P_SP], [P_Cost], [P_Special], [P_USPrice], [P_UKPrice], [P_EUPrice], [P_ForeignOrderedCURRID], [P_ForeignOrderedPrice], [P_Seriestitle], [P_MainAuthor], [P_Publisher], [P_Edition], [P_BindingCode], [P_PubDate], [P_PubPlace], [P_FlagText], [P_Comment], [P_Note], [P_Summary], [P_Description], [P_SupplierID], [P_BFSupplierCode], [P_DealID], [P_LastDateOrdered], [P_LastPriceOrdered], [P_LastQtyFirmOrdered], [P_LastQtySSOrdered], [P_LastQtyDelivered], [P_LastPriceDelivered], [P_LastDateDelivered], [P_DateLastStockTake], [P_LastDateSold], [P_LastQtySold], [P_LastPriceSold], [P_QtyLastStockTake], [P_LastStockTakeID], [P_CatalogHeadingID], [P_StckAgeQty6Mnths], [P_StckAgeQty12Mnths], [P_StckAgeQty18Mnths], [P_StckAgeQty18MnthsPlus], [P_StckAgeDate], [P_Obsolete], [P_OOS], [P_DateRecordAdded], [P_LastCopySerial], [OLDPID], [P_QtyCopiesOnHand], [P_QtyOnHand], [P_QtyReserved], [P_QtyExpectedBack], [P_QtyOnAppro], [P_QtyOnOrder], NULL, [P_QtyOnBackorder], [P_QtyTotalSold], [P_SpecialVAT], [P_VATRate], [P_BottomOfDocument], [P_Section], [P_BIC], [P_DateLastModified], [P_CatHead_ID], [P_Seesafe], [P_Status], [P_SkipBFWash], [P_DefaultDeliveryDays], [P_LastApproToTPID], [P_ReturnAvailability], [P_ProductType], [P_Code], [P_LoyaltyRate], [P_SalesCurrentMonth], [P_SalesYTD], [P_NDA], [P_ExcludeFromSales], [P_QtyOnHand_PreST], [P_CostLastStockTake], [P_LastForeignPrice], [P_LastFCID], [P_LastFCFactor], [P_ReRunQty], [P_ReRunAvgCost], [P_MasterCategory], [P_Weight], [P_Length], [P_Width], [P_SystemStatus], [P_MultibuyCode], [P_Core], [P_TinyImage]
   FROM [dbo].[tProduct]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   DROP TABLE [dbo].[tProduct]
GO

sp_rename N'[dbo].[tmp_tProduct]', N'tProduct'

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   ALTER TABLE [dbo].[tProduct] ADD CONSTRAINT [PK_tProduct] PRIMARY KEY NONCLUSTERED ([P_ID])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE INDEX [iAuthor] ON [dbo].[tProduct] ([P_MainAuthor])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE INDEX [iBIC] ON [dbo].[tProduct] ([P_BIC])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE INDEX [iCode] ON [dbo].[tProduct] ([P_Code])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE INDEX [idxSupplierID] ON [dbo].[tProduct] ([P_SupplierID])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE INDEX [iEAN] ON [dbo].[tProduct] ([P_EAN])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE INDEX [iProductType] ON [dbo].[tProduct] ([P_ProductType])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE INDEX [iProductTypeID] ON [dbo].[tProduct] ([P_ProductType_ID])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE INDEX [iSupplier] ON [dbo].[tProduct] ([P_SupplierID])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE CLUSTERED INDEX [iTitle] ON [dbo].[tProduct] ([P_Title])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE INDEX [old] ON [dbo].[tProduct] ([OLDPID])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1

exec('CREATE TRIGGER dbo.storeP ON dbo.tProduct 
AFTER INSERT
AS
DECLARE @DEFSTID INT
DECLARE @PID VARCHAR(50)
DECLARE @QtyOnHand INT
DECLARE @QtyReserved INT
DECLARE @QtyOnBackOrder INT
DECLARE @QtyOnOrder INT
DECLARE @QtyCopiesOnHand INT
DECLARE @QtyCopiesReserved INT
DECLARE @QtyOnAppro INT
DECLARE @DateLastStockTake datetime
DECLARE @QtyLastStockTake INT
DECLARE @LastReceived datetime
DECLARE @FirstReceived datetime

SELECT @pid  = i.P_ID,
        @QtyOnHand = i.P_QTYONHAND,
@QtyReserved = i.P_QTYReserved,
@QtyOnBackOrder = i.P_QtyOnBackOrder,
@QtyOnOrder = i.P_QtyOnOrder,
@QtyCopiesOnHand = i.P_QtyCopiesOnHand,
@QtyCopiesReserved = 0,
@QtyOnAppro = i.P_QtyOnAppro,
@DateLastStockTake = i.P_DateLastStockTake,
@QtyLastStockTake = i.P_QtyLastStockTake,
@LastReceived = i.P_LastdateDelivered


FROM INSERTED i

SELECT @DEFSTID  = c.CF_DEFAULTSTOREID 
FROM tCONFIGURATION c
INSERT INTO tSTOREP  (STP_P_ID,STP_ST_ID,STP_QTYONHAND,STP_QtyReserved) VALUES (@pid,@DEFSTID,@QtyOnHand,@QtyReserved)

IF @@ERROR <> 0   RAISERROR (''Trigger StoreP'', 16, 1, '''', '''')')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1

exec('CREATE TRIGGER [dbo].[trig_PriceChange] ON dbo.tProduct
FOR UPDATE
AS
BEGIN

BEGIN TRY

	IF update(P_SP) 
	BEGIN

		INSERT INTO tPriceChange 
		(PCH_PID,PCH_Date,PCH_Price)
		SELECT	ins.P_ID,GetDate(),
			ins.P_SP
		FROM inserted ins Left JOIN DELETED del ON ins.P_ID = del.P_ID
		WHERE ins.P_SP <> ISNULL(del.P_SP,0) AND ins.P_SP IS NOT NULL
		
	END
	SELECT 1
END TRY

BEGIN CATCH

	DECLARE @ErrorString NVARCHAR(4000)

	IF ERROR_PROCEDURE() <> OBJECT_NAME(@@PROCID)
	BEGIN
		PRINT ''X''
		SELECT @ErrorString = ERROR_MESSAGE() + '' X'';
		PRINT @ERRORSTRING
	END
	ELSE
	BEGIN
		PRINT ''Y''
		SELECT @ErrorString = ''Error Message: '' + ERROR_MESSAGE() + 
				'', Error Procedure: '' + RTRIM(CONVERT(CHAR,ERROR_PROCEDURE())) + '', Error Line: '' + RTRIM(CONVERT(Char,ERROR_LINE ()));
		SELECT @ERRORSTRING =  ERROR_MESSAGE() + '' X''
	END
	RAISERROR (@ERRORSTRING, 16,1)
END CATCH
END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1

exec('CREATE TRIGGER [dbo].[trigProdDelete] ON dbo.tProduct 
FOR DELETE 

		
AS
		INSERT INTO tProdUpdates 
		(PRU_Log_Type,
		PRU_P_ID		)
		SELECT	''DEL'',
			del.P_ID
		FROM Deleted del

IF dbo.GetProperty(''SUPPORTSMULTIBUY'') = ''TRUE''
BEGIN
DECLARE @ID INT
DECLARE @MBC_ID INT
DECLARE @OLDVAL VARCHAR(10)
DECLARE @PID UNiqueidentifier
		SELECT @PID = del.P_ID from deleted del
	--Add to or remove from multibuy category if updated
		SELECT @ID = ID FROM tDICTTYPES WHERE [SYSTEM] = ''SE''
		SELECT @MBC_ID =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''MBC'')
		DELETE FROM tPRODUCTSECTION WHERE PSEC_P_ID = @PID AND PSEC_SEC_ID = @MBC_ID
END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1

exec('CREATE TRIGGER [dbo].[trigProdUpdate] ON dbo.tProduct
FOR INSERT, UPDATE
AS


		INSERT INTO tProdUpdates 
		(PRU_Log_Type,
		PRU_P_ID,
		PRU_Code,
		PRU_EAN,
		PRU_Publisher,
		PRU_SeriesTitle,
		PRU_MainAuthor,
		PRU_Title,
		PRU_SP,
		PRU_SSP,
		PRU_VATRate,
		PRU_TriggerDate,
		PRU_PTID,
		PRU_SECID,
		PRU_NDA,
		PRU_MULTIBUYCODE)
		SELECT	''NEW'',
			ins.P_ID,
			ins.P_Code,
			ins.P_EAN,
			ins.P_Publisher,
			ins.P_SeriesTitle,
			ins.P_MainAuthor,
			LEFT(ins.P_Title,250),
			ins.P_SP,
			ins.P_SPecial,
			dbo.VATRATETOUSE(ins.P_SpecialVat,ins.P_VatRate),
			GetDate(),
			ins.P_ProductType_ID,
			vSectionMaster.PSEC_SEC_ID,
			ins.P_NDA,
			ins.P_MultibuyCode
		FROM inserted ins LEFT JOIN vSectionMaster ON ins.P_ID = vSectionMaster.PSEC_P_ID 
		Left JOIN deleted del ON ins.P_ID = del.P_ID 
		WHERE		ISNULL(ins.P_CODE,'''') <> ISNULL(del.P_Code,'''') or
					ISNULL(ins.P_EAN,'''') <> ISNULL(del.P_EAN,'''') or
					ISNULL(ins.P_Publisher,'''') <> ISNULL(del.P_Publisher,'''') or
					ISNULL(ins.P_SeriesTitle,'''') <> ISNULL(del.P_SeriesTitle,'''') or
					ISNULL(ins.P_MainAuthor,'''') <> ISNULL(del.P_MainAuthor,'''') or
					ISNULL(ins.P_Title,'''') <> ISNULL(del.P_Title,'''') or
					ISNULL(ins.P_SP,0) <> ISNULL(del.P_SP,0) or
					ISNULL(ins.P_SPecial,0) <> ISNULL(del.P_SPecial,0) or
					ISNULL(ins.P_VatRate,0) <> ISNULL(del.P_VatRate,0) or
					ISNULL(ins.P_SpecialVat,0) <> ISNULL(del.P_SpecialVat,0) or
					ISNULL(ins.P_ProductType_ID,0) <> ISNULL(del.P_ProductType_ID,0) or
					ISNULL(ins.P_NDA,'''') <> ISNULL(del.P_NDA,'''') or
					ISNULL(ins.P_MultibuyCode,'''') <> ISNULL(del.P_MultibuyCode,'''') or
					UPDATE (P_CODE)


IF dbo.GetProperty(''SUPPORTSMULTIBUY'') = ''TRUE''
BEGIN
DECLARE @ID INT
DECLARE @MBC_ID INT
DECLARE @OLDVAL VARCHAR(10)
DECLARE @NEWVAL VARCHAR(10)
DECLARE @PID UNiqueidentifier
		SELECT @OLDVAL = del.P_MultibuyCode from Deleted del
		SELECT @NEWVAL = ins.P_MultibuyCode,@PID = ins.P_ID from inserted ins
	--Add to or remove from multibuy category if updated
		IF ISNULL(@NEWVAL,'''') <> ISNULL(@OLDVAL,'''') 
		BEGIN
			SELECT @ID = ID FROM tDICTTYPES WHERE [SYSTEM] = ''SE''
			SELECT @MBC_ID =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''MBC'')
			IF ISNULL(@NEWVAL,'''') > ''''
			BEGIN
				IF	NOT (SELECT COUNT( PSEC_P_ID) FROM tProductSection WHERE PSEC_P_ID = @PID and PSEC_SEC_ID = @MBC_ID) > 0
					INSERT INTO tProductSection (PSEC_P_ID,PSEC_SEC_ID,PSEC_Priority)
					Values	(@PID,@MBC_ID,0) 
				UPDATE tProduct SET P_NDA = 1 WHERE P_ID = @PID
			END
			ELSE
			BEGIN
				UPDATE tProduct SET P_NDA = 0 WHERE P_ID = @PID
				DELETE FROM tPRODUCTSECTION WHERE PSEC_P_ID = @PID AND PSEC_SEC_ID = @MBC_ID
			END
		END		

			


END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1

exec('CREATE TRIGGER [dbo].[trigQtyOHChange] ON dbo.tProduct
AFTER INSERT, UPDATE, DELETE 
AS
BEGIN
DECLARE @QtyOHBody XML
DECLARE @DMLType CHAR(1)	
DECLARE @INSTALLATIONCODE VARCHAR(10)
DECLARE @CMD VARCHAR(200)
DECLARE @TMP VARCHAR(MAX)
DECLARE @RES INT
DECLARE @ERRMESS VARCHAR(500)
DECLARE @SOH INT
DECLARE @SOO INT
DECLARE @LASTDELIVEREDPRICE INT
DECLARE @SP INT
DECLARE @LASTDATEORDERED DATETIME
DECLARE @LASTQTYFIRMORDERED INT
DECLARE @LASTQTYSSORDERED INT

DECLARE @TOTALQTYSOLD INT
DECLARE @LASTDELIVEREDDATE DATETIME
DECLARE @LASTDATESOLD DATETIME
DECLARE @LASTQTYDELIVERED INT

DECLARE @PEAN VARCHAR(20)
DECLARE @PID UNIQUEIDENTIFIER
DECLARE @STORECODE VARCHAR(5)
DECLARE @ErrorString NVARCHAR(4000)

--declare @$prog varchar(50), 
--	@$errno int, 
--	@$errmsg varchar(4000), 
--	@$proc_section_nm varchar(50),
--	@$row_cnt INT,
--	@$error_db_name varchar(50), 
--	@$CreateUserName varchar(128),   -- last user changed the data 
--	@$CreateMachineName varchar(128), -- last machine changes-procedure were run from
--	@$CreateSource varchar(128)		-- last process that made a changes
--
--select @$errno = NULL,  @$errmsg = NULL,  @$proc_section_nm = NULL
--	,  @$prog = LEFT(object_name(@@procid),50), @$row_cnt = NULL
--	, @$error_db_name = db_name();
--

BEGIN TRY
	IF dbo.GETPROPERTY(''SOH_ON'') <>''TRUE''
		RETURN
	
	SELECT @INSTALLATIONCODE =  CF_INSTALLATIONCODE FROM tCONFIGURATION
	IF NOT EXISTS (SELECT * FROM inserted)
	BEGIN	
		SELECT	@PID = NULL
		SELECT @SOH = 0
		SELECT @PEAN = ''''
		SELECT @DMLType = ''D''
	END 
	-- after update or insert statement
	ELSE
	BEGIN
		SELECT	@PID = ISNULL(P_ID,''''), 
				@SOH = ISNULL(P_QTYONHAND,0), 
				@SOO = ISNULL(P_QTYONORDER,0), 
				@LASTDELIVEREDPRICE = ISNULL(P_LastPriceDelivered,0), 
				@SP = ISNULL(P_SP,0), 
				@LASTDATEORDERED = ISNULL(P_LastDateOrdered,0), 
				@LASTQTYFIRMORDERED = ISNULL(P_LastQtyFirmOrdered,0), 
				@LASTQTYSSORDERED = ISNULL(P_LastQtySSOrdered,0), 

				@TOTALQTYSOLD = ISNULL(P_QtyTotalSold,0), 
				@LASTDELIVEREDDATE = ISNULL(P_LastDateDelivered,0), 
				@LASTDATESOLD = ISNULL(P_LastDateSold,0), 
				@LASTQTYDELIVERED = ISNULL(P_LastQtyDelivered,0), 

				@PEAN = ISNULL(P_EAN,'''') 
		FROM inserted
		-- after update statement
		IF EXISTS (SELECT * FROM deleted)
			SELECT 	@DMLType = ''U''
		ELSE
			SELECT	@DMLType = ''I''
	END
	SELECT @QtyOHBody = 
		''<SOHMsg>
			<DMLType>'' + @DMLType + ''</DMLType>
			<PID>'' + CAST(@PID AS VARCHAR(40)) + ''</PID>
			<SOH>'' + CAST(@SOH as VARCHAR(10)) + ''</SOH>
			<SOO>'' + CAST(@SOO as VARCHAR(10)) + ''</SOO>
			<LDP>'' + CAST(@LASTDELIVEREDPRICE as VARCHAR(10)) + ''</LDP>
			<SP>'' + CAST(@SP as VARCHAR(10)) + ''</SP>
			<LDO>'' + CONVERT(VARCHAR(20),@LASTDATEORDERED,120) + ''</LDO>
			<QTYFIRMORDERED>'' + CAST(@LASTQTYFIRMORDERED as VARCHAR(10)) + ''</QTYFIRMORDERED>
			<QTYSSORDERED>'' + CAST(@LASTQTYSSORDERED as VARCHAR(10)) + ''</QTYSSORDERED>

			<TOTALSOLD>'' + CAST(@TOTALQTYSOLD as VARCHAR(10)) + ''</TOTALSOLD>
			<LDREC>'' + CONVERT(VARCHAR(20),@LASTDELIVEREDDATE,120) + ''</LDREC>
			<LDSOLD>'' + CONVERT(VARCHAR(20),@LASTDATESOLD,120) + ''</LDSOLD>
			<QTYLASTDELIVERED>'' + CAST(@LASTQTYDELIVERED as VARCHAR(10)) + ''</QTYLASTDELIVERED>

			<EAN>'' + CAST(@PEAN as VARCHAR(20)) + ''</EAN>
			<STCODE>'' + @INSTALLATIONCODE + ''</STCODE>
		</SOHMsg>''
	--INSERT INTO _tSBLog (SBL_Msg,SBL_Proc) VALUES (ISNULL(CAST(@QtyOHBody as VARCHAR(580)),''NULL''),''trigQtyOHChange'')

	SELECT @CMD = ''SOHSOURCE_'' + @INSTALLATIONCODE + ''_SERVICE''

	IF NOT @QTYOHBODY IS NULL 
		EXEC dbo._usp_SendXML @CMD,''SOHCONSUMER_SERVICE'',''SOH_CONTRACT'', ''SOH_MSG'', @QtyOHBody,@RES,@ERRMESS
END TRY
BEGIN CATCH
		IF ERROR_PROCEDURE() <> OBJECT_NAME(@@PROCID)
			SELECT @ErrorString = ERROR_MESSAGE();
		ELSE
			SELECT @ErrorString = ''Error Message: '' + ERROR_MESSAGE() + '', Error Procedure: '' + 
						ERROR_PROCEDURE() + '', Error Line: '' + CAST(ERROR_LINE() as VARCHAR(10));
		
		RAISERROR (@ErrorString, 16,1)
END CATCH


SET NOCOUNT OFF; 


END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1

exec('CREATE TRIGGER dbo.UniqueCode ON dbo.tProduct
FOR INSERT, UPDATE
AS
BEGIN
SET NOCOUNT ON 
IF UPDATE(P_CODE)
 IF (SELECT MAX(cnt) FROM (SELECT COUNT(i.P_CODE) as cnt from tPRODUCT,
  inserted i WHERE tPRODUCT.P_CODE=i.P_CODE AND i.P_CODE > ''''  GROUP BY i.P_CODE)x)>1
BEGIN
  ROLLBACK TRAN
  RAISERROR (''Duplicate CODE'', 16, 1)
END
IF UPDATE(P_EAN)
 IF (SELECT MAX(cnt) FROM (SELECT COUNT(i.P_EAN) as cnt from tPRODUCT,
  inserted i WHERE tPRODUCT.P_EAN=i.P_EAN GROUP BY i.P_EAN)x)>1
BEGIN
  ROLLBACK TRAN
  RAISERROR (''Duplicate EAN'', 16, 1)
END
END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.tProduct Table Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.tProduct Table'
END
GO


--
-- Script To Update dbo.tREORDER1 Table In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.tREORDER1 Table'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE TABLE [dbo].[tmp_tREORDER1] (
   [REORDD_ID] [int] IDENTITY (1, 1) NOT NULL,
   [STatus] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [PID] [uniqueidentifier] NULL,
   [COLID] [int] NULL,
   [REF] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [QTYCO] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [QTYPO] [int] NULL,
   [QtyPOUnissued] [int] NULL,
   [QTYAPP] [int] NULL,
   [LASTSIXMONTHS] [varchar] (60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [LASTSIXWEEKS] [varchar] (60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [QTYFIRM] [int] NULL,
   [QTYSS] [int] NULL,
   [PRCODE] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [DESCRIP] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [LASTSUPPLIERID] [int] NULL,
   [LASTSUPPLIERNAME] [varchar] (55) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [LASTDEALID] [int] NULL,
   [LASTDEALNAME] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [PT] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [PUBLISHER] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [TOTALSOLD] [int] NULL,
   [PRICE] [int] NULL,
   [ONHAND] [int] NULL,
   [LASTRECEIVEDDATE] [datetime] NULL,
   [LASTORDEREDDATE] [datetime] NULL,
   [LASTRECEIVEDQTY] [int] NULL,
   [LASTORDEREDQTYFIRM] [int] NULL,
   [LASTORDEREDQTYSS] [int] NULL,
   [LASTRECEIVEDPRICE] [int] NULL,
   [LASTORDEREDPRICE] [int] NULL,
   [CODate] [datetime] NULL,
   [STAFFID] [int] NULL
)
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   SET IDENTITY_INSERT [dbo].[tmp_tREORDER1] ON
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   INSERT INTO [dbo].[tmp_tREORDER1] ([REORDD_ID], [STatus], [PID], [COLID], [REF], [QTYCO], [QTYPO], [QtyPOUnissued], [QTYAPP], [LASTSIXMONTHS], [LASTSIXWEEKS], [QTYFIRM], [QTYSS], [PRCODE], [DESCRIP], [LASTSUPPLIERID], [LASTSUPPLIERNAME], [LASTDEALID], [LASTDEALNAME], [PT], [PUBLISHER], [TOTALSOLD], [PRICE], [ONHAND], [LASTRECEIVEDDATE], [LASTORDEREDDATE], [LASTRECEIVEDQTY], [LASTORDEREDQTYFIRM], [LASTORDEREDQTYSS], [LASTRECEIVEDPRICE], [LASTORDEREDPRICE], [CODate], [STAFFID])
   SELECT [REORDD_ID], [STatus], [PID], [COLID], [REF], [QTYCO], [QTYPO], NULL, [QTYAPP], [LASTSIXMONTHS], [LASTSIXWEEKS], [QTYFIRM], [QTYSS], [PRCODE], [DESCRIP], [LASTSUPPLIERID], [LASTSUPPLIERNAME], [LASTDEALID], [LASTDEALNAME], [PT], [PUBLISHER], [TOTALSOLD], [PRICE], [ONHAND], [LASTRECEIVEDDATE], [LASTORDEREDDATE], [LASTRECEIVEDQTY], [LASTORDEREDQTYFIRM], [LASTORDEREDQTYSS], [LASTRECEIVEDPRICE], [LASTORDEREDPRICE], [CODate], [STAFFID]
   FROM [dbo].[tREORDER1]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   SET IDENTITY_INSERT [dbo].[tmp_tREORDER1] OFF
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   DROP TABLE [dbo].[tREORDER1]
GO

sp_rename N'[dbo].[tmp_tREORDER1]', N'tREORDER1'

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.tREORDER1 Table Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.tREORDER1 Table'
END
GO

--
-- Script To Update dbo.tREORDER2 Table In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.tREORDER2 Table'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE TABLE [dbo].[tmp_tREORDER2] (
   [REORDD_ID] [int] IDENTITY (1, 1) NOT NULL,
   [STatus] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [CODate] [datetime] NULL,
   [PID] [uniqueidentifier] NULL,
   [COLID] [int] NULL,
   [REF] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [QTYCO] [int] NULL,
   [QTYPO] [int] NULL,
   [QtyPOUnissued] [int] NULL,
   [QTYAPP] [int] NULL,
   [LASTSIXMONTHS] [varchar] (60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [LASTSIXWEEKS] [varchar] (60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [QTYFIRM] [int] NULL,
   [QTYSS] [int] NULL,
   [PRCODE] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [DESCRIP] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [LASTSUPPLIERID] [int] NULL,
   [LASTSUPPLIERNAME] [varchar] (55) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [LASTDEALID] [int] NULL,
   [LASTDEALNAME] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [PT] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [PUBLISHER] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [TOTALSOLD] [int] NULL,
   [PRICE] [int] NULL,
   [ONHAND] [int] NULL,
   [LASTRECEIVEDDATE] [datetime] NULL,
   [LASTORDEREDDATE] [datetime] NULL,
   [LASTRECEIVEDQTY] [int] NULL,
   [LASTORDEREDQTYFIRM] [int] NULL,
   [LASTORDEREDQTYSS] [int] NULL,
   [LASTRECEIVEDPRICE] [int] NULL,
   [LASTORDEREDPRICE] [int] NULL
)
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   SET IDENTITY_INSERT [dbo].[tmp_tREORDER2] ON
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   INSERT INTO [dbo].[tmp_tREORDER2] ([REORDD_ID], [STatus], [CODate], [PID], [COLID], [REF], [QTYCO], [QTYPO], [QtyPOUnissued], [QTYAPP], [LASTSIXMONTHS], [LASTSIXWEEKS], [QTYFIRM], [QTYSS], [PRCODE], [DESCRIP], [LASTSUPPLIERID], [LASTSUPPLIERNAME], [LASTDEALID], [LASTDEALNAME], [PT], [PUBLISHER], [TOTALSOLD], [PRICE], [ONHAND], [LASTRECEIVEDDATE], [LASTORDEREDDATE], [LASTRECEIVEDQTY], [LASTORDEREDQTYFIRM], [LASTORDEREDQTYSS], [LASTRECEIVEDPRICE], [LASTORDEREDPRICE])
   SELECT [REORDD_ID], [STatus], [CODate], [PID], [COLID], [REF], [QTYCO], [QTYPO], NULL, [QTYAPP], [LASTSIXMONTHS], [LASTSIXWEEKS], [QTYFIRM], [QTYSS], [PRCODE], [DESCRIP], [LASTSUPPLIERID], [LASTSUPPLIERNAME], [LASTDEALID], [LASTDEALNAME], [PT], [PUBLISHER], [TOTALSOLD], [PRICE], [ONHAND], [LASTRECEIVEDDATE], [LASTORDEREDDATE], [LASTRECEIVEDQTY], [LASTORDEREDQTYFIRM], [LASTORDEREDQTYSS], [LASTRECEIVEDPRICE], [LASTORDEREDPRICE]
   FROM [dbo].[tREORDER2]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   SET IDENTITY_INSERT [dbo].[tmp_tREORDER2] OFF
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   DROP TABLE [dbo].[tREORDER2]
GO

sp_rename N'[dbo].[tmp_tREORDER2]', N'tREORDER2'

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.tREORDER2 Table Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.tREORDER2 Table'
END
GO

--
-- Script To Update dbo.tREORDERCUSTByCOL Table In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.tREORDERCUSTByCOL Table'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tREORDERCUSTByCOL_ID')
      ALTER TABLE [dbo].[tREORDERCUSTByCOL] DROP CONSTRAINT [DF_tREORDERCUSTByCOL_ID]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE TABLE [dbo].[tmp_tREORDERCUSTByCOL] (
   [WSNAME] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [STAFFID] [int] NULL,
   [COLID] [int] NULL,
   [STatus] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [PID] [uniqueidentifier] NULL,
   [REF] [varchar] (35) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [QTYCO] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [QTYPO] [int] NULL,
   [QtyPOUnissued] [int] NULL,
   [QTYAPP] [int] NULL,
   [LASTSIXMONTHS] [varchar] (60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [LASTSIXWEEKS] [varchar] (60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [QTYFIRM] [int] NULL,
   [QTYSS] [int] NULL,
   [PRCODE] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [DESCRIP] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [LASTSUPPLIERID] [int] NULL,
   [LASTSUPPLIERNAME] [varchar] (55) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [LASTDEALID] [int] NULL,
   [LASTDEALNAME] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [PT] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [PUBLISHER] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [TOTALSOLD] [int] NULL,
   [PRICE] [int] NULL,
   [ONHAND] [int] NULL,
   [LASTRECEIVEDDATE] [datetime] NULL,
   [LASTORDEREDDATE] [datetime] NULL,
   [LASTRECEIVEDQTY] [int] NULL,
   [LASTORDEREDQTYFIRM] [int] NULL,
   [LASTORDEREDQTYSS] [int] NULL,
   [LASTRECEIVEDPRICE] [int] NULL,
   [LASTORDEREDPRICE] [int] NULL,
   [CODate] [datetime] NULL,
   [TitleForSort] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [SlateName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [ForeignPrice] [int] NULL,
   [Category] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [ID] [varbinary] (50) NOT NULL CONSTRAINT [DF_tREORDERCUSTByCOL_ID] DEFAULT (newid())
)
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   INSERT INTO [dbo].[tmp_tREORDERCUSTByCOL] ([WSNAME], [STAFFID], [COLID], [STatus], [PID], [REF], [QTYCO], [QTYPO], [QtyPOUnissued], [QTYAPP], [LASTSIXMONTHS], [LASTSIXWEEKS], [QTYFIRM], [QTYSS], [PRCODE], [DESCRIP], [LASTSUPPLIERID], [LASTSUPPLIERNAME], [LASTDEALID], [LASTDEALNAME], [PT], [PUBLISHER], [TOTALSOLD], [PRICE], [ONHAND], [LASTRECEIVEDDATE], [LASTORDEREDDATE], [LASTRECEIVEDQTY], [LASTORDEREDQTYFIRM], [LASTORDEREDQTYSS], [LASTRECEIVEDPRICE], [LASTORDEREDPRICE], [CODate], [TitleForSort], [SlateName], [ForeignPrice], [Category], [ID])
   SELECT [WSNAME], [STAFFID], [COLID], [STatus], [PID], [REF], [QTYCO], [QTYPO], NULL, [QTYAPP], [LASTSIXMONTHS], [LASTSIXWEEKS], [QTYFIRM], [QTYSS], [PRCODE], [DESCRIP], [LASTSUPPLIERID], [LASTSUPPLIERNAME], [LASTDEALID], [LASTDEALNAME], [PT], [PUBLISHER], [TOTALSOLD], [PRICE], [ONHAND], [LASTRECEIVEDDATE], [LASTORDEREDDATE], [LASTRECEIVEDQTY], [LASTORDEREDQTYFIRM], [LASTORDEREDQTYSS], [LASTRECEIVEDPRICE], [LASTORDEREDPRICE], [CODate], [TitleForSort], [SlateName], [ForeignPrice], [Category], [ID]
   FROM [dbo].[tREORDERCUSTByCOL]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   DROP TABLE [dbo].[tREORDERCUSTByCOL]
GO

sp_rename N'[dbo].[tmp_tREORDERCUSTByCOL]', N'tREORDERCUSTByCOL'

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   ALTER TABLE [dbo].[tREORDERCUSTByCOL] ADD CONSTRAINT [PK_tREORDERCUSTByCOL] PRIMARY KEY CLUSTERED ([ID])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE INDEX [iUnique] ON [dbo].[tREORDERCUSTByCOL] ([WSNAME], [STAFFID], [COLID])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.tREORDERCUSTByCOL Table Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.tREORDERCUSTByCOL Table'
END
GO

--
-- Script To Update dbo.tREORDERGENERAL Table In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.tREORDERGENERAL Table'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF EXISTS (SELECT name FROM sysobjects WHERE name = N'DF_tREORDERGENERAL_ID')
      ALTER TABLE [dbo].[tREORDERGENERAL] DROP CONSTRAINT [DF_tREORDERGENERAL_ID]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE TABLE [dbo].[tmp_tREORDERGENERAL] (
   [WSNAME] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [STatus] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [CODate] [datetime] NULL,
   [PID] [uniqueidentifier] NULL,
   [COLID] [int] NULL,
   [REF] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [QTYCO] [int] NULL,
   [QTYPO] [int] NULL,
   [QtyPOUnissued] [int] NULL,
   [QTYAPP] [int] NULL,
   [QTYFIRM] [int] NULL,
   [QTYSS] [int] NULL,
   [PRCODE] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [DESCRIP] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [LASTSUPPLIERID] [int] NULL,
   [LASTSUPPLIERNAME] [varchar] (55) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [LASTDEALID] [int] NULL,
   [LASTDEALNAME] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [PT] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [PUBLISHER] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [TOTALSOLD] [int] NULL,
   [PRICE] [int] NULL,
   [ONHAND] [int] NULL,
   [LASTRECEIVEDDATE] [datetime] NULL,
   [LASTORDEREDDATE] [datetime] NULL,
   [LASTRECEIVEDQTY] [int] NULL,
   [LASTORDEREDQTYFIRM] [int] NULL,
   [LASTORDEREDQTYSS] [int] NULL,
   [LASTRECEIVEDPRICE] [int] NULL,
   [LASTORDEREDPRICE] [int] NULL,
   [LASTSIXMONTHS] [varchar] (60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [LASTSIXWEEKS] [varchar] (60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [TITLEForSORT] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [SLATENAME] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [ForeignPrice] [int] NULL,
   [Category] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [ID] [uniqueidentifier] NOT NULL CONSTRAINT [DF_tREORDERGENERAL_ID] DEFAULT (newid())
)
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   INSERT INTO [dbo].[tmp_tREORDERGENERAL] ([WSNAME], [STatus], [CODate], [PID], [COLID], [REF], [QTYCO], [QTYPO], [QtyPOUnissued], [QTYAPP], [QTYFIRM], [QTYSS], [PRCODE], [DESCRIP], [LASTSUPPLIERID], [LASTSUPPLIERNAME], [LASTDEALID], [LASTDEALNAME], [PT], [PUBLISHER], [TOTALSOLD], [PRICE], [ONHAND], [LASTRECEIVEDDATE], [LASTORDEREDDATE], [LASTRECEIVEDQTY], [LASTORDEREDQTYFIRM], [LASTORDEREDQTYSS], [LASTRECEIVEDPRICE], [LASTORDEREDPRICE], [LASTSIXMONTHS], [LASTSIXWEEKS], [TITLEForSORT], [SLATENAME], [ForeignPrice], [Category], [ID])
   SELECT [WSNAME], [STatus], [CODate], [PID], [COLID], [REF], [QTYCO], [QTYPO], NULL, [QTYAPP], [QTYFIRM], [QTYSS], [PRCODE], [DESCRIP], [LASTSUPPLIERID], [LASTSUPPLIERNAME], [LASTDEALID], [LASTDEALNAME], [PT], [PUBLISHER], [TOTALSOLD], [PRICE], [ONHAND], [LASTRECEIVEDDATE], [LASTORDEREDDATE], [LASTRECEIVEDQTY], [LASTORDEREDQTYFIRM], [LASTORDEREDQTYSS], [LASTRECEIVEDPRICE], [LASTORDEREDPRICE], [LASTSIXMONTHS], [LASTSIXWEEKS], [TITLEForSORT], [SLATENAME], [ForeignPrice], [Category], [ID]
   FROM [dbo].[tREORDERGENERAL]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   DROP TABLE [dbo].[tREORDERGENERAL]
GO

sp_rename N'[dbo].[tmp_tREORDERGENERAL]', N'tREORDERGENERAL'

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   ALTER TABLE [dbo].[tREORDERGENERAL] ADD CONSTRAINT [PK_tREORDERGENERAL] PRIMARY KEY CLUSTERED ([ID])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.tREORDERGENERAL Table Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.tREORDERGENERAL Table'
END
GO

--
-- Script To Create dbo.ahv_InvoicesPerSalesperson View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.ahv_InvoicesPerSalesperson View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.ahv_InvoicesPerSalesperson
AS
SELECT     dbo.tTR.TR_Code AS DocumentInvoicede, dbo.tTP.TP_ACNo AS CustomerAccountInvoicede, dbo.tTP.TP_Name AS CustomerName, 
                      dbo.tTP.TP_Initials AS CustomerInitials, dbo.tTP.TP_Title AS CustomerTitle, dbo.tTR.TR_Date AS DocumentNominalDate, 
                      dbo.tTR.TR_CaptureDate AS DocumentFirstCapturedDate, dbo.tTR.TR_ProcessingDate AS DocumentIssuedDate, dbo.CalcExtEXVAT2(dbo.tILine.IL_Qty, 
                      dbo.tILine.IL_Price, dbo.tILine.IL_DiscountRate, dbo.tProduct.P_VATRate, dbo.tCurrency.CURR_Divisor) AS OrderValExVat, 
                      dbo.CalcExtEXVAT2(dbo.tILine.IL_Qty, dbo.tILine.IL_Price, dbo.tILine.IL_DiscountRate, 0, dbo.tCurrency.CURR_Divisor) AS OrderValIncVat, 
                      dbo.tStaffMember.SM_Name AS StaffpersonName
FROM         dbo.tInvoice INNER JOIN
                      dbo.tTR INNER JOIN
                      dbo.tILine ON dbo.tTR.TR_ID = dbo.tILine.IL_TR_ID INNER JOIN
                      dbo.tStaffMember ON dbo.tTR.TR_STAFFID = dbo.tStaffMember.SM_ID INNER JOIN
                      dbo.tTP ON dbo.tTR.TR_TP_ID = dbo.tTP.TP_ID INNER JOIN
                      dbo.tProduct ON dbo.tILine.IL_P_ID = dbo.tProduct.P_ID ON dbo.tInvoice.I_ID = dbo.tTR.TR_ID CROSS JOIN
                      dbo.tCurrency INNER JOIN
                      dbo.tConfiguration ON dbo.tCurrency.CURR_ID = dbo.tConfiguration.CF_DefaultCurrID
WHERE     (dbo.tTR.TR_Status IN (3, 4))')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.ahv_InvoicesPerSalesperson View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.ahv_InvoicesPerSalesperson View'
END
GO

--
-- Script To Create dbo.Budget_Budget View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.Budget_Budget View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.Budget_Budget
AS
SELECT     TOP (100) PERCENT dbo.StartOfMonth(dbo.tBudget.B_BudgetMonth) AS BudgetMonthMonth, dbo.tBudget.B_DeliveriesBudget AS BudgetDeliveries, 
                      dbo.tBudget.B_ReturnsBudget AS BudgetReturns
FROM         dbo.tCurrency INNER JOIN
                      dbo.tConfiguration ON dbo.tCurrency.CURR_ID = dbo.tConfiguration.CF_DefaultCurrID CROSS JOIN
                      dbo.tBudget
ORDER BY BudgetMonthMonth DESC')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.Budget_Budget View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.Budget_Budget View'
END
GO

--
-- Script To Create dbo.Budget_Deliveries View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.Budget_Deliveries View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.Budget_Deliveries
AS
SELECT     TOP (100) PERCENT dbo.StartOfMonth(dbo.tDEL.DEL_SupplierInvoiceDate) AS DeliveryMonth, 
                      SUM(dbo.CurrFormat(dbo.CalcExt2(CASE WHEN TR_STATUS IN (3, 4) THEN DELL_QtySS + DELL_QtyFirm ELSE 0 END, dbo.tDELL.DELL_Price, 0, 
                      dbo.tCurrency.CURR_Divisor))) AS RetailValueReceivedIssued, 
                      SUM(dbo.CurrFormat(dbo.CalcExt2(CASE WHEN TR_STATUS = 2 THEN DELL_QtySS + DELL_QtyFirm ELSE 0 END, dbo.tDELL.DELL_Price, 0, 
                      dbo.tCurrency.CURR_Divisor))) AS RetailValueReceivedInProcess
FROM         dbo.tDELL INNER JOIN
                      dbo.tTR ON dbo.tDELL.DELL_TR_ID = dbo.tTR.TR_ID INNER JOIN
                      dbo.tDEL ON dbo.tTR.TR_ID = dbo.tDEL.DEL_ID CROSS JOIN
                      dbo.tCurrency INNER JOIN
                      dbo.tConfiguration ON dbo.tCurrency.CURR_ID = dbo.tConfiguration.CF_DefaultCurrID
GROUP BY dbo.StartOfMonth(dbo.tDEL.DEL_SupplierInvoiceDate)
ORDER BY dbo.StartOfMonth(dbo.tDEL.DEL_SupplierInvoiceDate) DESC')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.Budget_Deliveries View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.Budget_Deliveries View'
END
GO

--
-- Script To Create dbo.Budget_LastFourMonths_1 View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.Budget_LastFourMonths_1 View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.Budget_LastFourMonths_1
AS
SELECT     TOP (100) PERCENT dbo.StartOfMonth(GETDATE()) AS DeliveryMonth, SUM(dbo.CurrFormat(dbo.CalcExt2(CASE WHEN TR_STATUS IN (3, 4) 
                      THEN DELL_QtySS + DELL_QtyFirm ELSE 0 END, dbo.tDELL.DELL_Price, 0, 1))) AS RetailValueReceivedIssued, 
                      SUM(dbo.CurrFormat(dbo.CalcExt2(CASE WHEN TR_STATUS = 2 THEN DELL_QtySS + DELL_QtyFirm ELSE 0 END, dbo.tDELL.DELL_Price, 0, 1))) 
                      AS RetailValueReceivedInProcess
FROM         dbo.tDEL AS tDEL_1 INNER JOIN
                      dbo.tDELL ON tDEL_1.DEL_ID = dbo.tDELL.DELL_TR_ID INNER JOIN
                      dbo.tTR ON dbo.tDELL.DELL_TR_ID = dbo.tTR.TR_ID CROSS JOIN
                      dbo.tCurrency INNER JOIN
                      dbo.tConfiguration ON dbo.tCurrency.CURR_ID = dbo.tConfiguration.CF_DefaultCurrID
WHERE     (tDEL_1.DEL_SupplierInvoiceDate >= DATEADD(m, - 2, GETDATE()))')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.Budget_LastFourMonths_1 View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.Budget_LastFourMonths_1 View'
END
GO

--
-- Script To Create dbo.Budget_Orders View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.Budget_Orders View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.Budget_Orders
AS
SELECT     TOP (100) PERCENT dbo.StartOfMonth(dbo.tPOL.POL_ETA) AS DeliveryMonth, SUM(dbo.CurrFormat(dbo.CalcExt2(CASE WHEN TR_STATUS IN (3, 4) 
                      THEN dbo.tPOL.POL_QtySS + dbo.tPOL.POL_QtyFirm ELSE 0 END, dbo.tPOL.POL_Price, 0, dbo.tCurrency.CURR_Divisor))) 
                      AS OrdersAtRetailValueIssued, 
                      SUM(dbo.CurrFormat(dbo.CalcExt2(CASE WHEN TR_STATUS = 2 THEN dbo.tPOL.POL_QtySS + dbo.tPOL.POL_QtyFirm ELSE 0 END, 
                      dbo.tPOL.POL_Price, 0, dbo.tCurrency.CURR_Divisor))) AS OrdersAtRetailValueInProcess
FROM         dbo.tCurrency INNER JOIN
                      dbo.tConfiguration ON dbo.tCurrency.CURR_ID = dbo.tConfiguration.CF_DefaultCurrID CROSS JOIN
                      dbo.tTR INNER JOIN
                      dbo.tPOL ON dbo.tTR.TR_ID = dbo.tPOL.POL_TR_ID INNER JOIN
                      dbo.tPO ON dbo.tTR.TR_ID = dbo.tPO.PO_ID
WHERE     (dbo.tPOL.POL_DateReplaced IS NULL)
GROUP BY dbo.StartOfMonth(dbo.tPOL.POL_ETA)
ORDER BY DeliveryMonth DESC')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.Budget_Orders View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.Budget_Orders View'
END
GO

--
-- Script To Create dbo.Budget_Returns View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.Budget_Returns View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.Budget_Returns
AS
SELECT     dbo.StartOfMonth(dbo.tTR.TR_ProcessingDate) AS ReturnMonth, SUM(dbo.CurrFormat(dbo.CalcExt2(CASE WHEN TR_STATUS IN (3, 4) 
                      THEN ISNULL(RL_QTYRETURNED, 0) - ISNULL(RL_QTYREJECTED, 0) ELSE 0 END, ISNULL(dbo.tRL.RL_Price, 0), 0, 100))) 
                      AS RetailValueReturnsIssued, SUM(dbo.CurrFormat(dbo.CalcExt2(CASE WHEN TR_STATUS = 2 THEN ISNULL(RL_QTYRETURNED, 0) 
                      - ISNULL(RL_QTYREJECTED, 0) ELSE 0 END, ISNULL(dbo.tRL.RL_Price, 0), 0, 100))) AS RetailValueReturnsInProcess
FROM         dbo.tRL INNER JOIN
                      dbo.tTR ON dbo.tRL.RL_TR_ID = dbo.tTR.TR_ID INNER JOIN
                      dbo.tR ON dbo.tTR.TR_ID = dbo.tR.R_ID
GROUP BY dbo.StartOfMonth(dbo.tTR.TR_ProcessingDate)')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.Budget_Returns View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.Budget_Returns View'
END
GO

--
-- Script To Create dbo.Budget_Summary View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.Budget_Summary View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.Budget_Summary
AS
SELECT     TOP (100) PERCENT dbo.Budget_Budget.BudgetMonthMonth, dbo.Budget_Budget.BudgetDeliveries, dbo.Budget_Budget.BudgetReturns, 
                      dbo.Budget_Orders.OrdersAtRetailValueIssued, dbo.Budget_Orders.OrdersAtRetailValueInProcess, 
                      (dbo.Budget_Orders.OrdersAtRetailValueIssued + dbo.Budget_Orders.OrdersAtRetailValueInProcess) 
                      / dbo.Budget_Budget.BudgetDeliveries AS OrdersAgainstBudget, dbo.Budget_Deliveries.RetailValueReceivedIssued, 
                      dbo.Budget_Deliveries.RetailValueReceivedInProcess, 
                      (dbo.Budget_Deliveries.RetailValueReceivedIssued + dbo.Budget_Deliveries.RetailValueReceivedInProcess) 
                      / dbo.Budget_Budget.BudgetDeliveries AS DeliveriesAgainstBudget
FROM         dbo.Budget_Budget INNER JOIN
                      dbo.Budget_Orders ON dbo.Budget_Budget.BudgetMonthMonth = dbo.Budget_Orders.DeliveryMonth INNER JOIN
                      dbo.Budget_Deliveries ON dbo.Budget_Orders.DeliveryMonth = dbo.Budget_Deliveries.DeliveryMonth CROSS JOIN
                      dbo.Budget_LastFourMonths_1
ORDER BY dbo.Budget_Budget.BudgetMonthMonth')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.Budget_Summary View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.Budget_Summary View'
END
GO

--
-- Script To Delete dbo.vACCPACExport_CUST View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Deleting dbo.vACCPACExport_CUST View'
GO

   DROP VIEW [dbo].[vACCPACExport_CUST]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vACCPACExport_CUST View Deleted Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Delete dbo.vACCPACExport_CUST View'
END
GO

--
-- Script To Delete dbo.vACCPACExport_Cust_Short View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Deleting dbo.vACCPACExport_Cust_Short View'
GO

   DROP VIEW [dbo].[vACCPACExport_Cust_Short]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vACCPACExport_Cust_Short View Deleted Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Delete dbo.vACCPACExport_Cust_Short View'
END
GO

--
-- Script To Delete dbo.vACCPACExport_CUST_Short2 View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Deleting dbo.vACCPACExport_CUST_Short2 View'
GO

   DROP VIEW [dbo].[vACCPACExport_CUST_Short2]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vACCPACExport_CUST_Short2 View Deleted Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Delete dbo.vACCPACExport_CUST_Short2 View'
END
GO

--
-- Script To Create dbo.vBudget_Pivot1 View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.vBudget_Pivot1 View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


SET QUOTED_IDENTIFIER OFF
GO

SET ANSI_NULLS OFF
GO

exec('CREATE VIEW [dbo].[vBudget_Pivot1]
AS
SELECT     TOP (100) PERCENT DATEADD(m, 5, dbo.StartOfMonth(GETDATE())) AS h12, DATEADD(m, 4, dbo.StartOfMonth(GETDATE())) AS h11, DATEADD(m, 3, 
                      dbo.StartOfMonth(GETDATE())) AS h10, DATEADD(m, 2, dbo.StartOfMonth(GETDATE())) AS h09, DATEADD(m, 1, dbo.StartOfMonth(GETDATE())) AS h08, 
                      DATEADD(m, 0, dbo.StartOfMonth(GETDATE())) AS h07, DATEADD(m, - 1, dbo.StartOfMonth(GETDATE())) AS h06, DATEADD(m, - 2, 
                      dbo.StartOfMonth(GETDATE())) AS h05, DATEADD(m, - 3, dbo.StartOfMonth(GETDATE())) AS h04, DATEADD(m, - 4, dbo.StartOfMonth(GETDATE())) AS h03, 
                      DATEADD(m, - 5, dbo.StartOfMonth(GETDATE())) AS h02, DATEADD(m, - 6, dbo.StartOfMonth(GETDATE())) AS h01, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesBudget, 0) ELSE 0 END) 
                      AS m12_1, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesBudget, 0) 
                      ELSE 0 END) AS m11_1, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesBudget, 
                      0) ELSE 0 END) AS m10_1, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 2, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_DeliveriesBudget, 0) ELSE 0 END) AS m09_1, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 1, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesBudget, 0) ELSE 0 END) AS m08_1, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, 0, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesBudget, 0) ELSE 0 END) AS m07_1, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 1,
                       dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesBudget, 0) ELSE 0 END) AS m06_1, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, 
                      - 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesBudget, 0) ELSE 0 END) AS m05_1, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m,
                       - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesBudget, 0) ELSE 0 END) AS m04_1, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesBudget, 0) ELSE 0 END) AS m03_1,
                       SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesBudget, 0) ELSE 0 END) 
                      AS m02_1, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesBudget, 0) ELSE 0 END)
                       AS m01_1, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_ReturnsBudget, 0) ELSE 0 END) 
                      AS m12_2, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_ReturnsBudget, 0) ELSE 0 END) 
                      AS m11_2, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_ReturnsBudget, 0) ELSE 0 END) 
                      AS m10_2, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_ReturnsBudget, 0) ELSE 0 END) 
                      AS m09_2, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_ReturnsBudget, 0) ELSE 0 END) 
                      AS m08_2, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_ReturnsBudget, 0) ELSE 0 END) 
                      AS m07_2, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_ReturnsBudget, 0) ELSE 0 END) 
                      AS m06_2, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_ReturnsBudget, 0) ELSE 0 END) 
                      AS m05_2, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_ReturnsBudget, 0) ELSE 0 END) 
                      AS m04_2, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_ReturnsBudget, 0) ELSE 0 END) 
                      AS m03_2, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_ReturnsBudget, 0) ELSE 0 END) 
                      AS m02_2, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_ReturnsBudget, 0) ELSE 0 END) 
                      AS m01_2, 

SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReceivedIssued, 0) ELSE 0 END) 
                      AS m12_3, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReceivedIssued, 0) ELSE 0 END) 
                      AS m11_3, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReceivedIssued, 0) ELSE 0 END) 
                      AS m10_3, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReceivedIssued, 0) ELSE 0 END) 
                      AS m09_3, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReceivedIssued, 0) ELSE 0 END) 
                      AS m08_3, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReceivedIssued, 0) ELSE 0 END) 
                      AS m07_3, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReceivedIssued, 0) ELSE 0 END) 
                      AS m06_3, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReceivedIssued, 0) ELSE 0 END) 
                      AS m05_3, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReceivedIssued, 0) ELSE 0 END) 
                      AS m04_3, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReceivedIssued, 0) ELSE 0 END) 
                      AS m03_3, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReceivedIssued, 0) ELSE 0 END) 
                      AS m02_3, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReceivedIssued, 0) ELSE 0 END) 
                      AS m01_3, 

SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReturnsIssued, 0) ELSE 0 END) 
                      AS m12_4, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReturnsIssued, 0) ELSE 0 END) 
                      AS m11_4, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReturnsIssued, 0) ELSE 0 END) 
                      AS m10_4, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReturnsIssued, 0) ELSE 0 END) 
                      AS m09_4, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReturnsIssued, 0) ELSE 0 END) 
                      AS m08_4, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReturnsIssued, 0) ELSE 0 END) 
                      AS m07_4, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReturnsIssued, 0) ELSE 0 END) 
                      AS m06_4, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReturnsIssued, 0) ELSE 0 END) 
                      AS m05_4, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReturnsIssued, 0) ELSE 0 END) 
                      AS m04_4, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReturnsIssued, 0) ELSE 0 END) 
                      AS m03_4, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReturnsIssued, 0) ELSE 0 END) 
                      AS m02_4, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_RetailValueReturnsIssued, 0) ELSE 0 END) 
                      AS m01_4, 




SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAtRetailValueIssued, 0) 
                      ELSE 0 END) AS m12_5, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 4, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_OrdersAtRetailValueIssued, 0) ELSE 0 END) AS m11_5, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 3, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAtRetailValueIssued, 0) ELSE 0 END) AS m10_5, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAtRetailValueIssued, 0) ELSE 0 END) 
                      AS m09_5, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAtRetailValueIssued, 0) 
                      ELSE 0 END) AS m08_5, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 0, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_OrdersAtRetailValueIssued, 0) ELSE 0 END) AS m07_5, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 1, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAtRetailValueIssued, 0) ELSE 0 END) AS m06_5, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAtRetailValueIssued, 0) ELSE 0 END) 
                      AS m05_5, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAtRetailValueIssued, 0) 
                      ELSE 0 END) AS m04_5, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_OrdersAtRetailValueIssued, 0) ELSE 0 END) AS m03_5, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 5, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAtRetailValueIssued, 0) ELSE 0 END) AS m02_5, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAtRetailValueIssued, 0) ELSE 0 END) 
                      AS m01_5, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAtRetailValueInProcess, 0) 
                      ELSE 0 END) AS m12_6, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 4, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_OrdersAtRetailValueInProcess, 0) ELSE 0 END) AS m11_6, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 3, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAtRetailValueInProcess, 0) ELSE 0 END) AS m10_6, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAtRetailValueInProcess, 0) 
                      ELSE 0 END) AS m09_6, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 1, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_OrdersAtRetailValueInProcess, 0) ELSE 0 END) AS m08_6, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 0, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAtRetailValueInProcess, 0) ELSE 0 END) AS m07_6, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAtRetailValueInProcess, 0) 
                      ELSE 0 END) AS m06_6, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_OrdersAtRetailValueInProcess, 0) ELSE 0 END) AS m05_6, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 3, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAtRetailValueInProcess, 0) ELSE 0 END) AS m04_6, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAtRetailValueInProcess, 0) 
                      ELSE 0 END) AS m03_6, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_OrdersAtRetailValueInProcess, 0) ELSE 0 END) AS m02_6, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 6, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAtRetailValueInProcess, 0) ELSE 0 END) AS m01_6, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAgainstBudget, 0) ELSE 0 END) 
                      AS m12_7, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAgainstBudget, 0) 
                      ELSE 0 END) AS m11_7, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 3, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_OrdersAgainstBudget, 0) ELSE 0 END) AS m10_7, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 2, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAgainstBudget, 0) ELSE 0 END) AS m09_7, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m,
                       + 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAgainstBudget, 0) ELSE 0 END) AS m08_7, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 0, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAgainstBudget, 0) ELSE 0 END) 
                      AS m07_7, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAgainstBudget, 0) 
                      ELSE 0 END) AS m06_7, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_OrdersAgainstBudget, 0) ELSE 0 END) AS m05_7, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 3, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAgainstBudget, 0) ELSE 0 END) AS m04_7, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m,
                       - 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAgainstBudget, 0) ELSE 0 END) AS m03_7, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAgainstBudget, 0) ELSE 0 END) 
                      AS m02_7, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_OrdersAgainstBudget, 0) 
                      ELSE 0 END) AS m01_7, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 5, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_DeliveriesAgainstBudget, 0) ELSE 0 END) AS m12_8, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 4, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesAgainstBudget, 0) ELSE 0 END) AS m11_8, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 3, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesAgainstBudget, 0) ELSE 0 END) 
                      AS m10_8, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesAgainstBudget, 0) 
                      ELSE 0 END) AS m09_8, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 1, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_DeliveriesAgainstBudget, 0) ELSE 0 END) AS m08_8, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 0, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesAgainstBudget, 0) ELSE 0 END) AS m07_8, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesAgainstBudget, 0) ELSE 0 END) 
                      AS m06_8, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesAgainstBudget, 0) 
                      ELSE 0 END) AS m05_8, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 3, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_DeliveriesAgainstBudget, 0) ELSE 0 END) AS m04_8, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 4, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesAgainstBudget, 0) ELSE 0 END) AS m03_8, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesAgainstBudget, 0) ELSE 0 END) 
                      AS m02_8, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 6, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesAgainstBudget, 0) 
                      ELSE 0 END) AS m01_8, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 5, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_DeliveriesAgainstBudget_FourMonthAverage, 0) ELSE 0 END) AS m12_9, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 4, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesAgainstBudget_FourMonthAverage, 0) ELSE 0 END) AS m11_9, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 3, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_DeliveriesAgainstBudget_FourMonthAverage, 0) ELSE 0 END) AS m10_9, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 2, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesAgainstBudget_FourMonthAverage, 0) ELSE 0 END) AS m09_9, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 1, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_DeliveriesAgainstBudget_FourMonthAverage, 0) ELSE 0 END) AS m08_9, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, + 0, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesAgainstBudget_FourMonthAverage, 0) ELSE 0 END) AS m07_9, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 1, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesAgainstBudget_FourMonthAverage,
                       0) ELSE 0 END) AS m06_9, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 2, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_DeliveriesAgainstBudget_FourMonthAverage, 0) ELSE 0 END) AS m05_9, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 3, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesAgainstBudget_FourMonthAverage, 0) ELSE 0 END) AS m04_9, 
                      SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 4, dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesAgainstBudget_FourMonthAverage,
                       0) ELSE 0 END) AS m03_9, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 5, dbo.StartOfMonth(GetDate())) 
                      THEN ISNULL(X.B_DeliveriesAgainstBudget_FourMonthAverage, 0) ELSE 0 END) AS m02_9, SUM(CASE X.B_BUDGETMONTH WHEN DATEADD(m, - 6, 
                      dbo.StartOfMonth(GetDate())) THEN ISNULL(X.B_DeliveriesAgainstBudget_FourMonthAverage, 0) ELSE 0 END) AS m01_9
FROM         dbo.tBudget AS x')
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vBudget_Pivot1 View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.vBudget_Pivot1 View'
END
GO

--
-- Script To Create dbo.vExportDebtorsCreditNotesForAccounting_Pastel_Default View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.vExportDebtorsCreditNotesForAccounting_Pastel_Default View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.vExportDebtorsCreditNotesForAccounting_Pastel_Default
AS
SELECT     TOP (100) PERCENT dbo.tTR.TR_ID, SUM(CAST(dbo.CalcExt(dbo.tCNL.CNL_QTY, dbo.tCNL.CNL_Price, dbo.tCNL.CNL_DiscountRate) AS NUMERIC(18, 
                      2)) / dbo.tCurrency.CURR_Divisor) AS AMT, SUM(CAST(dbo.CalcExt(dbo.tCNL.CNL_QTY, dbo.tCNL.CNL_Price, dbo.tCNL.CNL_DiscountRate) 
                      AS NUMERIC(18, 2)) / dbo.tCurrency.CURR_Divisor - CAST(dbo.CalcExtEXVATb(dbo.tCNL.CNL_QTY, dbo.tCNL.CNL_Price, dbo.tCNL.CNL_DiscountRate, 
                      dbo.tCNL.CNL_VATRate, dbo.tTP.TP_VATable, dbo.tTP.TP_ShowVAT) AS NUMERIC(18, 2)) / dbo.tCurrency.CURR_Divisor) AS VAT, dbo.tTR.TR_Code, 
                      SUM(dbo.CalcExt2(dbo.tCNL.CNL_QTY, dbo.tProduct.P_Cost, 0, dbo.tCurrency.CURR_Divisor)) AS AVGCOST, dbo.tTP.TP_VATable, 
                      dbo.tTR.TR_ProcessingDate
FROM         dbo.tConfiguration INNER JOIN
                      dbo.tCurrency ON dbo.tConfiguration.CF_DefaultCurrID = dbo.tCurrency.CURR_ID CROSS JOIN
                      dbo.tProduct RIGHT OUTER JOIN
                      dbo.tTR INNER JOIN
                      dbo.tTP ON dbo.tTR.TR_TP_ID = dbo.tTP.TP_ID INNER JOIN
                      dbo.tCNL ON dbo.tTR.TR_ID = dbo.tCNL.CNL_TR_ID ON dbo.tProduct.P_ID = dbo.tCNL.CNL_P_ID LEFT OUTER JOIN
                      dbo.tPT ON dbo.tProduct.P_ProductType_ID = dbo.tPT.PT_ID
WHERE     (dbo.tTR.TR_Status IN (3, 4)) AND (dbo.tCNL.CNL_QTY > 0)
GROUP BY dbo.tTR.TR_ID, dbo.tTR.TR_Code, dbo.tTP.TP_VATable, dbo.tTR.TR_ProcessingDate')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vExportDebtorsCreditNotesForAccounting_Pastel_Default View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.vExportDebtorsCreditNotesForAccounting_Pastel_Default View'
END
GO

--
-- Script To Create dbo.vExportDebtorsCreditNotesForAccounting_Pastel_Special View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.vExportDebtorsCreditNotesForAccounting_Pastel_Special View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.vExportDebtorsCreditNotesForAccounting_Pastel_Special
AS
SELECT     TOP (100) PERCENT dbo.tTR.TR_ID, SUM(CAST(dbo.CalcExt(dbo.tCNL.CNL_QTY, dbo.tCNL.CNL_Price, dbo.tCNL.CNL_DiscountRate) AS NUMERIC(18, 
                      2)) / dbo.tCurrency.CURR_Divisor) AS AMT, SUM(CAST(dbo.CalcExt(dbo.tCNL.CNL_QTY, dbo.tCNL.CNL_Price, dbo.tCNL.CNL_DiscountRate) 
                      AS NUMERIC(18, 2)) / dbo.tCurrency.CURR_Divisor - CAST(dbo.CalcExtEXVATb(dbo.tCNL.CNL_QTY, dbo.tCNL.CNL_Price, dbo.tCNL.CNL_DiscountRate, 
                      dbo.tCNL.CNL_VATRate, dbo.tTP.TP_VATable, dbo.tTP.TP_ShowVAT) AS NUMERIC(18, 2)) / dbo.tCurrency.CURR_Divisor) AS VAT, dbo.tPT.PT_CRSALES, 
                      dbo.tPT.PT_CRSALES_CONTRA, dbo.tTR.TR_Code, SUM(CAST(dbo.CalcExtEXVATb(dbo.tCNL.CNL_QTY, dbo.tProduct.P_Cost, 0, 
                      ISNULL(dbo.tProduct.P_VATRate, dbo.tConfiguration.CF_VATRATE), dbo.tTP.TP_VATable, dbo.tTP.TP_ShowVAT) AS NUMERIC(18, 2)) 
                      / dbo.tCurrency.CURR_Divisor) AS AVGCOST, dbo.tProduct.P_ProductType, dbo.tTP.TP_VATable, dbo.tTR.TR_ProcessingDate
FROM         dbo.tConfiguration INNER JOIN
                      dbo.tCurrency ON dbo.tConfiguration.CF_DefaultCurrID = dbo.tCurrency.CURR_ID CROSS JOIN
                      dbo.tProduct RIGHT OUTER JOIN
                      dbo.tTR INNER JOIN
                      dbo.tTP ON dbo.tTR.TR_TP_ID = dbo.tTP.TP_ID INNER JOIN
                      dbo.tCNL ON dbo.tTR.TR_ID = dbo.tCNL.CNL_TR_ID ON dbo.tProduct.P_ID = dbo.tCNL.CNL_P_ID LEFT OUTER JOIN
                      dbo.tPT ON dbo.tProduct.P_ProductType_ID = dbo.tPT.PT_ID
WHERE     (dbo.tTR.TR_Status IN (3, 4))
GROUP BY dbo.tTR.TR_ID, dbo.tPT.PT_CRSALES, dbo.tPT.PT_CRSALES_CONTRA, dbo.tTR.TR_Code, dbo.tProduct.P_ProductType, dbo.tTP.TP_VATable, 
                      dbo.tTR.TR_ProcessingDate
HAVING      (dbo.tPT.PT_CRSALES > '''') AND (dbo.tPT.PT_CRSALES_CONTRA > '''')')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vExportDebtorsCreditNotesForAccounting_Pastel_Special View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.vExportDebtorsCreditNotesForAccounting_Pastel_Special View'
END
GO

--
-- Script To Create dbo.vExportDebtorsInvoicesForAccounting_Pastel_Default View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.vExportDebtorsInvoicesForAccounting_Pastel_Default View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.vExportDebtorsInvoicesForAccounting_Pastel_Default
AS
SELECT     TOP (100) PERCENT dbo.tTR.TR_ID, SUM(CAST(dbo.CalcExt(dbo.tILine.IL_Qty, dbo.tILine.IL_Price, dbo.tILine.IL_DiscountRate) AS NUMERIC(18, 0)) 
                      / dbo.tCurrency.CURR_Divisor) AS AMT, SUM(CAST(dbo.CalcExt(dbo.tILine.IL_Qty, dbo.tILine.IL_Price, dbo.tILine.IL_DiscountRate) AS NUMERIC(18, 0)) 
                      / dbo.tCurrency.CURR_Divisor - CAST(dbo.CalcExtEXVATb(dbo.tILine.IL_Qty, dbo.tILine.IL_Price, dbo.tILine.IL_DiscountRate, dbo.tILine.IL_VATRate, 
                      dbo.tTP.TP_VATable, dbo.tInvoice.I_ShowVAT) AS NUMERIC(18, 0)) / dbo.tCurrency.CURR_Divisor) AS VAT, SUM(dbo.CalcExt2(dbo.tILine.IL_Qty, 
                      dbo.tILine.IL_AvgCost, 0, dbo.tCurrency.CURR_Divisor)) AS AVGCOST, dbo.tTR.TR_Code, dbo.tTP.TP_VATable, dbo.tTR.TR_ProcessingDate, 
                      CASE WHEN ISNULL(TP_VATable, 1) = 1 THEN 1 ELSE 3 END AS TaxType
FROM         dbo.tConfiguration INNER JOIN
                      dbo.tCurrency ON dbo.tConfiguration.CF_DefaultCurrID = dbo.tCurrency.CURR_ID CROSS JOIN
                      dbo.tPT RIGHT OUTER JOIN
                      dbo.tProduct INNER JOIN
                      dbo.tTR INNER JOIN
                      dbo.tTP ON dbo.tTR.TR_TP_ID = dbo.tTP.TP_ID INNER JOIN
                      dbo.tInvoice ON dbo.tTR.TR_ID = dbo.tInvoice.I_ID INNER JOIN
                      dbo.tILine ON dbo.tTR.TR_ID = dbo.tILine.IL_TR_ID ON dbo.tProduct.P_ID = dbo.tILine.IL_P_ID ON 
                      dbo.tPT.PT_ID = dbo.tProduct.P_ProductType_ID
WHERE     (dbo.tTR.TR_Status = 4) AND (ISNULL(dbo.tInvoice.I_ProForma, 0) <> 1) AND (ISNULL(dbo.tILine.IL_Qty, 0) > 0) OR
                      (dbo.tTR.TR_Status = 4) AND (ISNULL(dbo.tInvoice.I_ProForma, 0) <> 1) AND (ISNULL(dbo.tILine.IL_Qty, 0) > 0)
GROUP BY dbo.tTR.TR_ID, dbo.tTR.TR_Code, dbo.tTP.TP_VATable, dbo.tTR.TR_ProcessingDate')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vExportDebtorsInvoicesForAccounting_Pastel_Default View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.vExportDebtorsInvoicesForAccounting_Pastel_Default View'
END
GO

--
-- Script To Create dbo.vExportDebtorsInvoicesForAccounting_Pastel_Special View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.vExportDebtorsInvoicesForAccounting_Pastel_Special View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.vExportDebtorsInvoicesForAccounting_Pastel_Special
AS
SELECT     TOP (100) PERCENT dbo.tTR.TR_ID, SUM(CAST(dbo.CalcExt(dbo.tILine.IL_Qty, dbo.tILine.IL_Price, dbo.tILine.IL_DiscountRate) AS NUMERIC(18, 0)) 
                      / dbo.tCurrency.CURR_Divisor) AS AMT, SUM(CAST(dbo.CalcExt(dbo.tILine.IL_Qty, dbo.tILine.IL_Price, dbo.tILine.IL_DiscountRate) AS NUMERIC(18, 0)) 
                      / dbo.tCurrency.CURR_Divisor - CAST(dbo.CalcExtEXVATb(dbo.tILine.IL_Qty, dbo.tILine.IL_Price, dbo.tILine.IL_DiscountRate, dbo.tILine.IL_VATRate, 
                      dbo.tTP.TP_VATable, dbo.tInvoice.I_ShowVAT) AS NUMERIC(18, 0)) / dbo.tCurrency.CURR_Divisor) AS VAT, dbo.tPT.PT_CRSALES, 
                      dbo.tPT.PT_CRSALES_CONTRA, SUM(dbo.CalcExt2(dbo.tILine.IL_Qty, dbo.tILine.IL_AvgCost, 0, dbo.tCurrency.CURR_Divisor)) AS AVGCOST, 
                      dbo.tTR.TR_Code, dbo.tProduct.P_ProductType, dbo.tTP.TP_VATable, dbo.tTR.TR_ProcessingDate
FROM         dbo.tConfiguration INNER JOIN
                      dbo.tCurrency ON dbo.tConfiguration.CF_DefaultCurrID = dbo.tCurrency.CURR_ID CROSS JOIN
                      dbo.tPT RIGHT OUTER JOIN
                      dbo.tProduct INNER JOIN
                      dbo.tTR INNER JOIN
                      dbo.tTP ON dbo.tTR.TR_TP_ID = dbo.tTP.TP_ID INNER JOIN
                      dbo.tInvoice ON dbo.tTR.TR_ID = dbo.tInvoice.I_ID INNER JOIN
                      dbo.tILine ON dbo.tTR.TR_ID = dbo.tILine.IL_TR_ID ON dbo.tProduct.P_ID = dbo.tILine.IL_P_ID ON 
                      dbo.tPT.PT_ID = dbo.tProduct.P_ProductType_ID
WHERE     (dbo.tTR.TR_Status = 4) AND (ISNULL(dbo.tInvoice.I_ProForma, 0) <> 1) AND (ISNULL(dbo.tILine.IL_Qty, 0) > 0)
GROUP BY dbo.tTR.TR_ID, dbo.tPT.PT_CRSALES, dbo.tPT.PT_CRSALES_CONTRA, dbo.tTR.TR_Code, dbo.tProduct.P_ProductType, dbo.tTP.TP_VATable, 
                      dbo.tTR.TR_ProcessingDate
HAVING      (dbo.tPT.PT_CRSALES > '''') AND (dbo.tPT.PT_CRSALES_CONTRA > '''')')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vExportDebtorsInvoicesForAccounting_Pastel_Special View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.vExportDebtorsInvoicesForAccounting_Pastel_Special View'
END
GO

--
-- Script To Create dbo.vPastelExport2_TaxInvoices View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.vPastelExport2_TaxInvoices View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.vPastelExport2_TaxInvoices
AS
SELECT     TOP 100 PERCENT PERIOD AS Perd, ''"'' + CONVERT(CHAR(10), DTE, 103) + ''"'' AS Dat, ''"'' + LTRIM(LEFT(GDC, 1)) + ''"'' AS GDC, ''"'' + LTRIM(LEFT(ACNO, 
                      7)) + ''"'' AS Acno, ''"'' + LTRIM(RIGHT(REFERENCE, 8)) + ''"'' AS Ref, ''"'' + LTRIM(RIGHT(DESCR, 36)) + ''"'' AS Descr, AMT, TAXTYPE, TAXAMT, 
                      ''"'' + LTRIM(LEFT(OPENITEM, 1)) + ''"'' AS Openitem, ''"'' + LTRIM(LEFT(COSTCODE, 5)) + ''"'' AS Costcode, ''"'' + LTRIM(LEFT(CONTRAACCOUNT, 7)) 
                      + ''"'' AS Contraaccount, EXCHANGERATE, BANKEXCHANGERATE, BATCHID, DISCOUNTTAX, DISCOUNTAMT, HOMEAMT
FROM         dbo.tPASTEL
WHERE     (DESCR = ''TAX_INVOICE'') AND (ACTION = 1)
ORDER BY Dat, Ref')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vPastelExport2_TaxInvoices View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.vPastelExport2_TaxInvoices View'
END
GO

--
-- Script To Create dbo.vPastelExport_Sales View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.vPastelExport_Sales View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.vPastelExport_Sales
AS
SELECT     TOP 100 PERCENT dbo.tTR.TR_ID, dbo.CalcExt2(dbo.tILine.IL_Qty, dbo.tILine.IL_Price, dbo.tILine.IL_DiscountRate, dbo.tCurrency.CURR_Divisor) 
                      AS AMT, dbo.CalcExtVAT2(dbo.tILine.IL_Qty, dbo.tILine.IL_Price, dbo.tILine.IL_DiscountRate, dbo.tILine.IL_VATRate, dbo.tCurrency.CURR_Divisor) 
                      AS VAT
FROM         dbo.tTR INNER JOIN
                      dbo.tTP ON dbo.tTR.TR_TP_ID = dbo.tTP.TP_ID INNER JOIN
                      dbo.tInvoice ON dbo.tTR.TR_ID = dbo.tInvoice.I_ID INNER JOIN
                      dbo.tILine ON dbo.tTR.TR_ID = dbo.tILine.IL_TR_ID CROSS JOIN
                      dbo.tConfiguration INNER JOIN
                      dbo.tCurrency ON dbo.tConfiguration.CF_DefaultCurrID = dbo.tCurrency.CURR_ID
WHERE     (dbo.tInvoice.I_ProForma <> 1) AND (dbo.tTR.TR_Status = 3 OR
                      dbo.tTR.TR_Status = 4)
GROUP BY dbo.tTR.TR_ID, dbo.CalcExt2(dbo.tILine.IL_Qty, dbo.tILine.IL_Price, dbo.tILine.IL_DiscountRate, dbo.tCurrency.CURR_Divisor), 
                      dbo.CalcExtVAT2(dbo.tILine.IL_Qty, dbo.tILine.IL_Price, dbo.tILine.IL_DiscountRate, dbo.tILine.IL_VATRate, dbo.tCurrency.CURR_Divisor)')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vPastelExport_Sales View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.vPastelExport_Sales View'
END
GO

--
-- Script To Create dbo.vPastelExport_SInvoice View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.vPastelExport_SInvoice View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.vPastelExport_SInvoice
AS
SELECT     TOP 100 PERCENT dbo.tTR.TR_ID, SUM(dbo.CalcExt2(dbo.tDELL.DELL_QtyTotal, dbo.tDELL.DELL_Price, dbo.tDELL.DELL_Discount, 
                      dbo.tCurrency.CURR_Divisor)) AS AMT, SUM(dbo.CalcExtVAT2(dbo.tDELL.DELL_QtyTotal, dbo.tDELL.DELL_Price, dbo.tDELL.DELL_Discount, 
                      dbo.tProduct.P_VATRate, dbo.tCurrency.CURR_Divisor)) AS VAT
FROM         dbo.tProduct INNER JOIN
                      dbo.tTR INNER JOIN
                      dbo.tDEL ON dbo.tTR.TR_ID = dbo.tDEL.DEL_ID INNER JOIN
                      dbo.tDELL ON dbo.tTR.TR_ID = dbo.tDELL.DELL_TR_ID ON dbo.tProduct.P_ID = dbo.tDELL.DELL_P_ID CROSS JOIN
                      dbo.tConfiguration INNER JOIN
                      dbo.tCurrency ON dbo.tConfiguration.CF_LocalCurrID = dbo.tCurrency.CURR_ID
WHERE     (dbo.tTR.TR_Status = 3) OR
                      (dbo.tTR.TR_Status = 4)
GROUP BY dbo.tTR.TR_ID
ORDER BY dbo.tTR.TR_ID DESC')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vPastelExport_SInvoice View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.vPastelExport_SInvoice View'
END
GO

--
-- Script To Create dbo.vPASTELExport_Supp View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.vPASTELExport_Supp View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.vPASTELExport_Supp
AS
SELECT     TOP 100 PERCENT ''"'' + dbo.STRIPDOUBLEQUOTES(PC_ACNO) + ''"'' AS Expr2, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_Name) + ''"'' AS Expr3, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_ADD1) + ''"'' AS Expr4, 
                      ''"'' + dbo.STRIPDOUBLEQUOTES(PC_ADD2) + ''"'' AS Expr5, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_ADD3) + ''"'' AS Expr6, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_ADD4) + ''"'' AS Expr7, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_ADD5) + ''"'' AS Expr8, 
                      ''"'' + dbo.STRIPDOUBLEQUOTES(PC_Tel) + ''"'' AS Expr9, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_Fax) + ''"'' AS Expr10, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_Contact) + ''"'' AS Expr11, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_DEFAULTTAX) 
                      + ''"'' AS Expr12, ''""'' AS EarlyPaymentTerms, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_DADD1) + ''"'' AS Expr1, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_DADD2) + ''"'' AS Expr13, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_DADD3) 
                      + ''"'' AS Expr14, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_DADD4) + ''"'' AS Expr15, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_DADD5) + ''"'' AS Expr16, ''"'' + CASE dbo.STRIPDOUBLEQUOTES(PC_BLOCKED) 
                      WHEN 1 THEN ''Y'' WHEN 0 THEN ''N'' END + ''"'' AS Expr17, ''""'' AS Exclusive, ''""'' AS StatementsMessage, ''"N"'' AS OpenItem, 
                      ''"'' + dbo.STRIPDOUBLEQUOTES(PC_CurrencyCode) + ''"'' AS e, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_PaymentTerms) + ''"'' AS f, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_CreditLimit) + ''"'' AS g, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_MobilePhone) 
                      + ''"'' AS h, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_Email) + ''"'' AS i, ''""'' AS j, ''""'' AS k, ''""'' AS l, ''"'' + dbo.STRIPDOUBLEQUOTES(PC_CountryCode) + ''"'' AS m, ''""'' AS Expr19, ''""'' AS n, ''""'' AS Expr20, 
                      ''""'' AS Expr21, ''""'' AS q, ''""'' AS r, ''""'' AS s, ''""'' AS t, ''""'' AS u, ''""'' AS taxcode, ''""'' AS defaultcontra
FROM         dbo.tPASTEL_SUPP_EXPORT
WHERE     (PC_Action = 1)
ORDER BY PC_ACNO')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vPASTELExport_Supp View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.vPASTELExport_Supp View'
END
GO

--
-- Script To Create dbo.vPASTELExport_TAXCREDITNOTES View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.vPASTELExport_TAXCREDITNOTES View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.vPASTELExport_TAXCREDITNOTES
AS
SELECT     TOP 100 PERCENT PERIOD, ''"'' + CONVERT(CHAR(10), DTE, 103) + ''"'' AS DTE1, ''"'' + LTRIM(LEFT(GDC, 1)) + ''"'' AS GDC, ''"'' + LTRIM(LEFT(ACNO, 7)) 
                      + ''"'' AS ACNO, ''"'' + LTRIM(RIGHT(REFERENCE, 8)) + ''"'' AS REF, ''"'' + LTRIM(LEFT(DESCR, 36)) + ''"'' AS DESCR, AMT * - 1 AS AMT, TAXTYPE, 
                      TAXAMT * - 1 AS TAXAMT, ''"'' + LTRIM(LEFT(OPENITEM, 1)) + ''"'' AS OPENITEM, ''"'' + LTRIM(LEFT(COSTCODE, 5)) + ''"'' AS COSTCODE, 
                      ''"'' + LTRIM(LEFT(CONTRAACCOUNT, 7)) + ''"'' AS CONTRAACCOUNT, EXCHANGERATE, BANKEXCHANGERATE, BATCHID, DISCOUNTTAX, 
                      DISCOUNTAMT * - 1 AS DISCOUNTAMT, HOMEAMT * - 1 AS HOMEAMT
FROM         dbo.tPASTEL
WHERE     (DESCR = ''TAX_CREDITNOTE'') AND (ACTION = 1)
ORDER BY DTE1, REF')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vPASTELExport_TAXCREDITNOTES View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.vPASTELExport_TAXCREDITNOTES View'
END
GO

--
-- Script To Create dbo.vPastelExport_TAXINVOICES View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.vPastelExport_TAXINVOICES View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.vPastelExport_TAXINVOICES
AS
SELECT     TOP 100 PERCENT PERIOD, ''"'' + CONVERT(CHAR(10), DTE, 103) + ''"'' AS DTE1, ''"'' + LTRIM(LEFT(GDC, 1)) + ''"'' AS GDC, ''"'' + LTRIM(LEFT(ACNO, 7)) 
                      + ''"'' AS ACNO, ''"'' + LTRIM(RIGHT(REFERENCE, 8)) + ''"'' AS REF, ''"'' + LTRIM(RIGHT(DESCR, 36)) + ''"'' AS DESCR, AMT, TAXTYPE, TAXAMT, 
                      ''"'' + LTRIM(LEFT(OPENITEM, 1)) + ''"'' AS OPENITEM, ''"'' + LTRIM(LEFT(COSTCODE, 5)) + ''"'' AS COSTCODE, ''"'' + LTRIM(LEFT(CONTRAACCOUNT, 7)) 
                      + ''"'' AS CONTRAACCOUNT, EXCHANGERATE AS Expr1, BANKEXCHANGERATE, BATCHID, DISCOUNTTAX, DISCOUNTAMT, HOMEAMT
FROM         dbo.tPASTEL
WHERE     (DESCR = ''TAX_INVOICE'') AND (ACTION = 1)
ORDER BY DTE1, REF')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vPastelExport_TAXINVOICES View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.vPastelExport_TAXINVOICES View'
END
GO

--
-- Script To Create dbo.vPOLSInProcessPerPID View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.vPOLSInProcessPerPID View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE VIEW dbo.vPOLSInProcessPerPID
AS
SELECT     dbo.tPOL.POL_P_ID AS PID, SUM(dbo.NonNegative(ISNULL(dbo.tPOL.POL_QtySS, 0) + ISNULL(dbo.tPOL.POL_QtyFirm, 0))) AS OO, 
                      dbo.tPOL.POL_TR_ID AS POID
FROM         dbo.tPOL INNER JOIN
                      dbo.tTR ON dbo.tPOL.POL_TR_ID = dbo.tTR.TR_ID
WHERE     (dbo.tTR.TR_Status = 2) AND (ISNULL(dbo.tPOL.POL_DateReplaced, 0) = 0)
GROUP BY dbo.tPOL.POL_P_ID, dbo.tPOL.POL_TR_ID')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vPOLSInProcessPerPID View Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.vPOLSInProcessPerPID View'
END
GO

--
-- Script To Update dbo.vReorder_Browsed View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.vReorder_Browsed View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER VIEW dbo.vReorder_Browsed
AS
SELECT     dbo.tProduct.P_ID, dbo.ProductStatusF(dbo.tProduct.P_Status, dbo.tProduct.P_Obsolete) + dbo.tProduct.P_Title AS Description, dbo.tProduct.P_EAN, 
                      dbo.tProduct.P_MainAuthor, dbo.tProduct.P_Publisher, dbo.tProduct.P_LastQtyFirmOrdered, dbo.tProduct.P_LastQtySSOrdered, 
                      dbo.tProduct.P_LastQtyDelivered, dbo.tProduct.P_LastPriceDelivered, dbo.tProduct.P_QtyOnHand, dbo.tProduct.P_QtyOnOrder, 
                      dbo.tProduct.P_QtyOnBackorder, dbo.tProduct.P_LastDateDelivered, dbo.tProduct.P_LastDateOrdered, dbo.tProduct.P_LastPriceOrdered, 
                      dbo.tDeal.DL_Description, dbo.tTP.TP_Name AS SUPPLIERNAME, dbo.tProduct.P_SupplierID, dbo.tProduct.P_DealID, dbo.tProduct.P_Code, 
                      dbo.tProduct.P_RRP, dbo.tProduct.P_SP, dbo.tPT.PT_Code AS PT, dbo.tProduct.P_QtyOnAppro, dbo.FormatSales(ISNULL(dbo.vWeekSalesPivot.w1, 0), 
                      ISNULL(dbo.vWeekSalesPivot.w2, 0), ISNULL(dbo.vWeekSalesPivot.w3, 0), ISNULL(dbo.vWeekSalesPivot.w4, 0), ISNULL(dbo.vWeekSalesPivot.w5, 0), 
                      ISNULL(dbo.vWeekSalesPivot.w6, 0), ''w'') AS LASTSIXWEEKS, dbo.FormatSales(ISNULL(dbo.vMonthSalesPivot.m1, 0), 
                      ISNULL(dbo.vMonthSalesPivot.m2, 0), ISNULL(dbo.vMonthSalesPivot.m3, 0), ISNULL(dbo.vMonthSalesPivot.m4, 0), ISNULL(dbo.vMonthSalesPivot.m5, 
                      0), ISNULL(dbo.vMonthSalesPivot.m6, 0), ''m'') AS LASTSIXMONTHS, dbo.tProduct.P_Title AS TitleForSort, 0 AS Qty, 
                      dbo.GetForeignPrice(dbo.tProduct.P_EUPrice, dbo.tProduct.P_UKPrice, dbo.tProduct.P_USPrice, dbo.tCurrency.CURR_SYSTEM) AS ForeignPrice, 
                      dbo.FlattenSections(dbo.tProduct.P_ID) AS Sections, dbo.GetWorkingDeal(dbo.tProduct.P_DealID, dbo.tTP.TP_ID) AS DefaultDeal, 
                      dbo.tProduct.P_QtyOnOrder_UnIssued
FROM         dbo.tCurrency RIGHT OUTER JOIN
                      dbo.tTP ON dbo.tCurrency.CURR_ID = dbo.tTP.TP_CURR_ID RIGHT OUTER JOIN
                      dbo.tPT INNER JOIN
                      dbo.tProduct ON dbo.tPT.PT_ID = dbo.tProduct.P_ProductType_ID INNER JOIN
                      dbo.tmpBrowsedProducts ON dbo.tProduct.P_ID = dbo.tmpBrowsedProducts.PID LEFT OUTER JOIN
                      dbo.vMonthSalesPivot ON dbo.tProduct.P_ID = dbo.vMonthSalesPivot.PID LEFT OUTER JOIN
                      dbo.vWeekSalesPivot ON dbo.tProduct.P_ID = dbo.vWeekSalesPivot.PID LEFT OUTER JOIN
                      dbo.tDeal ON dbo.tProduct.P_DealID = dbo.tDeal.DL_ID ON dbo.tTP.TP_ID = dbo.tProduct.P_SupplierID
WHERE     (ISNULL(dbo.tProduct.P_Obsolete, 0) <> 1)
GROUP BY dbo.tProduct.P_ID, dbo.ProductStatusF(dbo.tProduct.P_Status, dbo.tProduct.P_Obsolete) + dbo.tProduct.P_Title, dbo.tProduct.P_EAN, 
                      dbo.tProduct.P_MainAuthor, dbo.tProduct.P_Publisher, dbo.tProduct.P_LastQtyFirmOrdered, dbo.tProduct.P_LastQtySSOrdered, 
                      dbo.tProduct.P_LastQtyDelivered, dbo.tProduct.P_LastPriceDelivered, dbo.tProduct.P_QtyOnHand, dbo.tProduct.P_QtyOnOrder, 
                      dbo.tProduct.P_QtyOnBackorder, dbo.tProduct.P_LastDateDelivered, dbo.tProduct.P_LastDateOrdered, dbo.tProduct.P_LastPriceOrdered, 
                      dbo.tDeal.DL_Description, dbo.tTP.TP_Name, dbo.tProduct.P_SupplierID, dbo.tProduct.P_DealID, dbo.tProduct.P_Code, dbo.tProduct.P_RRP, 
                      dbo.tProduct.P_SP, dbo.tPT.PT_Code, dbo.tProduct.P_QtyOnAppro, dbo.FormatSales(ISNULL(dbo.vMonthSalesPivot.m1, 0), 
                      ISNULL(dbo.vMonthSalesPivot.m2, 0), ISNULL(dbo.vMonthSalesPivot.m3, 0), ISNULL(dbo.vMonthSalesPivot.m4, 0), ISNULL(dbo.vMonthSalesPivot.m5, 
                      0), ISNULL(dbo.vMonthSalesPivot.m6, 0), ''m''), dbo.FormatSales(ISNULL(dbo.vWeekSalesPivot.w1, 0), ISNULL(dbo.vWeekSalesPivot.w2, 0), 
                      ISNULL(dbo.vWeekSalesPivot.w3, 0), ISNULL(dbo.vWeekSalesPivot.w4, 0), ISNULL(dbo.vWeekSalesPivot.w5, 0), ISNULL(dbo.vWeekSalesPivot.w6, 0), 
                      ''w''), dbo.tProduct.P_Title, dbo.GetForeignPrice(dbo.tProduct.P_EUPrice, dbo.tProduct.P_UKPrice, dbo.tProduct.P_USPrice, 
                      dbo.tCurrency.CURR_SYSTEM), dbo.FlattenSections(dbo.tProduct.P_ID), dbo.GetWorkingDeal(dbo.tProduct.P_DealID, dbo.tTP.TP_ID), 
                      dbo.tProduct.P_QtyOnOrder_UnIssued')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vReorder_Browsed View Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.vReorder_Browsed View'
END
GO

--
-- Script To Update dbo.vReorder_SALES View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.vReorder_SALES View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER VIEW dbo.vReorder_SALES
AS
SELECT     dbo.tProduct.P_ID, dbo.ProductStatusF(dbo.tProduct.P_Status, dbo.tProduct.P_Obsolete) + dbo.tProduct.P_Title AS Description, dbo.tProduct.P_EAN, 
                      dbo.tProduct.P_MainAuthor, dbo.tProduct.P_Publisher, dbo.tProduct.P_LastQtyFirmOrdered, dbo.tProduct.P_LastQtySSOrdered, 
                      dbo.tProduct.P_LastQtyDelivered, dbo.tProduct.P_LastPriceDelivered, dbo.tProduct.P_QtyOnHand, dbo.tProduct.P_QtyOnOrder, 
                      dbo.tProduct.P_QtyOnBackorder, dbo.tProduct.P_LastDateDelivered, dbo.tProduct.P_LastDateOrdered, dbo.tProduct.P_LastPriceOrdered, 
                      dbo.tDeal.DL_Description, dbo.tTP.TP_Name AS SUPPLIERNAME, dbo.tProduct.P_SupplierID, dbo.tProduct.P_DealID, dbo.tProduct.P_Code, 
                      SUM(dbo.vAllSalesOnly_Consolidated.Qty) AS QTY, dbo.tProduct.P_RRP, dbo.tProduct.P_SP, MAX(dbo.vAllSalesOnly_Consolidated.Dte) AS Dte, 
                      dbo.tPT.PT_Code AS PT, dbo.tProduct.P_QtyOnAppro, dbo.FormatSales(ISNULL(dbo.vWeekSalesPivot.w1, 0), ISNULL(dbo.vWeekSalesPivot.w2, 0), 
                      ISNULL(dbo.vWeekSalesPivot.w3, 0), ISNULL(dbo.vWeekSalesPivot.w4, 0), ISNULL(dbo.vWeekSalesPivot.w5, 0), ISNULL(dbo.vWeekSalesPivot.w6, 0), 
                      ''w'') AS LASTSIXWEEKS, dbo.FormatSales(ISNULL(dbo.vMonthSalesPivot.m1, 0), ISNULL(dbo.vMonthSalesPivot.m2, 0), 
                      ISNULL(dbo.vMonthSalesPivot.m3, 0), ISNULL(dbo.vMonthSalesPivot.m4, 0), ISNULL(dbo.vMonthSalesPivot.m5, 0), ISNULL(dbo.vMonthSalesPivot.m6, 
                      0), ''m'') AS LASTSIXMONTHS, dbo.tProduct.P_Title, dbo.GetForeignPrice(dbo.tProduct.P_EUPrice, dbo.tProduct.P_UKPrice, dbo.tProduct.P_USPrice, 
                      dbo.tCurrency.CURR_SYSTEM) AS ForeignCurrency, dbo.FlattenSections(dbo.tProduct.P_ID) AS Sections, dbo.GetWorkingDeal(dbo.tProduct.P_DealID, 
                      dbo.tTP.TP_ID) AS DefaultDeal, ISNULL(dbo.tProduct.P_QtyOnOrder_UnIssued, 0) AS P_QtyOnOrder_UnIssued
FROM         dbo.tCurrency RIGHT OUTER JOIN
                      dbo.tTP ON dbo.tCurrency.CURR_ID = dbo.tTP.TP_CURR_ID RIGHT OUTER JOIN
                      dbo.tPT INNER JOIN
                      dbo.tProduct ON dbo.tPT.PT_ID = dbo.tProduct.P_ProductType_ID INNER JOIN
                      dbo.vAllSalesOnly_Consolidated ON dbo.tProduct.P_ID = dbo.vAllSalesOnly_Consolidated.PID LEFT OUTER JOIN
                      dbo.vMonthSalesPivot ON dbo.tProduct.P_ID = dbo.vMonthSalesPivot.PID LEFT OUTER JOIN
                      dbo.vWeekSalesPivot ON dbo.tProduct.P_ID = dbo.vWeekSalesPivot.PID LEFT OUTER JOIN
                      dbo.tDeal ON dbo.tProduct.P_DealID = dbo.tDeal.DL_ID ON dbo.tTP.TP_ID = dbo.tProduct.P_SupplierID
WHERE     (ISNULL(dbo.tProduct.P_Obsolete, 0) <> 1)
GROUP BY dbo.tProduct.P_ID, dbo.ProductStatusF(dbo.tProduct.P_Status, dbo.tProduct.P_Obsolete) + dbo.tProduct.P_Title, dbo.tProduct.P_EAN, 
                      dbo.tProduct.P_MainAuthor, dbo.tProduct.P_Publisher, dbo.tProduct.P_LastQtyFirmOrdered, dbo.tProduct.P_LastQtySSOrdered, 
                      dbo.tProduct.P_LastQtyDelivered, dbo.tProduct.P_LastPriceDelivered, dbo.tProduct.P_QtyOnHand, dbo.tProduct.P_QtyOnOrder, 
                      dbo.tProduct.P_QtyOnBackorder, dbo.tProduct.P_LastDateDelivered, dbo.tProduct.P_LastDateOrdered, dbo.tProduct.P_LastPriceOrdered, 
                      dbo.tDeal.DL_Description, dbo.tTP.TP_Name, dbo.tProduct.P_SupplierID, dbo.tProduct.P_DealID, dbo.tProduct.P_Code, dbo.tProduct.P_RRP, 
                      dbo.tProduct.P_SP, dbo.tPT.PT_Code, dbo.tProduct.P_QtyOnAppro, dbo.FormatSales(ISNULL(dbo.vMonthSalesPivot.m1, 0), 
                      ISNULL(dbo.vMonthSalesPivot.m2, 0), ISNULL(dbo.vMonthSalesPivot.m3, 0), ISNULL(dbo.vMonthSalesPivot.m4, 0), ISNULL(dbo.vMonthSalesPivot.m5, 
                      0), ISNULL(dbo.vMonthSalesPivot.m6, 0), ''m''), dbo.FormatSales(ISNULL(dbo.vWeekSalesPivot.w1, 0), ISNULL(dbo.vWeekSalesPivot.w2, 0), 
                      ISNULL(dbo.vWeekSalesPivot.w3, 0), ISNULL(dbo.vWeekSalesPivot.w4, 0), ISNULL(dbo.vWeekSalesPivot.w5, 0), ISNULL(dbo.vWeekSalesPivot.w6, 0), 
                      ''w''), dbo.tProduct.P_Title, dbo.GetForeignPrice(dbo.tProduct.P_EUPrice, dbo.tProduct.P_UKPrice, dbo.tProduct.P_USPrice, 
                      dbo.tCurrency.CURR_SYSTEM), dbo.FlattenSections(dbo.tProduct.P_ID), dbo.GetWorkingDeal(dbo.tProduct.P_DealID, dbo.tTP.TP_ID), 
                      ISNULL(dbo.tProduct.P_QtyOnOrder_UnIssued, 0)')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vReorder_SALES View Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.vReorder_SALES View'
END
GO

--
-- Script To Update dbo.vReorder_TRANSFERS View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.vReorder_TRANSFERS View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER VIEW dbo.vReorder_TRANSFERS
AS
SELECT     dbo.tProduct.P_ID, dbo.ProductStatusF(dbo.tProduct.P_Status, dbo.tProduct.P_Obsolete) + dbo.tProduct.P_Title AS Description, dbo.tProduct.P_EAN, 
                      dbo.tProduct.P_MainAuthor, dbo.tProduct.P_Publisher, dbo.tProduct.P_LastQtyFirmOrdered, dbo.tProduct.P_LastQtySSOrdered, 
                      dbo.tProduct.P_LastQtyDelivered, dbo.tProduct.P_LastPriceDelivered, dbo.tProduct.P_QtyOnHand, dbo.tProduct.P_QtyOnOrder, 
                      dbo.tProduct.P_QtyOnBackorder, dbo.tProduct.P_LastDateDelivered, dbo.tProduct.P_LastDateOrdered, dbo.tProduct.P_LastPriceOrdered, 
                      dbo.tDeal.DL_Description, dbo.tTP.TP_Name AS SUPPLIERNAME, dbo.tProduct.P_SupplierID, dbo.tProduct.P_DealID, dbo.tProduct.P_Code, 
                      SUM(ISNULL(dbo.tSales_CYr.SCY_Qty, 0)) AS QTY, dbo.tProduct.P_RRP, dbo.tProduct.P_SP, MAX(dbo.tTR.TR_Date) AS dte, dbo.tPT.PT_Code AS PT, 
                      dbo.tProduct.P_QtyOnAppro, dbo.FormatSales(ISNULL(dbo.vWeekSalesPivot.w1, 0), ISNULL(dbo.vWeekSalesPivot.w2, 0), 
                      ISNULL(dbo.vWeekSalesPivot.w3, 0), ISNULL(dbo.vWeekSalesPivot.w4, 0), ISNULL(dbo.vWeekSalesPivot.w5, 0), ISNULL(dbo.vWeekSalesPivot.w6, 0), 
                      ''w'') AS LASTSIXWEEKS, dbo.FormatSales(ISNULL(dbo.vMonthSalesPivot.m1, 0), ISNULL(dbo.vMonthSalesPivot.m2, 0), 
                      ISNULL(dbo.vMonthSalesPivot.m3, 0), ISNULL(dbo.vMonthSalesPivot.m4, 0), ISNULL(dbo.vMonthSalesPivot.m5, 0), ISNULL(dbo.vMonthSalesPivot.m6, 
                      0), ''m'') AS LASTSIXMONTHS, dbo.tProduct.P_Title, dbo.GetForeignPrice(dbo.tProduct.P_EUPrice, dbo.tProduct.P_UKPrice, dbo.tProduct.P_USPrice, 
                      dbo.tCurrency.CURR_SYSTEM) AS ForeignCurrency, dbo.FlattenSections(dbo.tProduct.P_ID) AS Sections, dbo.tTFRL.TFRL_TR_ID, 
                      dbo.GetWorkingDeal(dbo.tProduct.P_DealID, dbo.tTP.TP_ID) AS DefaultDeal, ISNULL(dbo.tProduct.P_QtyOnOrder_UnIssued, 0) 
                      AS P_QtyOnOrder_UnIssued
FROM         dbo.tCurrency RIGHT OUTER JOIN
                      dbo.tTP ON dbo.tCurrency.CURR_ID = dbo.tTP.TP_CURR_ID RIGHT OUTER JOIN
                      dbo.tPT INNER JOIN
                      dbo.tProduct ON dbo.tPT.PT_ID = dbo.tProduct.P_ProductType_ID INNER JOIN
                      dbo.tTFRL ON dbo.tProduct.P_ID = dbo.tTFRL.TFRL_P_ID INNER JOIN
                      dbo.tTFR ON dbo.tTFRL.TFRL_TR_ID = dbo.tTFR.TFR_ID INNER JOIN
                      dbo.tTR ON dbo.tTFRL.TFRL_TR_ID = dbo.tTR.TR_ID LEFT OUTER JOIN
                      dbo.tSales_CYr ON dbo.tProduct.P_ID = dbo.tSales_CYr.SCY_P_ID LEFT OUTER JOIN
                      dbo.vMonthSalesPivot ON dbo.tProduct.P_ID = dbo.vMonthSalesPivot.PID LEFT OUTER JOIN
                      dbo.vWeekSalesPivot ON dbo.tProduct.P_ID = dbo.vWeekSalesPivot.PID LEFT OUTER JOIN
                      dbo.tDeal ON dbo.tProduct.P_DealID = dbo.tDeal.DL_ID ON dbo.tTP.TP_ID = dbo.tProduct.P_SupplierID
WHERE     (ISNULL(dbo.tProduct.P_Obsolete, 0) <> 1) AND (dbo.tTFR.TFR_INOUT = ''OUT'')
GROUP BY dbo.tProduct.P_ID, dbo.ProductStatusF(dbo.tProduct.P_Status, dbo.tProduct.P_Obsolete) + dbo.tProduct.P_Title, dbo.tProduct.P_EAN, 
                      dbo.tProduct.P_MainAuthor, dbo.tProduct.P_Publisher, dbo.tProduct.P_LastQtyFirmOrdered, dbo.tProduct.P_LastQtySSOrdered, 
                      dbo.tProduct.P_LastQtyDelivered, dbo.tProduct.P_LastPriceDelivered, dbo.tProduct.P_QtyOnHand, dbo.tProduct.P_QtyOnOrder, 
                      dbo.tProduct.P_QtyOnBackorder, dbo.tProduct.P_LastDateDelivered, dbo.tProduct.P_LastDateOrdered, dbo.tProduct.P_LastPriceOrdered, 
                      dbo.tDeal.DL_Description, dbo.tTP.TP_Name, dbo.tProduct.P_SupplierID, dbo.tProduct.P_DealID, dbo.tProduct.P_Code, dbo.tProduct.P_RRP, 
                      dbo.tProduct.P_SP, dbo.tPT.PT_Code, dbo.tProduct.P_QtyOnAppro, dbo.FormatSales(ISNULL(dbo.vWeekSalesPivot.w1, 0), 
                      ISNULL(dbo.vWeekSalesPivot.w2, 0), ISNULL(dbo.vWeekSalesPivot.w3, 0), ISNULL(dbo.vWeekSalesPivot.w4, 0), ISNULL(dbo.vWeekSalesPivot.w5, 0), 
                      ISNULL(dbo.vWeekSalesPivot.w6, 0), ''w''), dbo.FormatSales(ISNULL(dbo.vMonthSalesPivot.m1, 0), ISNULL(dbo.vMonthSalesPivot.m2, 0), 
                      ISNULL(dbo.vMonthSalesPivot.m3, 0), ISNULL(dbo.vMonthSalesPivot.m4, 0), ISNULL(dbo.vMonthSalesPivot.m5, 0), ISNULL(dbo.vMonthSalesPivot.m6, 
                      0), ''m''), dbo.tProduct.P_Title, dbo.GetForeignPrice(dbo.tProduct.P_EUPrice, dbo.tProduct.P_UKPrice, dbo.tProduct.P_USPrice, 
                      dbo.tCurrency.CURR_SYSTEM), dbo.FlattenSections(dbo.tProduct.P_ID), dbo.tTFRL.TFRL_TR_ID, dbo.GetWorkingDeal(dbo.tProduct.P_DealID, 
                      dbo.tTP.TP_ID), ISNULL(dbo.tProduct.P_QtyOnOrder_UnIssued, 0)')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vReorder_TRANSFERS View Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.vReorder_TRANSFERS View'
END
GO

--
-- Script To Update dbo.vReorderCust View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.vReorderCust View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER VIEW dbo.vReorderCust
AS
SELECT     dbo.tProduct.P_ID, dbo.ProductStatusF(dbo.tProduct.P_Status, dbo.tProduct.P_Obsolete) + dbo.tProduct.P_Title AS Description, dbo.tProduct.P_EAN, 
                      dbo.tProduct.P_MainAuthor, dbo.tProduct.P_Publisher, dbo.tProduct.P_LastQtyFirmOrdered, dbo.tProduct.P_LastQtySSOrdered, 
                      dbo.tProduct.P_LastQtyDelivered, dbo.tProduct.P_LastPriceDelivered, dbo.tProduct.P_QtyOnHand, dbo.tProduct.P_QtyOnOrder, 
                      dbo.tProduct.P_QtyOnBackorder, dbo.tProduct.P_QtyTotalSold, dbo.tProduct.P_LastDateDelivered, dbo.tProduct.P_LastDateOrdered, 
                      dbo.tProduct.P_LastPriceOrdered, dbo.tDeal.DL_Description, tTP_1.TP_Name AS SUPPLIERNAME, dbo.tProduct.P_SupplierID, dbo.tProduct.P_DealID, 
                      dbo.tProduct.P_Code, dbo.tProduct.P_RRP, dbo.tProduct.P_SP, dbo.tPT.PT_Code AS PT, dbo.tProduct.P_QtyOnAppro, 
                      dbo.FormatSales(ISNULL(dbo.vWeekSalesPivot.w1, 0), ISNULL(dbo.vWeekSalesPivot.w2, 0), ISNULL(dbo.vWeekSalesPivot.w3, 0), 
                      ISNULL(dbo.vWeekSalesPivot.w4, 0), ISNULL(dbo.vWeekSalesPivot.w5, 0), ISNULL(dbo.vWeekSalesPivot.w6, 0), ''w'') AS LASTSIXWEEKS, 
                      dbo.FormatSales(ISNULL(dbo.vMonthSalesPivot.m1, 0), ISNULL(dbo.vMonthSalesPivot.m2, 0), ISNULL(dbo.vMonthSalesPivot.m3, 0), 
                      ISNULL(dbo.vMonthSalesPivot.m4, 0), ISNULL(dbo.vMonthSalesPivot.m5, 0), ISNULL(dbo.vMonthSalesPivot.m6, 0), ''m'') AS LSTSIXMONTHS, 
                      dbo.tProduct.P_Status, dbo.tProduct.P_Title, ISNULL(dbo.tCOL.COL_ActionTaken, 0) AS ActionTaken, dbo.GetForeignPrice(dbo.tProduct.P_EUPrice, 
                      dbo.tProduct.P_UKPrice, dbo.tProduct.P_USPrice, dbo.tCurrency.CURR_SYSTEM) AS ForeignPrice, dbo.FlattenSections(dbo.tProduct.P_ID) AS Sections, 
                      dbo.GetWorkingDeal(dbo.tProduct.P_DealID, tTP_1.TP_ID) AS DefaultDeal, dbo.tProduct.P_QtyOnOrder_UnIssued
FROM         dbo.tCurrency RIGHT OUTER JOIN
                      dbo.tTP AS tTP_1 ON dbo.tCurrency.CURR_ID = tTP_1.TP_CURR_ID RIGHT OUTER JOIN
                      dbo.tCOL INNER JOIN
                      dbo.tProduct ON dbo.tCOL.COL_P_ID = dbo.tProduct.P_ID INNER JOIN
                      dbo.tCO ON dbo.tCOL.COL_TR_ID = dbo.tCO.CO_ID INNER JOIN
                      dbo.tPT ON dbo.tProduct.P_ProductType_ID = dbo.tPT.PT_ID INNER JOIN
                      dbo.tTR ON dbo.tCO.CO_ID = dbo.tTR.TR_ID INNER JOIN
                      dbo.tTP ON dbo.tTR.TR_TP_ID = dbo.tTP.TP_ID LEFT OUTER JOIN
                      dbo.vMonthSalesPivot ON dbo.tProduct.P_ID = dbo.vMonthSalesPivot.PID LEFT OUTER JOIN
                      dbo.vWeekSalesPivot ON dbo.tProduct.P_ID = dbo.vWeekSalesPivot.PID LEFT OUTER JOIN
                      dbo.vValidPOLs ON dbo.tCOL.COL_ID = dbo.vValidPOLs.POL_COLID LEFT OUTER JOIN
                      dbo.tDeal ON dbo.tProduct.P_DealID = dbo.tDeal.DL_ID ON tTP_1.TP_ID = dbo.tProduct.P_SupplierID
WHERE     (NOT (dbo.tCOL.COL_Fulfilled IN (''FUL'', ''CAN'', ''8''))) AND (dbo.tCO.CO_OrderType = 1) AND (dbo.vValidPOLs.POL_COLID IS NULL) AND 
                      (dbo.tTR.TR_Status = 3) AND (NOT (ISNULL(dbo.tCOL.COL_ActionTaken, 0) IN (3)))
GROUP BY dbo.tProduct.P_ID, dbo.ProductStatusF(dbo.tProduct.P_Status, dbo.tProduct.P_Obsolete) + dbo.tProduct.P_Title, dbo.tProduct.P_EAN, 
                      dbo.tProduct.P_MainAuthor, dbo.tProduct.P_Publisher, dbo.tProduct.P_LastQtyFirmOrdered, dbo.tProduct.P_LastQtySSOrdered, 
                      dbo.tProduct.P_LastQtyDelivered, dbo.tProduct.P_LastPriceDelivered, dbo.tProduct.P_QtyOnHand, dbo.tProduct.P_QtyOnOrder, 
                      dbo.tProduct.P_QtyOnBackorder, dbo.tProduct.P_QtyTotalSold, dbo.tProduct.P_LastDateDelivered, dbo.tProduct.P_LastDateOrdered, 
                      dbo.tProduct.P_LastPriceOrdered, dbo.tDeal.DL_Description, tTP_1.TP_Name, dbo.tProduct.P_SupplierID, dbo.tProduct.P_DealID, dbo.tProduct.P_Code, 
                      dbo.tProduct.P_RRP, dbo.tProduct.P_SP, dbo.tPT.PT_Code, dbo.tProduct.P_QtyOnAppro, dbo.FormatSales(ISNULL(dbo.vWeekSalesPivot.w1, 0), 
                      ISNULL(dbo.vWeekSalesPivot.w2, 0), ISNULL(dbo.vWeekSalesPivot.w3, 0), ISNULL(dbo.vWeekSalesPivot.w4, 0), ISNULL(dbo.vWeekSalesPivot.w5, 0), 
                      ISNULL(dbo.vWeekSalesPivot.w6, 0), ''w''), dbo.FormatSales(ISNULL(dbo.vMonthSalesPivot.m1, 0), ISNULL(dbo.vMonthSalesPivot.m2, 0), 
                      ISNULL(dbo.vMonthSalesPivot.m3, 0), ISNULL(dbo.vMonthSalesPivot.m4, 0), ISNULL(dbo.vMonthSalesPivot.m5, 0), ISNULL(dbo.vMonthSalesPivot.m6, 
                      0), ''m''), dbo.tProduct.P_Status, dbo.tProduct.P_Title, ISNULL(dbo.tCOL.COL_ActionTaken, 0), dbo.GetForeignPrice(dbo.tProduct.P_EUPrice, 
                      dbo.tProduct.P_UKPrice, dbo.tProduct.P_USPrice, dbo.tCurrency.CURR_SYSTEM), dbo.FlattenSections(dbo.tProduct.P_ID), 
                      dbo.GetWorkingDeal(dbo.tProduct.P_DealID, tTP_1.TP_ID), dbo.tProduct.P_QtyOnOrder_UnIssued')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vReorderCust View Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.vReorderCust View'
END
GO

--
-- Script To Update dbo.vReorderCustByCOL View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.vReorderCustByCOL View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER VIEW dbo.vReorderCustByCOL
AS
SELECT     tTR_2.TR_Status AS status2, dbo.tProduct.P_ID, dbo.ProductStatusF(dbo.tProduct.P_Status, dbo.tProduct.P_Obsolete) 
                      + dbo.tProduct.P_Title AS Description, dbo.tProduct.P_EAN, dbo.tProduct.P_MainAuthor, dbo.tProduct.P_Publisher, dbo.tProduct.P_LastQtyFirmOrdered, 
                      dbo.tProduct.P_LastQtySSOrdered, dbo.tProduct.P_LastQtyDelivered, dbo.tProduct.P_LastPriceDelivered, dbo.tProduct.P_QtyOnHand, 
                      dbo.tProduct.P_QtyOnOrder, dbo.tProduct.P_QtyOnBackorder, dbo.tProduct.P_QtyTotalSold, dbo.tProduct.P_LastDateDelivered, 
                      dbo.tProduct.P_LastDateOrdered, dbo.tProduct.P_LastPriceOrdered, dbo.tDeal.DL_Description, tTP_1.TP_Name AS SUPPLIERNAME, 
                      dbo.tProduct.P_SupplierID, dbo.tProduct.P_DealID, dbo.tProduct.P_Code, dbo.tProduct.P_RRP, dbo.tProduct.P_SP, tTR_2.TR_Date, 
                      dbo.tPT.PT_Code AS PT, dbo.tCOL.COL_Ref, dbo.tCO.CO_OrderNUm, dbo.tCOL.COL_ID, dbo.tCOL.COL_Qty, dbo.tCOL.COL_QtyDispatched, 
                      tTR_2.TR_STAFFID, dbo.tProduct.P_QtyOnAppro, dbo.FormatSales(ISNULL(dbo.vWeekSalesPivot.w1, 0), ISNULL(dbo.vWeekSalesPivot.w2, 0), 
                      ISNULL(dbo.vWeekSalesPivot.w3, 0), ISNULL(dbo.vWeekSalesPivot.w4, 0), ISNULL(dbo.vWeekSalesPivot.w5, 0), ISNULL(dbo.vWeekSalesPivot.w6, 0), 
                      ''w'') AS LASTSIXWEEKS, dbo.FormatSales(ISNULL(dbo.vMonthSalesPivot.m1, 0), ISNULL(dbo.vMonthSalesPivot.m2, 0), 
                      ISNULL(dbo.vMonthSalesPivot.m3, 0), ISNULL(dbo.vMonthSalesPivot.m4, 0), ISNULL(dbo.vMonthSalesPivot.m5, 0), ISNULL(dbo.vMonthSalesPivot.m6, 
                      0), ''m'') AS LASTSIXMONTHS, dbo.tProduct.P_Title, dbo.tCOL.COL_ActionTaken AS ActionTaken, dbo.tTP.TP_ACNo, 
                      dbo.GetForeignPrice(dbo.tProduct.P_EUPrice, dbo.tProduct.P_UKPrice, dbo.tProduct.P_USPrice, dbo.tCurrency.CURR_SYSTEM) AS ForeignPrice, 
                      dbo.FlattenSections(dbo.tProduct.P_ID) AS Sections, dbo.GetWorkingDeal(dbo.tProduct.P_DealID, tTP_1.TP_ID) AS DefaultDeal, 
                      dbo.tProduct.P_QtyOnOrder_UnIssued
FROM         dbo.tCurrency RIGHT OUTER JOIN
                      dbo.tTP AS tTP_1 ON dbo.tCurrency.CURR_ID = tTP_1.TP_CURR_ID RIGHT OUTER JOIN
                      dbo.tCOL INNER JOIN
                      dbo.tTR AS tTR_2 ON dbo.tCOL.COL_TR_ID = tTR_2.TR_ID INNER JOIN
                      dbo.tProduct ON dbo.tCOL.COL_P_ID = dbo.tProduct.P_ID INNER JOIN
                      dbo.tCO ON dbo.tCOL.COL_TR_ID = dbo.tCO.CO_ID INNER JOIN
                      dbo.tPT ON dbo.tProduct.P_ProductType_ID = dbo.tPT.PT_ID INNER JOIN
                      dbo.tTP ON tTR_2.TR_TP_ID = dbo.tTP.TP_ID LEFT OUTER JOIN
                      dbo.vMonthSalesPivot ON dbo.tProduct.P_ID = dbo.vMonthSalesPivot.PID LEFT OUTER JOIN
                      dbo.vWeekSalesPivot ON dbo.tProduct.P_ID = dbo.vWeekSalesPivot.PID LEFT OUTER JOIN
                      dbo.vValidPOLs ON dbo.tCOL.COL_ID = dbo.vValidPOLs.POL_COLID LEFT OUTER JOIN
                      dbo.tDeal ON dbo.tProduct.P_DealID = dbo.tDeal.DL_ID ON tTP_1.TP_ID = dbo.tProduct.P_SupplierID
WHERE     (dbo.tCOL.COL_Fulfilled <> ''FUL'') AND (dbo.tCOL.COL_Fulfilled <> ''CAN'') AND (dbo.tCOL.COL_Fulfilled <> ''8'') AND (dbo.tCO.CO_OrderType = 1) AND 
                      (dbo.tCOL.COL_ActionTaken <> 3) AND (dbo.vValidPOLs.POL_COLID IS NULL)
GROUP BY tTR_2.TR_Status, dbo.tProduct.P_ID, dbo.ProductStatusF(dbo.tProduct.P_Status, dbo.tProduct.P_Obsolete) + dbo.tProduct.P_Title, 
                      dbo.tProduct.P_EAN, dbo.tProduct.P_MainAuthor, dbo.tProduct.P_Publisher, dbo.tProduct.P_LastQtyFirmOrdered, dbo.tProduct.P_LastQtySSOrdered, 
                      dbo.tProduct.P_LastQtyDelivered, dbo.tProduct.P_LastPriceDelivered, dbo.tProduct.P_QtyOnHand, dbo.tProduct.P_QtyOnOrder, 
                      dbo.tProduct.P_QtyOnBackorder, dbo.tProduct.P_QtyTotalSold, dbo.tProduct.P_LastDateDelivered, dbo.tProduct.P_LastDateOrdered, 
                      dbo.tProduct.P_LastPriceOrdered, dbo.tDeal.DL_Description, tTP_1.TP_Name, dbo.tProduct.P_SupplierID, dbo.tProduct.P_DealID, dbo.tProduct.P_Code, 
                      dbo.tProduct.P_RRP, dbo.tProduct.P_SP, tTR_2.TR_Date, dbo.tPT.PT_Code, dbo.tCOL.COL_Ref, dbo.tCO.CO_OrderNUm, dbo.tCOL.COL_ID, 
                      dbo.tCOL.COL_Qty, dbo.tCOL.COL_QtyDispatched, tTR_2.TR_STAFFID, dbo.tProduct.P_QtyOnAppro, 
                      dbo.FormatSales(ISNULL(dbo.vMonthSalesPivot.m1, 0), ISNULL(dbo.vMonthSalesPivot.m2, 0), ISNULL(dbo.vMonthSalesPivot.m3, 0), 
                      ISNULL(dbo.vMonthSalesPivot.m4, 0), ISNULL(dbo.vMonthSalesPivot.m5, 0), ISNULL(dbo.vMonthSalesPivot.m6, 0), ''m''), 
                      dbo.FormatSales(ISNULL(dbo.vWeekSalesPivot.w1, 0), ISNULL(dbo.vWeekSalesPivot.w2, 0), ISNULL(dbo.vWeekSalesPivot.w3, 0), 
                      ISNULL(dbo.vWeekSalesPivot.w4, 0), ISNULL(dbo.vWeekSalesPivot.w5, 0), ISNULL(dbo.vWeekSalesPivot.w6, 0), ''w''), dbo.tProduct.P_Title, 
                      dbo.tCOL.COL_ActionTaken, dbo.tTP.TP_ACNo, dbo.GetForeignPrice(dbo.tProduct.P_EUPrice, dbo.tProduct.P_UKPrice, dbo.tProduct.P_USPrice, 
                      dbo.tCurrency.CURR_SYSTEM), dbo.FlattenSections(dbo.tProduct.P_ID), dbo.GetWorkingDeal(dbo.tProduct.P_DealID, tTP_1.TP_ID), 
                      dbo.tProduct.P_QtyOnOrder_UnIssued
HAVING      (tTR_2.TR_Status = 3)')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vReorderCustByCOL View Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.vReorderCustByCOL View'
END
GO

--
-- Script To Delete dbo.vvvPastelExport2_TaxInvoices_Master View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Deleting dbo.vvvPastelExport2_TaxInvoices_Master View'
GO

   DROP VIEW [dbo].[vvvPastelExport2_TaxInvoices_Master]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vvvPastelExport2_TaxInvoices_Master View Deleted Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Delete dbo.vvvPastelExport2_TaxInvoices_Master View'
END
GO

--
-- Script To Delete dbo.vvvPastelExport_Sales_Master View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Deleting dbo.vvvPastelExport_Sales_Master View'
GO

   DROP VIEW [dbo].[vvvPastelExport_Sales_Master]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vvvPastelExport_Sales_Master View Deleted Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Delete dbo.vvvPastelExport_Sales_Master View'
END
GO

--
-- Script To Delete dbo.vvvPastelExport_SInvoice_Master View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Deleting dbo.vvvPastelExport_SInvoice_Master View'
GO

   DROP VIEW [dbo].[vvvPastelExport_SInvoice_Master]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vvvPastelExport_SInvoice_Master View Deleted Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Delete dbo.vvvPastelExport_SInvoice_Master View'
END
GO

--
-- Script To Delete dbo.vvvPASTELExport_Supp_Master View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Deleting dbo.vvvPASTELExport_Supp_Master View'
GO

   DROP VIEW [dbo].[vvvPASTELExport_Supp_Master]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vvvPASTELExport_Supp_Master View Deleted Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Delete dbo.vvvPASTELExport_Supp_Master View'
END
GO

--
-- Script To Delete dbo.vvvPASTELExport_TAXCREDITNOTES_Master View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Deleting dbo.vvvPASTELExport_TAXCREDITNOTES_Master View'
GO

   DROP VIEW [dbo].[vvvPASTELExport_TAXCREDITNOTES_Master]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vvvPASTELExport_TAXCREDITNOTES_Master View Deleted Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Delete dbo.vvvPASTELExport_TAXCREDITNOTES_Master View'
END
GO

--
-- Script To Delete dbo.vvvPastelExport_TAXINVOICES_Master View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Deleting dbo.vvvPastelExport_TAXINVOICES_Master View'
GO

   DROP VIEW [dbo].[vvvPastelExport_TAXINVOICES_Master]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vvvPastelExport_TAXINVOICES_Master View Deleted Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Delete dbo.vvvPastelExport_TAXINVOICES_Master View'
END
GO

--
-- Script To Update dbo.zReorder_SalesandTransfers View In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.zReorder_SalesandTransfers View'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER VIEW dbo.zReorder_SalesandTransfers
AS
SELECT     P_ID, Description, P_EAN, P_MainAuthor, P_Publisher, P_LastQtyFirmOrdered, P_LastQtySSOrdered, P_LastQtyDelivered, P_LastPriceDelivered, 
                      P_QtyOnHand, P_QtyOnOrder, P_QtyOnBackorder, P_LastDateDelivered, P_LastDateOrdered, P_LastPriceOrdered, DL_Description, SUPPLIERNAME, 
                      P_SupplierID, P_DealID, P_Code, QTY, P_RRP, P_SP, Dte, PT, P_QtyOnAppro, LASTSIXWEEKS, LASTSIXMONTHS, P_Title AS TITLEFORSORT, 
                      Sections, DefaultDeal, P_QtyOnOrder_UnIssued
FROM         dbo.vReorder_SALES
UNION
SELECT     P_ID, Description, P_EAN, P_MainAuthor, P_Publisher, P_LastQtyFirmOrdered, P_LastQtySSOrdered, P_LastQtyDelivered, P_LastPriceDelivered, 
                      P_QtyOnHand, P_QtyOnOrder, P_QtyOnBackorder, P_LastDateDelivered, P_LastDateOrdered, P_LastPriceOrdered, DL_Description, SUPPLIERNAME, 
                      P_SupplierID, P_DealID, P_Code, QTY, P_RRP, P_SP, dte, PT, P_QtyOnAppro, LASTSIXWEEKS, LASTSIXMONTHS, P_Title, Sections, DefaultDeal, 
                      P_QtyOnOrder_UnIssued
FROM         dbo.vReorder_TRANSFERS')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.zReorder_SalesandTransfers View Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.zReorder_SalesandTransfers View'
END
GO

--
-- Script To Update dbo._actp_Invocation_RQ_Consumer Procedure In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo._actp_Invocation_RQ_Consumer Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER PROCEDURE [dbo].[_actp_Invocation_RQ_Consumer]
AS
BEGIN
DECLARE @DATEFROM DATETIME
DECLARE @DATETO DATETIME
DECLARE @handle UNIQUEIDENTIFIER
DECLARE @messageTypeName SYSNAME
DECLARE @messageBody XML
DECLARE @ErrorString NVARCHAR(4000)
DECLARE @INVOCATION_RQ_Body XML
DECLARE @CUSTOMERSTATS_Body XML
DECLARE @RES INT
DECLARE @ERRMESS VARCHAR(500)
DECLARE @EAN VARCHAR(20)
DECLARE @INSTALLATIONCODE VARCHAR(10)
DECLARE @REQID VARCHAR(20)
DECLARE @hdoc INT  
DECLARE @CustCount INT
DECLARE @AddressCount INT
DECLARE @LastUpdateFromCentral DATETIME
DECLARE @LastUpdateToCentral DATETIME
DECLARE @LASTADDEDACNO VARCHAR(20)
DECLARE @LASTADDEDACDATE DATETIME
DECLARE @QtyOHBody XML
DECLARE @RESULTXML XML
DECLARE @TT VARCHAR(80)

DECLARE @DATERANGE TABLE 
	(
		DATEFROM VARCHAR(20),
		DATETO VARCHAR(20)
	)

	DECLARE @T TABLE 
	(
		ACNO VARCHAR(20)
	)

BEGIN TRANSACTION;

	BEGIN TRY;
			SELECT @INSTALLATIONCODE = CF_INSTALLATIONCODE FROM tConfiguration;
			RECEIVE TOP(1)
				@handle = conversation_handle,
				@messageTypeName = message_type_name,
				@messageBody = message_body
			FROM [INVOCATION_RQ_Q];
			
			IF @handle IS NOT NULL
			BEGIN
				IF @messageTypeName = N''CUSTOMERSTATS_RQ_MSG''
				BEGIN
--=====================================================================
							INSERT INTO _tSBLog (SBL_Msg,SBL_Proc) VALUES (''INVOCATION_RQ_MSG message received'',''[_actp_Invocation_RQ_Consumer]'')
			-----------prelims------------------------------
					SELECT @CustCount =  COUNT(dbo.tTP.TP_ID)
						FROM         dbo.tTP INNER JOIN
							  dbo.tTP_IG ON dbo.tTP.TP_ID = dbo.tTP_IG.TPIG_TP_ID INNER JOIN
							  dbo.tDict ON dbo.tTP_IG.TPIG_IG_ID = dbo.tDict.DICT_ID
						WHERE     (dbo.tDict.DICT_System = ''L1'')
					SELECT @AddressCount =      COUNT(dbo.tAdd.ADD_ID)
						FROM         dbo.tTP_IG INNER JOIN
							  dbo.tDict ON dbo.tTP_IG.TPIG_IG_ID = dbo.tDict.DICT_ID INNER JOIN
							  dbo.tAdd ON dbo.tTP_IG.TPIG_TP_ID = dbo.tAdd.ADD_TP_ID
						WHERE     (dbo.tDict.DICT_System = ''L1'')


					SELECT @LastUpdateFromCentral =  T_LastLoyaltyUpdateFromCentral FROM _tTimerStat;
					SELECT @LastUpdateToCentral =  T_LastLoyaltyTransmission FROM _tTimerStat;

					SELECT @LASTADDEDACNO = dbo.tTP.TP_ACNo, @LASTADDEDACDATE = dbo.tTP.TP_DateRecordAdded
						FROM         dbo.tTP INNER JOIN
								  (SELECT     MAX(tTP_1.TP_ID) AS MAXID
									FROM          dbo.tDict INNER JOIN
														   dbo.tTP_IG ON dbo.tDict.DICT_ID = dbo.tTP_IG.TPIG_IG_ID INNER JOIN
														   dbo.tTP AS tTP_1 ON dbo.tTP_IG.TPIG_TP_ID = tTP_1.TP_ID
									WHERE      (dbo.tDict.DICT_System = ''L1'')) AS m ON dbo.tTP.TP_ID = m.MAXID INNER JOIN
							  dbo.tConfiguration ON dbo.tTP.TP_LOYALTYHOMESTOREID = dbo.tConfiguration.CF_DefaultStoreID;
			-----------------------------
					SELECT @CUSTOMERSTATS_Body = 
						''<CUSTOMERSTATS_MSG>
							<BRCode>'' + ISNULL(@INSTALLATIONCODE,''NULL'') + ''</BRCode>
							<CustCount>'' + CAST(ISNULL(@CustCount,0) AS VARCHAR(10)) + ''</CustCount>
							<AddressCount>'' + CAST(ISNULL(@AddressCount,0) AS VARCHAR(10)) + ''</AddressCount>
							<LastupdateFromCentral>'' + ISNULL(CONVERT(VARCHAR(30),@LastUpdateFromCentral,120),'''') + ''</LastupdateFromCentral>
							<LastSentToCentral>'' + ISNULL(CONVERT(VARCHAR(30),@LastUpdateToCentral,120),'''') + ''</LastSentToCentral>
							<LastCustAdded>'' + ISNULL(@LASTADDEDACNO,'''') + ''</LastCustAdded>
							<LastCustAddedDate>'' + ISNULL(CONVERT(VARCHAR(30),@LASTADDEDACDATE,120),'''') + ''</LastCustAddedDate>
							<StatusDate>'' + ISNULL(CONVERT(VARCHAR(30),GETDATE(),120),'''') + ''</StatusDate>
						</CUSTOMERSTATS_MSG>'';
					SEND ON CONVERSATION @handle MESSAGE TYPE CUSTOMERSTATS_RS_MSG (@CUSTOMERSTATS_Body);
							INSERT INTO _tSBLog (SBL_Msg,SBL_Proc) VALUES (''INVOCATION_RS_MSG message sent'',''[_actp_Invocation_RQ_Consumer]'')

					COMMIT TRANSACTION;
				END
--=======================================================================

				IF @messageTypeName = N''CASHUPS_RQ_MSG''
				BEGIN
					IF dbo.GETPROPERTY(''Cashup_Reporting_ON'') =''TRUE''
					BEGIN
						EXEC sp_xml_preparedocument @hdoc OUTPUT, @messageBody
						INSERT INTO @DATERANGE( DATEFROM,DATETO)
							SELECT y.DATEFROM,y.DATETO
							FROM OPENXML ( @hdoc, ''/DATERANGE_MSG'', 3 ) WITH (
								DATEFROM VARCHAR(20) ''DATEFROM'',
								DATETO VARCHAR(20) ''DATETO''
								) AS y
						EXEC sp_xml_removedocument @hdoc
						SELECT @DATEFROM = CAST(DATEFROM as DATETIME),@DATETO = CAST(DATETO as DATETIME) FROM @DATERANGE;
						SELECT @TT = ''from:'' + ISNULL(DATEFROM,''empty'') + ''to:'' +  ISNULL(DATETO,''empty'') FROM @DATERANGE;
						EXEC dbo._GetCashupsForExportXML @DATEFROM,@DATETO, @RESULTXML OUTPUT;
						SEND ON CONVERSATION @handle MESSAGE TYPE CASHUPS_RS_MSG (@RESULTXML)
					END
					COMMIT TRANSACTION;
				END
--=======================================================================

				IF @messageTypeName = N''COLS_RQ_MSG''
				BEGIN
					IF dbo.GETPROPERTY(''COLS_Reporting_ON'') =''TRUE''
					BEGIN
						EXEC sp_xml_preparedocument @hdoc OUTPUT, @messageBody
						INSERT INTO @DATERANGE( DATEFROM,DATETO)
							SELECT y.DATEFROM,y.DATETO
							FROM OPENXML ( @hdoc, ''/BranchSelection'', 3 ) WITH (
								DATEFROM VARCHAR(20) ''DATEFROM'',
								DATETO VARCHAR(20) ''DATETO''
								) AS y
						EXEC sp_xml_removedocument @hdoc
						SELECT @DATEFROM = CAST(DATEFROM as DATETIME),@DATETO = CAST(DATETO as DATETIME) FROM @DATERANGE;
						EXEC dbo._GetCOLSForExportXML @DATEFROM,@DATETO, @RESULTXML OUTPUT;
						SEND ON CONVERSATION @handle MESSAGE TYPE COLS_RS_MSG (@RESULTXML)
					END
					COMMIT TRANSACTION;
				END
--=======================================================================

				IF @messageTypeName = N''CUSTOMERSET_RQ_MSG''
				BEGIN
					INSERT INTO _tSBLog (SBL_Msg,SBL_Proc) VALUES (LEFT(''POS 1. CUSTOMERSET_RQ_MSG'',600),''[_actp_Invocation_RQ_Consumer]'')

					EXEC sp_xml_preparedocument @hdoc OUTPUT, @messageBody

					INSERT INTO @T(Acno)
						SELECT x.Acno
						FROM OPENXML ( @hdoc, ''/AcnoSelection/DetailLine'', 2 ) WITH (
							Acno VARCHAR(20) ''ACNO''
							) AS x;

					EXEC sp_xml_removedocument @hdoc
					INSERT INTO _tCQ (CQ_INT,CQ_TYPE,CQ_FieldsUpdated)
						SELECT TP_ID,''TP'',''All''
						FROM  @T t JOIN tTP tp ON t.Acno = tp.TP_Acno;
					INSERT INTO _tSBLog (SBL_Msg,SBL_Proc) 
					VALUES (LEFT(''POS 2. CUSTOMERSET_RQ_MSG Inserted rows:'' + CAST(@@ROWCOUNT as varchar(5)),600) ,''[_actp_Invocation_RQ_Consumer]'')

					EXECUTE _SendLoyaltyCustomers
					COMMIT TRANSACTION;
				END
				IF @messageTypeName = N''SOHALL_RQ_MSG''
				BEGIN
					SELECT @QtyOHBody =  
						(SELECT ISNULL(P_QTYONHAND,0) SOH,ISNULL(P_EAN,'''') EAN FROM tPRODUCT WHERE ISNULL(P_QTYONHAND,0) > 0  FOR XML AUTO, TYPE)
					SELECT @QtyOHBody = 
						''<SOHMsg>
							<List>'' + CAST(@QtyOHBody as VARCHAR(MAX)) + ''</List>
							<STCODE>'' + @INSTALLATIONCODE + ''</STCODE>
						</SOHMsg>'';
					SEND ON CONVERSATION @handle MESSAGE TYPE SOHALL_RS_MSG (@QtyOHBody);

					INSERT INTO _tSBLog (SBL_Msg,SBL_Proc) VALUES (''POS AfterSOH return message.'' ,''[_actp_Invocation_RQ_Consumer]'')
					COMMIT TRANSACTION;
				END
--=======================================================================
--				IF @messageTypeName = N''SALESSET_RQ_MSG''
--				BEGIN
--					EXEC sp_xml_preparedocument @hdoc OUTPUT, @messageBody
--
--					INSERT INTO @T(Acno)
--						SELECT x.Acno,CASE WHEN x.SELECTED IN (''-1'',''1'',''TRUE'') THEN 1 ELSE 0 END
--						FROM OPENXML ( @hdoc, ''/DetailLine/ACNO'', 3 ) WITH (
--							Acno VARCHAR(20) ''Acno''
--							) AS x
--
--					EXEC sp_xml_removedocument @hdoc
--
--					UPDATE tExchange set EXCH_SentToCentralAt = NULL WHERE EXCH_SALEDATE >= @Dte AND  EXCH_SALEDATE <= dbo.EndOfDay(@Dte)
--					EXEC _SendExchanges
--				END
--=======================================================================

				ELSE IF @messageTypeName = N''BUDGETDATA_MSG''
				BEGIN
					INSERT INTO _tSBLog (SBL_Msg,SBL_Proc) VALUES (''BUDGETDATA_MSG.'' ,''[_actp_Invocation_RQ_Consumer]'')

--DECLARE @hdoc INT  
--DECLARE  @messageBody XML
--DECLARE @BUDGETMONTH DATETIME
--DECLARE @DELIVERY INT
--DECLARE @RETURN INT
--SELECT @BUDGETMONTH = ''2010-03-01''
--SELECT @DELIVERY = 120
--SELECT @RETURN = 5
--		SELECT @messageBody = 
--			''<Msg>
--				<MSGTYPE>BUDGET_UPDATE</MSGTYPE>
--				<BUDGETMONTH>'' + CONVERT(VARCHAR(30),@BUDGETMONTH,120) + ''</BUDGETMONTH>
--				<DELIVERY>'' + CAST(@DELIVERY AS VARCHAR(40)) + ''</DELIVERY>
--				<RETURN>'' + CAST(@RETURN AS VARCHAR(40)) + ''</RETURN>
--			</Msg>''

					EXEC sp_xml_preparedocument @hdoc OUTPUT, @messageBody
DECLARE @MSGTYPE VARCHAR(50)
						SELECT @MSGTYPE = MSGTYPE
						FROM OPENXML ( @hdoc, ''/Msg'', 2 ) WITH (
							MSGTYPE VARCHAR(50) ''MSGTYPE''
							)

						If @MSGTYPE = ''BUDGET_UPDATE''
						BEGIN
							SELECT @MSGTYPE =''SUCESS''
						END
					EXEC sp_xml_removedocument @hdoc

--SELECT @MSGTYPE



					COMMIT TRANSACTION;
				END


				ELSE IF @messageTypeName = N''EndOfStream''
				BEGIN
					END CONVERSATION @handle;
					COMMIT TRANSACTION;
				END

				ELSE IF   @messageTypeName = N''http://schemas.microsoft.com/SQL/ServiceBroker/EndDialog''
				BEGIN
					END CONVERSATION @handle;
					COMMIT TRANSACTION;
				END

				ELSE IF @messageTypeName = N''http://schemas.microsoft.com/SQL/ServiceBroker/Error''
				BEGIN
					END CONVERSATION @handle;
					DECLARE @error INT;
					DECLARE @description NVARCHAR(4000);
					WITH XMLNAMESPACES (''http://schemas.microsoft.com/SQL/ServiceBroker/Error'' AS ssb)
					SELECT
						@error = CAST(@messageBody AS XML).value(''(//ssb:Error/ssb:Code)[1]'', ''INT''),
						@description = CAST(@messageBody AS XML).value(''(//ssb:Error/ssb:Description)[1]'', ''NVARCHAR(4000)'')
					RAISERROR(N''Received error Code:%i Description:''''%s'''''', 16, 1, @error, @description) WITH LOG;
					ROLLBACK TRANSACTION;
					INSERT INTO _tSBLog (SBL_Msg,SBL_Proc) VALUES (LEFT(''Pos 3:error '' + @description,600),''[_actp_Invocation_RQ_Consumer]'')
				END
				ELSE
				BEGIN
					COMMIT TRANSACTION;
					INSERT INTO _tSBLog (SBL_Msg,SBL_Proc) VALUES (''Pos 4:error '' + @messageTypeName ,''[_actp_Invocation_RQ_Consumer]'')
				END
			END  
			ELSE
			BEGIN
				COMMIT TRANSACTION;
				INSERT INTO _tSBLog(SBL_MSG,SBL_PROC) VALUES (''POS 5:Rollback: handle NULL @messageTypeName = '' + @messageTypeName ,''[_actp_Invocation_RQ_Consumer]'')
			END
			
	END TRY
	BEGIN CATCH
				if (XACT_STATE()) = -1
				BEGIN
					 rollback transaction;
					INSERT INTO _tSBLog(SBL_MSG,SBL_PROC) 
							VALUES (LEFT(''CATCH:ROLLBACK'' + @ErrorString  + '': '' + CAST(@messageBody AS VARCHAR(MAX)),580),''[_actp_Invocation_RQ_Consumer]'')
				END;
				-- Test whether the transaction is active and valid.
				if (XACT_STATE()) = 1
				BEGIN
					IF ERROR_PROCEDURE() <> OBJECT_NAME(@@PROCID)
						SELECT @ErrorString = ERROR_MESSAGE();
					ELSE
						SELECT @ErrorString = ''Error Message: '' + ERROR_MESSAGE() + '', Error Procedure: '' + 
										RTRIM(CONVERT(CHAR,ERROR_PROCEDURE())) + '', Error Line: '' + RTRIM(CONVERT(Char,ERROR_LINE ()));
					RAISERROR (@ErrorString, 16,1)

					--Update the log with error message
					INSERT INTO _tSBLog(SBL_MSG,SBL_PROC) 
						VALUES (LEFT(''CATCH:COMMIT'' + @ErrorString  + '': '' + CAST(@messageBody AS VARCHAR(MAX)),580),''[_actp_Invocation_RQ_Consumer]'')
					Commit transaction;	
				END
	END CATCH
END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo._actp_Invocation_RQ_Consumer Procedure Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo._actp_Invocation_RQ_Consumer Procedure'
END
GO

--
-- Script To Create dbo.BackupDatabases Procedure In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.BackupDatabases Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE PROCEDURE [dbo].[BackupDatabases] (@Path VARCHAR(200))
AS
BEGIN
		DECLARE @name VARCHAR(50) -- database name 
		DECLARE @fileName VARCHAR(256) -- filename for backup 
		DECLARE @fileDate VARCHAR(20) -- used for file name
		DECLARE @DELETEPATH VARCHAR(300)

		SELECT @DELETEPATH = ''DEL '' + ''"'' + @Path + ''*.BAK"''
		SELECT @DELETEPATH
		EXEC xp_cmdshell @DELETEPATH

		SELECT @fileDate = CONVERT(VARCHAR(20),GETDATE(),112)

		DECLARE db_cursor CURSOR FOR 
		SELECT name 
		FROM master.dbo.sysdatabases 
	--	WHERE name NOT IN (''master'',''model'',''msdb'',''tempdb'') 
		WHERE name NOT IN (''tempdb'',''PBKSBU'',''STS_Config'',''STS_SBSERVER_1'')

		OPEN db_cursor  
		FETCH NEXT FROM db_cursor INTO @name  

		WHILE @@FETCH_STATUS = 0  
		BEGIN  
			   SET @fileName = @path + @name + ''_'' + @fileDate + ''.BAK'' 
			   BACKUP DATABASE @name TO DISK = @fileName 

			   FETCH NEXT FROM db_cursor INTO @name  
		END  

		CLOSE db_cursor  
		DEALLOCATE db_cursor
END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.BackupDatabases Procedure Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.BackupDatabases Procedure'
END
GO

--
-- Script To Create dbo.Budget_Update_InDayend Procedure In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.Budget_Update_InDayend Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE PROCEDURE dbo.Budget_Update_InDayend (@Date Datetime)
AS
BEGIN
DECLARE @i INT
DECLARE @Dte DATETIME
	--Insert rows into tBudget for last 12 months and next 12 months and set to 1 if they don''t exist else leave alone
	SELECT @i = -12
	WHILE @i <= 12
	BEGIN
		SELECT @DTE = dbo.StartOfMonth(DATEADD(m,@i,GetDate()))
		If NOT EXISTS(SELECT * FROM tBUDGET WHERE B_BUDGETMONTH = @DTE)
		INSERT INTO tBUDGET (B_BUDGETMONTH,B_DeliveriesBudget,B_ReturnsBudget) VALUES (@DTE,1,1)
		SELECT @i = @i + 1
	END
	--update orders with ETAs for all deliveries from 5 months back
	UPDATE tBudget SET B_OrdersAtRetailValueIssued  = OrdersAtRetailValueIssued, B_OrdersAtRetailValueInProcess = OrdersAtRetailValueInProcess 
	FROM Budget_Orders 
	WHERE DELIVERYMONTH = B_BudgetMonth AND B_BUDGETMONTH > dbo.StartOfMonth(DateAdd(m,-5,@Date))

	--update Returns at processed date for all deliveries from 5 months back
	UPDATE tBudget SET B_RetailValueReturnsIssued  = RetailValueReturnsIssued, B_RetailValueReturnsInProcess = RetailValueReturnsInProcess 
	FROM Budget_Returns
	WHERE ReturnMonth = B_BudgetMonth AND B_BUDGETMONTH > dbo.StartOfMonth(DateAdd(m,-5,@Date))

	--update orders with ETAs for all deliveries from 5 months back
	UPDATE tBudget SET B_RetailValueReceivedIssued  = RetailValueReceivedIssued, B_RetailValueReceivedInProcess = RetailValueReceivedInProcess
	FROM Budget_Deliveries
	WHERE DELIVERYMONTH = B_BudgetMonth AND B_BUDGETMONTH > dbo.StartOfMonth(DateAdd(m,-5,@Date))

	--Update ratios
	UPDATE tBudget SET	B_OrdersAgainstBudget= (B_OrdersAtRetailValueIssued + B_OrdersAtRetailValueInProcess)/B_DeliveriesBudget, 
						B_DeliveriesAgainstBudget = (B_RetailValueReceivedIssued + B_RetailValueReceivedInProcess)/B_DeliveriesBudget
	WHERE B_BudgetMonth > DATEADD(m,-12,@Date)

	--Update four month folling average ratio
	UPDATE tBUDGET SET B_DeliveriesAgainstBudget_FourMonthAverage = (RetailValueReceivedIssued+RetailValueReceivedInProcess)/B_DeliveriesBudget 
		FROM BUDGET_LastFourMonths_1 WHERE DELIVERYMONTH = dbo.StartOfMonth(@Date)

END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.Budget_Update_InDayend Procedure Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.Budget_Update_InDayend Procedure'
END
GO

--
-- Script To Update dbo.CreateorAddtoCOFromXML Procedure In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.CreateorAddtoCOFromXML Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER PROCEDURE [dbo].[CreateorAddtoCOFromXML] @DOCXML XML
AS
BEGIN
	SET NOCOUNT ON;
DECLARE @VATRATE NUMERIC(6,2)
DECLARE @DEFAULTVATRATE NUMERIC(6,2)
DECLARE @COMPCODE VARCHAR(10)
DECLARE @COMPID INT
DECLARE @DOCCODE VARCHAR(20)
DECLARE @EAN VARCHAR(20)
DECLARE @hdoc INT   
DECLARE @PID UNIQUEIDENTIFIER
DECLARE @TFDOCCODE VARCHAR(15)
DECLARE @TFDOCDATE VARCHAR(25)
DECLARE @TFFROM VARCHAR(40)
DECLARE @NEWTRID INT
DECLARE @XML XML
DECLARE @CMD VARCHAR(200)
DECLARE @RES INT
DECLARE @ERRMESS VARCHAR(1000)
DECLARE @strSTAFFID VARCHAR(5)
DECLARE @TPID INT
DECLARE @TRID  INT
DECLARE @COL_PRICE INT
DECLARE @CO TABLE 
(
	StaffMember VARCHAR(15)  
)
DECLARE @COL TABLE 
(
	PID VARCHAR(50)
)
declare @$prog varchar(50), 
	@$errno int, 
	@$errmsg varchar(4000), 
	@$proc_section_nm varchar(50),
	@$row_cnt INT,
	@$error_db_name varchar(50), 
	@$CreateUserName varchar(128),   -- last user changed the data 
	@$CreateMachineName varchar(128), -- last machine changes-procedure were run from
	@$CreateSource varchar(128)		-- last process that made a changes

select @$errno = NULL,  @$errmsg = NULL,  @$proc_section_nm = NULL
	,  @$prog = LEFT(object_name(@@procid),50), @$row_cnt = NULL
	, @$error_db_name = db_name();


BEGIN TRY
----DECLARE @DOCXML XML
----SELECT @DOCXML = ''<doc_SpecialOrderAddition><MessageType>SpecialOrderAddition</MessageType><MessageCreationDate>201003190842</MessageCreationDate><StaffMember>2</StaffMember><DetailLines><ITEM><PID>{B63F82FE-62C2-477B-A242-0F792E923D2B}</PID><CodeF>978-0-07-159158-4*</CodeF><Description> GREENSPAN''''S BUBBLES - T</Description><Author></Author><Distributor></Distributor><Qtys/><Publisher></Publisher><LocalPrice>R.00</LocalPrice><EAN>9780071591584</EAN><Obsolete>False</Obsolete></ITEM></DetailLines></doc_SpecialOrderAddition>''


	    EXEC sp_xml_preparedocument @hdoc OUTPUT, @DOCXML
		INSERT INTO @CO( StaffMember)
			SELECT x.StaffMember
			FROM OPENXML ( @hdoc, ''/doc_SpecialOrderAddition'', 2 ) WITH (
				StaffMember VARCHAR(5)
				) AS x
		EXEC sp_xml_removedocument @hdoc

		EXEC sp_xml_preparedocument @hdoc OUTPUT, @DOCXML

		INSERT INTO @COL( PID)
			SELECT x.PID
			FROM OPENXML ( @hdoc, ''/doc_SpecialOrderAddition/DetailLines/ITEM'', 2 ) WITH (
				PID VARCHAR(50)
				) AS x
		EXEC sp_xml_removedocument @hdoc



	SELECT @strSTAFFID = StaffMember FROM @CO
	SELECT @TPID = ISNULL(TP_ID,0) FROM tTP WHERE TP_NAME = ''SPECIAL_ORDER_'' +  @strSTAFFID

	--CREATE Special customer record if not exists
	IF ISNULL(@TPID,0) = 0 
	BEGIN
		INSERT INTO tTP (TP_NAME,TP_NOTE,TP_ROLE) VALUES (''SPECIAL_ORDER_'' +  @strSTAFFID,''Special customer account for supervisor staff'',3)
		SELECT @TPID = SCOPE_IDENTITY()
	END

	SELECT @TRID = ISNULL(TR_ID,0) FROM tTR JOIN tCO ON TR_ID = CO_ID JOIN tTP ON TR_TP_ID = TP_ID WHERE TP_ID = @TPID AND TR_STATUS = 2
	SELECT @DEFAULTVATRATE = CF_VATRATE,@COMPID =  CF_DEFAULTCOMPANYID,@COMPCODE = COMP_CODE
	FROM tCOMPANY JOIN tCONFIGURATION ON CF_DEFAULTCOMPANYID = COMP_ID

	IF ISNULL(@TRID,0) = 0
	BEGIN
		EXEC sp_GetNextCode 19,@DOCCode OUTPUT
		SELECT @DOCCode = ISNULL(@COMPCODE,'''') + ''CSP'' + @DOCCode

		INSERT INTO tTR (TR_COMP_ID,TR_CODE,TR_DATE,TR_CAPTUREDATE,TR_TP_ID,TR_STATUS,TR_TYPE,TR_NOTE) 
		VALUES (@COMPID,@DOCCode,GETDATE(),GetDate(),@TPID,2,1, ''Special supervisor''''s order'')

		SELECT @TRID = SCOPE_IDENTITY()

		INSERT INTO tCO (CO_ID,CO_OrderType) 
		VALUES (@TRID,1)
	END

	DECLARE cur CURSOR FOR
	SELECT PID FROM @COL
	OPEN cur
	FETCH NEXT FROM cur INTO @PID
	WHILE @@FETCH_STATUS = 0
	BEGIN
		SELECT  @VATRATE = P_VATRATE,@COL_PRICE = P_RRP FROM tPRODUCT WHERE P_ID = @PID
		INSERT INTO tCOL (COL_TR_ID,COL_P_ID,COL_QTY,COL_QTYFIRM,COL_FULFILLED,COL_PRICE,COL_REF) 
			VALUES	(@TRID,@PID,1,1,''OS'',@COL_PRICE,''Store order'')
		FETCH NEXT FROM cur INTO @PID
	END
	CLOSE cur
	DEALLOCATE cur




END TRY
BEGIN CATCH
set @$errmsg = Left(''Error '' +
		CASE
			WHEN @$errno > 0 THEN CAST(@$errno as varchar)
			ELSE Cast(ERROR_NUMBER() as varchar)
		END + ''in proc '' + isnull(@$prog,'' '') + '' '' + 
		CASE 
			WHEN @$errno > 0 THEN isnull(@$errmsg,'' '') 
			ELSE isnull(@$errmsg,'' '') + ISNULL(ERROR_MESSAGE(),'''')
		END ,4000);

raiserror (@$errmsg, 16, 1); 

--EXEC dbo.ERROR_LOG_2005 @ERROR_LOG_PROGRAM_NM  = @$prog,  
--		@ERROR_LOG_PROGRAM_SECTION_NM  = @$proc_section_nm,
--		@ERROR_LOG_ERROR_NO  = @$errno,  
--		@ERROR_LOG_ERROR_DSC  = @$errmsg,
--		@ERROR_DB_NAME  = @$error_db_name
		-- set the error if not set
		declare @t VARCHAR(1000)
		select @T = @$prog + ''/'' + @$proc_section_nm
		EXEC dbo.SAVELog @$errmsg ,@t,NULL,NULL






--	IF @@TRANCOUNT > 0
--		ROLLBACK TRAN
--	DECLARE @ErrorString NVARCHAR(4000)
--	IF ERROR_PROCEDURE() <> OBJECT_NAME(@@PROCID)
--		SELECT @ErrorString = ERROR_MESSAGE();
--	ELSE
--		SELECT @ErrorString = ''Error Message: '' + ERROR_MESSAGE() + '', Error Procedure: '' + RTRIM(CONVERT(CHAR,ERROR_PROCEDURE())) + '', Error Line: '' + RTRIM(CONVERT(Char,ERROR_LINE ()));
--	RAISERROR (@ErrorString, 16,1)
--	--Update the log with error message
--	INSERT INTO _tSBLog (SBL_Msg,SBL_Proc) VALUES (LEFT(''Pos 2: '' + LEFT(@ErrorString,500),600),''trigQtyOHChange'')
END CATCH











END




RETURN')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.CreateorAddtoCOFromXML Procedure Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.CreateorAddtoCOFromXML Procedure'
END
GO

--
-- Script To Update dbo.ExportCreditorsTrading_Pastel Procedure In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.ExportCreditorsTrading_Pastel Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER PROCEDURE [dbo].[ExportCreditorsTrading_Pastel](@SMID INTEGER,@LASTTRIDEXPORTED INT,
						@SINCEDATE DATETIME = NULL,@PERIOD INTEGER,@RES INTEGER OUTPUT) AS
BEGIN
DECLARE @TRID INTEGER
DECLARE @TRTYPE INTEGER
DECLARE @TRDATE DATETIME
DECLARE @TRPROCESSINGDATE DATETIME
DECLARE @TPACNO VARCHAR(15)
DECLARE @TRDOCNO VARCHAR(15)
DECLARE @AMT NUMERIC(18,2)
DECLARE @VAT NUMERIC(18,2)
DECLARE @PT_PURCHASES_CONTRA VARCHAR(30)
DECLARE @OPID as INTEGER
DECLARE @UsePapyrusExchangeRateWhenExportingToAccounting BIT
DECLARE @LASTTRID INTEGER
DECLARE @FIRSTTRCODE VARCHAR(20)
DECLARE @LASTTRCODE VARCHAR(20)
DECLARE @RESULT VARCHAR(200)
DECLARE @PREVTRID INTEGER
DECLARE @TRSTATUS INT
DECLARE @EXCHANGERATE NUMERIC(11,7)
DECLARE @ERR INTEGER
DECLARE @VATABLE INT
DECLARE @ISFOREIGN BIT
DECLARE @NS_AMT NUMERIC(18,2)
DECLARE @IsNonStock BIT
DECLARE @CURRENTTRID INT
DECLARE @SUPPLIERINVOICEREF VARCHAR(50)
DECLARE @SUPPLIERINVOICEDATE DATETIME
DECLARE @UseSuppliersInvoiceDateOnPosting VARCHAR(5)
DECLARE @INVENTORYACCOUNTINGMODEL VARCHAR(20)
DECLARE @ICA VARCHAR(20)
DECLARE @COSA VARCHAR(20)
DECLARE @SA VARCHAR(20)
DECLARE @EXPORTDATE DATETIME

	BEGIN TRANSACTION
		DELETE FROM tPASTEL
		SELECT @EXPORTDATE = GETDATE()
		SELECT @UsePapyrusExchangeRateWhenExportingToAccounting = dbo.GetProperty(''UsePapyrusExchangeRateWhenExportingToAccounting'')
		INSERT INTO tOPERATION (OP_TYPE,OP_STARTEDAT,OP_ENDEDAT,OP_STARTEDBYID,OP_RESULT,OP_FULLREPORT) 
			VALUES (14,GetDate(),GetDate(), @SMID,1,@RESULT)
		SELECT @OPID = SCOPE_IDENTITY()

		SELECT @UseSuppliersInvoiceDateOnPosting = dbo.GetProperty(''UseSuppliersInvoiceDateOnPosting'')
		SELECT @INVENTORYACCOUNTINGMODEL = dbo.GetProperty(''InventoryAccountingModel'')
		SELECT @ICA = dbo.GetProperty(''InventoryControlAccount'')
		SELECT @COSA = dbo.GetProperty(''CostOfSalesAccount'')
		SELECT @SA = dbo.GetProperty(''SalesAccount'')

		DECLARE curPASTEL CURSOR FOR
		SELECT TR_ID,TR_TYPE,TR_DATE,TP_ACNO,TR_CODE,TR_STATUS FROM tTR LEFT JOIN tTP ON TR_TP_ID = TP_ID 
			WHERE TR_PROCESSINGDATE >= @SINCEDATE  AND TR_TYPE IN (4,11) AND TR_STATUS IN (3,4)  ORDER BY TR_PROCESSINGDATE

		OPEN curPASTEL
		FETCH NEXT FROM curPASTEL INTO @TRID,@TRTYPE,@TRDATE,@TPACNO,@TRDOCNO,@TRSTATUS
		SELECT @FIRSTTRCODE = @TRDOCNO
		WHILE @@FETCH_STATUS = 0
		BEGIN
			IF @TRTYPE =4    --GRN
			BEGIN
				DECLARE CurITEM CURSOR FOR
					SELECT	TR_PROCESSINGDATE,
							CASE WHEN P_PRODUCTTYPE IN (''B'',''G'',''M'') THEN 0 ELSE 1 END as NonStock, 
							CASE WHEN ISLOCAL = 0 THEN NS_FOREIGN ELSE NS_LOCAL END,
							AMT,CASE WHEN P_PRODUCTTYPE not IN (''B'',''G'',''M'')THEN 0 ELSE VAT END,PT_PURCHASES_CONTRA ,ISNULL(CONVERTTOLOCAL,1) ,
							TP_VATABLE,TR_ID,DEL_SUPPLIERINVOICEREF,DEL_SUPPLIERINVOICEDATE
					FROM vExportCreditorsInvoicesForAccounting_PASTEL WHERE TR_ID = @TRID
				OPEN curITEM
				FETCH NEXT FROM CurITEM INTO @TRPROCESSINGDATE, @ISNONSTOCK,@NS_AMT,@AMT,@VAT,
							@PT_PURCHASES_CONTRA,@EXCHANGERATE,@VATABLE,@CURRENTTRID,
							@SUPPLIERINVOICEREF,@SUPPLIERINVOICEDATE
				WHILE @@FETCH_STATUS = 0 AND @CURRENTTRID = @TRID
				BEGIN
					INSERT INTO tPASTEL (PERIOD,DTE,GDC,ACNO,REFERENCE,
							DESCR,AMT,TaxType,TAXAMT,OPENITEM,COSTCODE,
							CONTRAACCOUNT,EXCHANGERATE,BANKEXCHANGERATE,BATCHID,DISCOUNTTAX,DISCOUNTAMT,HOMEAMT,[ACTION],ProcessingDate) 
						VALUES (@PERIOD,
							CASE WHEN @UseSuppliersInvoiceDateOnPosting = ''TRUE'' THEN @SUPPLIERINVOICEDATE ELSE @TRDATE END,
							''C'',@TPACNO,RIGHT(@SUPPLIERINVOICEREF,8),''P-''+ @TRDOCNO,
							CASE WHEN @ISNONSTOCK = 0 THEN @AMT * -1 ELSE @NS_AMT * -1 END,
							CASE WHEN @EXCHANGERATE = 1 THEN 1 ELSE 0 END,  --this determines the tax indicator that goes to Pastel (1 taxable,0 not)
							@VAT * -1,'' '',''     '',
							CASE WHEN @INVENTORYACCOUNTINGMODEL = ''PERPETUAL'' THEN @ICA ELSE @PT_PURCHASES_CONTRA END,
							CASE WHEN @UsePapyrusExchangeRateWhenExportingToAccounting = 1 THEN @EXCHANGERATE ELSE 0 END,
							CASE WHEN @UsePapyrusExchangeRateWhenExportingToAccounting = 1 THEN @EXCHANGERATE ELSE 0 END,
							0,0,0,0,1,@TRPROCESSINGDATE)
					FETCH NEXT FROM CurITEM INTO @TRPROCESSINGDATE,@ISNONSTOCK,@NS_AMT,@AMT,@VAT,@PT_PURCHASES_CONTRA,
						@EXCHANGERATE,@VATABLE,@CURRENTTRID,@SUPPLIERINVOICEREF,@SUPPLIERINVOICEDATE
				END
				CLOSE CurITEM
				DEALLOCATE CurITEM
				SELECT @ERR = @@ERROR
				IF @ERR <> 0 GOTO LEAVE
			END
			IF @TRTYPE =11  --Return to suppliers
			BEGIN
				DECLARE CurITEM CURSOR FOR
					SELECT	TR_PROCESSINGDATE,
							CASE WHEN P_PRODUCTTYPE IN (''B'',''G'',''M'') THEN 0 ELSE 1 END as NonStock, 
							CASE WHEN ISLOCAL = 0 THEN NS_FOREIGN ELSE NS_LOCAL END,
							AMT,CASE WHEN P_PRODUCTTYPE not IN (''B'',''G'',''M'')THEN 0 ELSE VAT END,PT_PURCHASES_CONTRA ,ISNULL(CONVERTTOLOCAL,1) ,
							TP_VATABLE,TR_ID,DEL_SUPPLIERINVOICEREF,DEL_SUPPLIERINVOICEDATE
					FROM vExportCreditorsReturnsForAccounting_PASTEL WHERE TR_ID = @TRID
				OPEN curITEM
				FETCH NEXT FROM CurITEM INTO @TRPROCESSINGDATE,@ISNONSTOCK,@NS_AMT,@AMT,@VAT,@PT_PURCHASES_CONTRA,
								@EXCHANGERATE,@VATABLE,@CURRENTTRID,@SUPPLIERINVOICEREF,@SUPPLIERINVOICEDATE
				WHILE @@FETCH_STATUS = 0 AND @CURRENTTRID = @TRID
				BEGIN
					INSERT INTO tPASTEL (PERIOD,DTE,GDC,ACNO,REFERENCE,DESCR,
					AMT,
					TaxType,
					TAXAMT,OPENITEM,COSTCODE,
					CONTRAACCOUNT,EXCHANGERATE,BANKEXCHANGERATE,BATCHID,DISCOUNTTAX,DISCOUNTAMT,HOMEAMT,[ACTION],ProcessingDate) 
						VALUES (@PERIOD,@TRDATE,''C'',@TPACNO,RIGHT(@SUPPLIERINVOICEREF,8),''R-'' + @TRDOCNO,
							CASE WHEN @ISNONSTOCK = 0 THEN @AMT ELSE @NS_AMT END,
							CASE WHEN @EXCHANGERATE = 1 THEN 1 ELSE 0 END,  --this determines the tax indicator that goes to Pastel (1 taxable,0 not)
							@VAT * -1,'' '',''     '',
							CASE WHEN @INVENTORYACCOUNTINGMODEL = ''PERPETUAL'' THEN @ICA ELSE @PT_PURCHASES_CONTRA END,
							CASE WHEN @UsePapyrusExchangeRateWhenExportingToAccounting = 1 THEN @EXCHANGERATE ELSE 0 END,
							CASE WHEN @UsePapyrusExchangeRateWhenExportingToAccounting = 1 THEN @EXCHANGERATE ELSE 0 END,
							0,0,0,0,1,@TRPROCESSINGDATE)
					FETCH NEXT FROM CurITEM INTO @TRPROCESSINGDATE,@ISNONSTOCK,@NS_AMT,@AMT,@VAT,@PT_PURCHASES_CONTRA,
						@EXCHANGERATE,@VATABLE,@CURRENTTRID,@SUPPLIERINVOICEREF,@SUPPLIERINVOICEDATE
				END
				CLOSE CurITEM
				DEALLOCATE CurITEM

				SELECT @ERR = @@ERROR
				IF @ERR <> 0 GOTO LEAVE
			END
		
			SELECT @PREVTRID = @TRID
			SELECT @LASTTRCODE =  @TRDOCNO
			FETCH NEXT FROM curPASTEL INTO @TRID,@TRTYPE,@TRDATE,@TPACNO,@TRDOCNO,@TRSTATUS
		
		END

		CLOSE curPASTEL
		DEALLOCATE curPASTEL

		UPDATE tOPERATION SET OP_FULLREPORT = ''Export Creditors invoices and returns to PASTEL '' + @FIRSTTRCODE + '' to '' + @LASTTRCODE 
				WHERE OP_ID = @OPID
		IF ISNULL(@PREVTRID,0) > 0 
			UPDATE tCONFIGURATION SET CF_Last_TR_Exported = @PREVTRID

		UPDATE tIECONTROL SET IEC_LastTRID = @PREVTRID WHERE IEC_NAME = ''EXPORTCREDITORSTRADING''
	COMMIT TRANSACTION
		SELECT @RES =0
		RETURN
	LEAVE:
		ROLLBACK TRANSACTION
		SELECT @RES = @ERR
		RETURN

END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.ExportCreditorsTrading_Pastel Procedure Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.ExportCreditorsTrading_Pastel Procedure'
END
GO

--
-- Script To Delete dbo.ExportCreditorsTrading_PastelMaster Procedure In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Deleting dbo.ExportCreditorsTrading_PastelMaster Procedure'
GO

   DROP PROCEDURE [dbo].[ExportCreditorsTrading_PastelMaster]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.ExportCreditorsTrading_PastelMaster Procedure Deleted Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Delete dbo.ExportCreditorsTrading_PastelMaster Procedure'
END
GO

--
-- Script To Create dbo.ExportCustomers_Pastel Procedure In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.ExportCustomers_Pastel Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


SET ANSI_NULLS OFF
GO

exec('CREATE PROCEDURE dbo.ExportCustomers_Pastel(@SMID INTEGER,@SINCEDATE DATETIME = NULL,@RES INTEGER OUTPUT) AS

DECLARE @TRID INTEGER
DECLARE @TRTYPE INTEGER
DECLARE @TRDATE DATETIME
DECLARE @TPACNO VARCHAR(15)
DECLARE @TRDOCNO VARCHAR(15)
DECLARE @AMT NUMERIC(15,2)
DECLARE @VAT NUMERIC(15,2)

DECLARE @OPID as INTEGER

DECLARE @LASTTRID INTEGER
DECLARE @FIRSTTRCODE VARCHAR(20)
DECLARE @LASTTRCODE VARCHAR(20)
DECLARE @RESULT VARCHAR(200)
DECLARE @PREVTRID INTEGER

DECLARE @ERR INTEGER

BEGIN TRANSACTION

	DELETE FROM tPASTEL_CUST_EXPORT

	
	INSERT INTO tOPERATION (OP_TYPE,OP_STARTEDAT,OP_ENDEDAT,OP_STARTEDBYID,OP_RESULT,OP_FULLREPORT) 
		VALUES (15,GetDate(),GetDate(), @SMID,1,@RESULT)
	SELECT @OPID = SCOPE_IDENTITY()

INSERT INTO tPASTEL_CUST_EXPORT
   SELECT Left(LTRIM(TP_ACNO),6), Left(dbo.FullName2F(LTRIM(TP_TITLE) , LTRIM(TP_INITIALS), LTRIM(TP_NAME)),40), Left(LTRIM(add1),30),Left(LTRIM(add2),30),Left(LTRIM(add3),30),Left(LTRIM(add4),30),Left(LTRIM(add5),30),
	Left(LTRIM(tel),16),Left(LTRIM(fax),16),Left(LTRIM(contactperson),16),Left(LTRIM(dadd1),30),Left(LTRIM(dadd2),30),Left(LTRIM(dadd3),30),Left(LTRIM(dadd4),30),Left(LTRIM(dadd5),30),
	TP_BLOCKED,LEFT(CAST(discount as VARCHAR(5)),5),LEFT(LTRIM(mobilephone),16),LEFT(LTRIM(email),50),LEFT(LTRIM(TAXREFERENCE),50),1
   FROM vExportCustomersForAccounting_PASTEL
  WHERE ISNULL(TP_DATELASTMODIFIED,''1900-01-01'') >= @SINCEDATE


	UPDATE tOPERATION SET OP_FULLREPORT = ''Export customers to PASTEL ''  WHERE OP_ID = @OPID


COMMIT TRANSACTION
	SELECT @RES =0
	RETURN
LEAVE:
	ROLLBACK TRANSACTION
	SELECT @RES = @ERR

	RETURN')
GO

SET ANSI_NULLS ON
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.ExportCustomers_Pastel Procedure Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.ExportCustomers_Pastel Procedure'
END
GO

--
-- Script To Create dbo.GetAllTableSizes Procedure In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.GetAllTableSizes Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE PROCEDURE [dbo].[GetAllTableSizes]
AS
/*
    Obtains spaced used data for ALL user tables in the database
*/
DECLARE @TableName VARCHAR(100)    --For storing values in the cursor

--Cursor to get the name of all user tables from the sysobjects listing
DECLARE tableCursor CURSOR
FOR 
select [name]
from dbo.sysobjects 
where  OBJECTPROPERTY(id, N''IsUserTable'') = 1
FOR READ ONLY

--A procedure level temp table to store the results
CREATE TABLE #TempTable
(
    tableName varchar(100),
    numberofRows varchar(100),
    reservedSize varchar(50),
    dataSize varchar(50),
    indexSize varchar(50),
    unusedSize varchar(50)
)

--Open the cursor
OPEN tableCursor

--Get the first table name from the cursor
FETCH NEXT FROM tableCursor INTO @TableName

--Loop until the cursor was not able to fetch
WHILE (@@Fetch_Status >= 0)
BEGIN
    --Dump the results of the sp_spaceused query to the temp table
    INSERT  #TempTable
        EXEC sp_spaceused @TableName

    --Get the next table name
    FETCH NEXT FROM tableCursor INTO @TableName
END

--Get rid of the cursor
CLOSE tableCursor
DEALLOCATE tableCursor

--Select all records so we can use the reults
SELECT * 
FROM #TempTable  ORDER BY CAST(numberofRows AS INT) DESC

--Final cleanup!
DROP TABLE #TempTable')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.GetAllTableSizes Procedure Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.GetAllTableSizes Procedure'
END
GO

--
-- Script To Update dbo.InitializeProperties Procedure In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.InitializeProperties Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER PROCEDURE [dbo].[InitializeProperties]

AS


--1	Nielsen
--2	FTP settings (support)
--3	Supports
--4	Dayend procedure
--5	User interface control
--6	Local defaults
--7	Local settings
--8	Accounting settings
--9	Central
--10	Email
--11	Statements
--12	Service Broker

BEGIN

SELECT [PROPT_ID] FROM tPropertyType WHERE PROPT_ID = 1
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tPropertyType]([PROPT_ID], [PROPT_DESCRIPTION]) 
		VALUES (1, ''Nielsen'')
ELSE
	UPDATE [dbo].tPropertyType SET [PROPT_DESCRIPTION] = ''Nielsen'' 
		WHERE PROPT_ID = 1
SELECT [PROPT_ID] FROM tPropertyType WHERE PROPT_ID = 2
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tPropertyType]([PROPT_ID], [PROPT_DESCRIPTION]) 
		VALUES (2, ''FTP settings (support)'')
ELSE
	UPDATE [dbo].tPropertyType SET [PROPT_DESCRIPTION] = ''FTP settings (support)'' 
		WHERE PROPT_ID = 2
SELECT [PROPT_ID] FROM tPropertyType WHERE PROPT_ID = 3
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tPropertyType]([PROPT_ID], [PROPT_DESCRIPTION]) 
		VALUES (3, ''Supports'')
ELSE
	UPDATE [dbo].tPropertyType SET [PROPT_DESCRIPTION] = ''Supports'' 
		WHERE PROPT_ID = 3
SELECT [PROPT_ID] FROM tPropertyType WHERE PROPT_ID = 4
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tPropertyType]([PROPT_ID], [PROPT_DESCRIPTION]) 
		VALUES (4, ''Dayend procedure'')
ELSE
	UPDATE [dbo].tPropertyType SET [PROPT_DESCRIPTION] = ''Dayend procedure'' 
		WHERE PROPT_ID = 4
SELECT [PROPT_ID] FROM tPropertyType WHERE PROPT_ID = 5
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tPropertyType]([PROPT_ID], [PROPT_DESCRIPTION]) 
		VALUES (5, ''User interface control'')
ELSE
	UPDATE [dbo].tPropertyType SET [PROPT_DESCRIPTION] = ''User interface control'' 
		WHERE PROPT_ID = 5
SELECT [PROPT_ID] FROM tPropertyType WHERE PROPT_ID = 6
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tPropertyType]([PROPT_ID], [PROPT_DESCRIPTION]) 
		VALUES (6, ''Local defaults'')
ELSE
	UPDATE [dbo].tPropertyType SET [PROPT_DESCRIPTION] = ''Local defaults'' 
		WHERE PROPT_ID = 6
SELECT [PROPT_ID] FROM tPropertyType WHERE PROPT_ID = 7
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tPropertyType]([PROPT_ID], [PROPT_DESCRIPTION]) 
	VALUES (7, ''Local settings'')
ELSE
	UPDATE [dbo].tPropertyType SET [PROPT_DESCRIPTION] = ''Local settings'' 
		WHERE PROPT_ID = 7
SELECT [PROPT_ID] FROM tPropertyType WHERE PROPT_ID = 8
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tPropertyType]([PROPT_ID], [PROPT_DESCRIPTION]) 
		VALUES (8, ''Accounting settings'')
ELSE
	UPDATE [dbo].tPropertyType SET [PROPT_DESCRIPTION] = ''Accounting settings'' 
		WHERE PROPT_ID = 8
SELECT [PROPT_ID] FROM tPropertyType WHERE PROPT_ID = 9
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tPropertyType]([PROPT_ID], [PROPT_DESCRIPTION]) 
		VALUES (9, ''Central'')
ELSE
	UPDATE [dbo].tPropertyType SET [PROPT_DESCRIPTION] = ''Central'' 
		WHERE PROPT_ID = 9
SELECT [PROPT_ID] FROM tPropertyType WHERE PROPT_ID = 10
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tPropertyType]([PROPT_ID], [PROPT_DESCRIPTION]) 
		VALUES (10, ''Email'')
ELSE
	UPDATE [dbo].tPropertyType SET [PROPT_DESCRIPTION] = ''Email'' 
		WHERE PROPT_ID = 10

SELECT [PROPT_ID] FROM tPropertyType WHERE PROPT_ID = 11
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tPropertyType]([PROPT_ID], [PROPT_DESCRIPTION]) 
		VALUES (11, ''Statements'')
ELSE
	UPDATE [dbo].tPropertyType SET [PROPT_DESCRIPTION] = ''Statements'' 
		WHERE PROPT_ID = 11

SELECT [PROPT_ID] FROM tPropertyType WHERE PROPT_ID = 12
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tPropertyType]([PROPT_ID], [PROPT_DESCRIPTION]) 
		VALUES (12, ''Service Broker'')
ELSE
	UPDATE [dbo].tPropertyType SET [PROPT_DESCRIPTION] = ''Service Broker'' 
		WHERE PROPT_ID = 12

SELECT [PROPT_ID] FROM tPropertyType WHERE PROPT_ID = 13
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tPropertyType]([PROPT_ID], [PROPT_DESCRIPTION]) 
		VALUES (13, ''Paths'')
ELSE
	UPDATE [dbo].tPropertyType SET [PROPT_DESCRIPTION] = ''Paths'' 
		WHERE PROPT_ID = 13




SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''ADMINISTRATOREMAIL''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) VALUES (''ADMINISTRATOREMAIL'', '''', ''not used'', NULL,''L'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''not used''   ,[PropertyTypeID] = 7,[Scope] = ''L''
		WHERE [PropertyKey] = ''ADMINISTRATOREMAIL''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''ALLOWANTIQUARIANSEARCH''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''ALLOWANTIQUARIANSEARCH'', ''1'', ''Allow the operator to search antiquarian books (books with individual copy records)'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Allow the operator to search antiquarian books (books with individual copy records)''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''ALLOWANTIQUARIANSEARCH''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''AllowGeneralStock''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''AllowGeneralStock'', ''TRUE'', ''Display menus for general stock as well as default (book stock)'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Display menus for general stock as well as default (book stock)''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''AllowGeneralStock''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''AllowInvoiceDateOverride''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''AllowInvoiceDateOverride'', ''FALSE'', ''Allow operator to override the nominal date of an invoice when issuing it. If false, then it takes the date it is issued.'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Allow operator to override the nominal date of an invoice when issuing it. If false, then it takes the date it is issued.''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''AllowInvoiceDateOverride''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''AllowPODateOverride''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''AllowPODateOverride'', ''FALSE'', ''Allow operator to override the nominal date of an purchase order when issuing it. If false, then it takes the date it is issued.'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Allow operator to override the nominal date of an purchase order when issuing it. If false, then it takes the date it is issued.''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''AllowPODateOverride''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''AllowsInvoicePicking''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) VALUES (''AllowsInvoicePicking'', ''FALSE'', ''If TRUE, additional status of ''''Picking'''' is supported'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, additional status of ''''Picking'''' is supported''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''AllowsInvoicePicking''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''AllowsSSInvoicing''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) VALUES (''AllowsSSInvoicing'', ''FALSE'', ''If TRUE, seesafe invoicing is supported'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, seesafe invoicing is supported''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''AllowsSSInvoicing''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''AllowsZeroDiscountPOs''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''AllowsZeroDiscountPOs'', ''TRUE'', ''Allows purchase orders to be issued without a specified deal - therefore with zero discount (affects report of value of o/s orders)'', 
				NULL,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Allows purchase orders to be issued without a specified deal - therefore with zero discount (affects report of value of o/s orders)'' ,
		[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''AllowsZeroDiscountPOs''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''AllowZeropricedPOLines''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''AllowZeropricedPOLines'', ''FALSE'', ''Allow operator to capture purchase order lines with prices of zero.'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Allow operator to capture purchase order lines with prices of zero.''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''AllowZeropricedPOLines''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''ApplyExtraChargesToCostType''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''ApplyExtraChargesToCostType'', '''', ''IF BY_QTY, then extra charges on the delivery transaction are allocated to individual items cost by averaging the extra charge'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''IF BY_QTY, then extra charges on the delivery transaction are allocated to individual items cost by averaging the extra charge''  ,
		[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''ApplyExtraChargesToCostType''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''BACKUP''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''BACKUP'', ''REMOVABLE'', ''Share name of the folder on the removable device to which the backup is written (e.g. A ZIP drive). Not applicable if backing up to CD. Or else a fixed path to a folder'', 4,''L'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Share name of the folder on the removable device to which the backup is written (e.g. A ZIP drive). Not applicable if backing up to CD. Or else a fixed path to a folder''  ,[PropertyTypeID] = 4,[Scope] = ''L''
		WHERE [PropertyKey] = ''BACKUP''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''BACKUPCOMPRESSION''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''BACKUPCOMPRESSION'', ''FALSE'', ''If TRUE, compresses the .BAK file by zipping it.'', 4,''L'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, compresses the .BAK file by zipping it.''  ,[PropertyTypeID] = 4,[Scope] = ''L''
		WHERE [PropertyKey] = ''BACKUPCOMPRESSION''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''BACKUPMEDIUM''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''BACKUPMEDIUM'', ''DISK'', ''Can be DISK or CD, depending whether the target drive is a CD or a disk'', 4,''L'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Can be DISK or CD, depending whether the target drive is a CD or a disk''  ,[PropertyTypeID] = 4,[Scope] = ''L''
		WHERE [PropertyKey] = ''BACKUPMEDIUM''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''BlindCashup''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''BlindCashup'', ''FALSE'', ''If TRUE then cashup is captured and locked before totals are displayed by computer.'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE then cashup is captured and locked before totals are displayed by computer.''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''BlindCashup''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''BOOKFINDFACET''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''BOOKFINDFACET'', ''WEBK'', ''WEBK for Compact World product and PMBK for Premier product'', 1,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''WEBK for Compact World product and PMBK for Premier product''  ,[PropertyTypeID] = 1,[Scope] = ''C''
		WHERE [PropertyKey] = ''BOOKFINDFACET''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''BOOKFINDISBN13ENABLED''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''BOOKFINDISBN13ENABLED'', ''TRUE'', ''Support searching on Nielsen in Papyrus using the 13 digit ISBN number. (not all versions of Nielsen support it as at June 2007)'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Support searching on Nielsen in Papyrus using the 13 digit ISBN number. (not all versions of Nielsen support it as at June 2007)''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''BOOKFINDISBN13ENABLED''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''BOOKFINDROOT''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''BOOKFINDROOT'', ''C:\BOOKFIND'', ''Folder into which the Nielsen product is installed'', 1,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Folder into which the Nielsen product is installed''  ,[PropertyTypeID] = 1,[Scope] = ''C''
		WHERE [PropertyKey] = ''BOOKFINDROOT''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CanEditCOs''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''CanEditCOs'', ''FALSE'', ''If TRUE, operator can edit issued customer orders'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, operator can edit issued customer orders''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''CanEditCOs''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CanEditPOs''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''CanEditPOs'', ''FALSE'', ''If TRUE, operator can edit issued purchase orders'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, operator can edit issued purchase orders''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''CanEditPOs''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CanEditQUs''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''CanEditQUs'', ''FALSE'', ''If TRUE, operator can edit issued quotations'', 7,''L'')
ELSE
	UPDATE [dbo].[tProperty] 
SET [PropertyDescription] = ''If TRUE, operator can edit issued quotations''  ,[PropertyTypeID] = 7,[Scope] = ''L''
		WHERE [PropertyKey] = ''CanEditQUs''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CashDrawerKick''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''CashDrawerKick'', ''7'', ''Signal to send to drawer to open it; e.g. 7,10,13 would be chr(7) + chr(10 + chr(13) '', 7,''L'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Signal to send to drawer to open it; e.g. 7,10,13 would be chr(7) + chr(10 + chr(13) '' ,[PropertyTypeID] = 7,[Scope] = ''L''
		WHERE [PropertyKey] = ''CashDrawerKick''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''Cashup_extended''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''Cashup_extended'', ''FALSE'', ''If TRUE, the alternative cashup sheet is used '', 3,''L'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE, the alternative cashup sheet is used ''  ,[PropertyTypeID] = 3,[Scope] = ''L''
		WHERE [PropertyKey] = ''Cashup_extended''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CDTYPE''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''CDTYPE'', ''RO'', ''RO for read-only CD and RW for rewritable CD'', 4,''L'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''RO for read-only CD and RW for rewritable CD''  ,[PropertyTypeID] = 4,[Scope] = ''L''
		WHERE [PropertyKey] = ''CDTYPE''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CENTRALFTPADDRESS''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''CENTRALFTPADDRESS'', ''ftp.whitaker.co.uk'', ''FTP folder on Central site from which stock is shared and loyalty files uploaded'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''FTP folder on Central site from which stock is shared and loyalty files uploaded''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''CENTRALFTPADDRESS''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CENTRALFTPFOLDER''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''CENTRALFTPFOLDER'', ''/bt000SA1/temp/CENTRAL'', ''Default folder for FTP folder on Central site from which stock is shared and loyalty files uploaded'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET	[PropertyDescription] = ''Default folder for FTP folder on Central site from which stock is shared and loyalty files uploaded''  ,
									[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''CENTRALFTPFOLDER''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CENTRALFTPPASSWORD''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''CENTRALFTPPASSWORD'', ''1beach'', ''Password for FTP folder on Central site from which stock is shared and loyalty files uploaded'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET	[PropertyDescription] = ''Password for FTP folder on Central site from which stock is shared and loyalty files uploaded''  ,
									[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''CENTRALFTPPASSWORD''


SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CENTRALFTPUSERNAME''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''CENTRALFTPUSERNAME'', ''bt000SA1'', ''User name for FTP folder on Central site from which stock is shared and loyalty files uploaded'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET	[PropertyDescription] = ''User name for FTP folder on Central site from which stock is shared and loyalty files uploaded''  ,
									[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''CENTRALFTPUSERNAME''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CheckRefsOnCO''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''CheckRefsOnCO'', '''', ''Check if a given CO reference number has already been captured.'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Check if a given CO reference number has already been captured.''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''CheckRefsOnCO''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CheckRefsOnGRN''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''CheckRefsOnGRN'', '''', ''Check if a given GRN invoice number has already been captured.'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Check if a given GRN invoice number has already been captured.''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''CheckRefsOnGRN''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''BookFindFeedInputFormatFilePath''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''BookFindFeedInputFormatFilePath'', ''C:\PBKS\TEMPLATES\ImportBookfindFeedFormat.XML'', ''Path to the format file for Bookfind Feed input'', 13,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Path to the format file for Bookfind Feed input''  ,[PropertyTypeID] = 13,[Scope] = ''C''
		WHERE [PropertyKey] = ''BookFindFeedInputFormatFilePath''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''ClipboardInputFormatFilePath''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''ClipboardInputFormatFilePath'', ''C:\PBKS\TEMPLATES\ImportClipboardFormat.xml'', ''Path to the format file for clipboard input'', 13,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Path to the format file for clipboard input''  ,[PropertyTypeID] = 13,[Scope] = ''C''
		WHERE [PropertyKey] = ''ClipboardInputFormatFilePath''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''COMPORTNumber''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''COMPORTNumber'', ''1'', ''Used for FrontDesk application scanner'', 7,''L'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Used for FrontDesk application scanner''  ,[PropertyTypeID] = 7,[Scope] = ''L''
		WHERE [PropertyKey] = ''COMPORTNumber''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''COMPORTSETTINGS''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''COMPORTSETTINGS'', ''9600,n,8,1'', ''Used for FrontDesk application scanner'', 7,''L'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Used for FrontDesk application scanner''  ,[PropertyTypeID] = 7,[Scope] = ''L''
		WHERE [PropertyKey] = ''COMPORTSETTINGS''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CONNECTIONNAME''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''CONNECTIONNAME'', ''BDC Wireless'', ''Name of the internet connection (where dial-up is used)'', 7,''L'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Name of the internet connection (where dial-up is used)''  ,[PropertyTypeID] = 7,[Scope] = ''L''
		WHERE [PropertyKey] = ''CONNECTIONNAME''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CONTRA_ACCOUNT_INV''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''CONTRA_ACCOUNT_INV'', '''', ''Used with Pastel'', 8,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Used with Pastel''  ,[PropertyTypeID] = 8,[Scope] = ''C''
		WHERE [PropertyKey] = ''CONTRA_ACCOUNT_INV''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CONTRA_ACCOUNT_SINV''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''CONTRA_ACCOUNT_SINV'', '''', ''Used with Pastel'', 8,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Used with Pastel''  ,[PropertyTypeID] = 8,[Scope] = ''C''
		WHERE [PropertyKey] = ''CONTRA_ACCOUNT_SINV''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CostOfSalesAccount''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
	VALUES (''CostOfSalesAccount'', ''2000000'', ''The number of cost of sales account in the accounting system'', 8,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''The number of cost of sales account in the accounting system)''  ,[PropertyTypeID] = 8,[Scope] = ''C''
		WHERE [PropertyKey] = ''CostOfSalesAccount''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CreditNoteDocCodeInsert''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''CreditNoteDocCodeInsert'', ''CN'', ''Prefix to use on a credit note'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Prefix to use on a credit note''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''CreditNoteDocCodeInsert''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CustomerAcnoSequence''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''CustomerAcnoSequence'', ''CA'', ''If CA then customer name followed by acno, else acno followed by customer name'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If CA then customer name followed by acno, else acno followed by customer name''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''CustomerAcnoSequence''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CustomerInputFormatFilePath''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''CustomerInputFormatFilePath'', ''C:\PBKS\TEMPLATES\ImportCustomerFormat.xml'', ''Path to the format file for customer input'', 13,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Path to the format file for customer input''  ,[PropertyTypeID] = 13,[Scope] = ''C''
		WHERE [PropertyKey] = ''CustomerInputFormatFilePath''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''DamagedReturns''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''DamagedReturns'', ''FALSE'', ''If TRUE, Credit notes can track damaged return stock '', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE, Credit notes can track damaged return stock ''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''DamagedReturns''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''DefaultAccountingAccno''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''DefaultAccountingAccno'', '''', ''Used for quick invoices and proformas to associate ad-hoc customers with a single Debtor account on the accounting system'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Used for quick invoices and proformas to associate ad-hoc customers with a single Debtor account on the accounting system''  ,
		[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''DefaultAccountingAccno''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''DEFAULTAREACODE''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''DEFAULTAREACODE'', ''021'', ''Default landline area code for the installation'', 6,''L'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Default landline area code for the installation''  ,[PropertyTypeID] = 6,[Scope] = ''L''
		WHERE [PropertyKey] = ''DEFAULTAREACODE''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''DELAYINSECONDS''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''DELAYINSECONDS'', ''60'', ''unknown'', NULL,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''unknown''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''DELAYINSECONDS''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''DeliveryStyle''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''DeliveryStyle'', ''STD'', ''If BB then allows capture of Product type, category and Multiby at receiving.'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If BB then allows capture of Product type, category and Multiby at receiving.''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''DeliveryStyle''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''DetectOverInvoicing''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''DetectOverInvoicing'', ''FALSE'', ''unknown'', NULL,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''unknown''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''DetectOverInvoicing''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''DiscountToCalculateDefaultCostForTransfers''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''DiscountToCalculateDefaultCostForTransfers'', ''0'', ''unknown'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''unknown''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''DiscountToCalculateDefaultCostForTransfers''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''EDIENABLED''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''EDIENABLED'', ''TRUE'', ''Is EDI enabled '', 3,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Is EDI enabled''  ,[PropertyTypeID] = 3,[Scope] = ''C''
		WHERE [PropertyKey] = ''EDIENABLED''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''JavaMemoryAllocation''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''JavaMemoryAllocation'', ''256'', ''The number of megabytes allocated to Java - FOP uses Java -large documents will need more memory - suggest 256. '', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''The number of megabytes allocated to Java - FOP uses Java -large documents will need more memory - suggest 256.''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''JavaMemoryAllocation''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''EDIFTPADDRESS''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''EDIFTPADDRESS'', ''111.111.111.111'', ''FTP address on EDI site from which Orders are received'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''FTP address on EDI site from which Orders are received''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''EDIFTPADDRESS''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''EDIFTPFOLDER''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''EDIFTPFOLDER'', ''/bt000SA1/temp/EDI'', ''FTP folder on EDI site from which Orders are received'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''FTP folder on EDI site from which Orders are received''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''EDIFTPFOLDER''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''EDIFTPPASSWORD''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''EDIFTPPASSWORD'', ''1beach'', ''Password for EDI FTP folder '', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Password for EDI FTP folder''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''EDIFTPPASSWORD''


SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''EDIFTPUSERNAME''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''EDIFTPUSERNAME'', ''bt000SA1'', ''User name for FTP folder on EDI site '', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''User name for FTP folder on EDI site''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''EDIFTPUSERNAME''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''EMail_INV_ShowHTML''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''EMail_INV_ShowHTML'', ''TRUE'', ''If TRUE, invoices show HTML version in the email'', 10,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, invoices show HTML version in the email''  ,[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''EMail_INV_ShowHTML''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''Email_PO_ShowHTML''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''Email_PO_ShowHTML'', ''TRUE'', ''If TRUE, purchase orders show HTML version in the email'', 10,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, purchase orders show HTML version in the email''  ,[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''Email_PO_ShowHTML''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''EMail_QUOTE_ShowHTML''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''EMail_QUOTE_ShowHTML'', ''TRUE'', ''If TRUE, quotations show HTML version in the email'', 10,''L'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, quotations show HTML version in the email''  ,[PropertyTypeID] = 10,[Scope] = ''L''
		WHERE [PropertyKey] = ''EMail_QUOTE_ShowHTML''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''Email_SalesOrder_ShowHTML''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''Email_SalesOrder_ShowHTML'', ''TRUE'', ''If TRUE, sales orders show HTML version in the email'', 10,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, sales orders show HTML version in the email''  ,[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''Email_SalesOrder_ShowHTML''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''EmailAddressForTesting''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''EmailAddressForTesting'', ''david@papyrussoftware.co.za'', ''Email address to which test emails are sent, usually the local address'', 10,''L'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Email address to which test emails are sent, usually the local address'' ,[PropertyTypeID] = 10,[Scope] = ''L''
		WHERE [PropertyKey] = ''EmailAddressForTesting''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''EmailFrom''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''EmailFrom'', ''bookcott@hermanus.co.za'', ''The senders email to reflect on the email. (Used for direct emailing)'', 10,''L'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''The senders email to reflect on the email. (Used for direct emailing)''  ,[PropertyTypeID] = 10,[Scope] = ''L''
		WHERE [PropertyKey] = ''EmailFrom''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''EMailINV''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''EMailINV'', ''FALSE'', ''If TRUE, invoices are emailed'', 10,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, invoices are emailed''  ,[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''EMailINV''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''EmailPO''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''EmailPO'', ''TRUE'', ''If TRUE, purchase orders are emailed'', 10,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, purchase orders are emailed''  ,[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''EmailPO''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''EMailQUOTE''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''EMailQUOTE'', ''FALSE'', ''If TRUE, quotations are emailed'', 10,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, quotations are emailed''  ,[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''EMailQUOTE''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''EmailSalesOrder''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''EmailSalesOrder'', ''TRUE'', ''If TRUE, sales orders are emailed'', 10,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, sales orders are emailed''  ,[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''EmailSalesOrder''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''ENABLEBOOKCLUBRETURN''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''ENABLEBOOKCLUBRETURN'', ''FALSE'', ''Unknown'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Unknown''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''ENABLEBOOKCLUBRETURN''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''ExcludeActionedFromReorder''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''ExcludeActionedFromReorder'', ''TRUE'', ''If TRUE, Customer orders actioned by purchase orders will not appear in the reorder slate.'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE, Customer orders actioned by purchase orders will not appear in the reorder slate.'' ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''ExcludeActionedFromReorder''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''EXPORTTOPASTELENABLED''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''EXPORTTOPASTELENABLED'', ''TRUE'', ''Allow export to Pastel'', 8,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Allow export to Pastel''  ,[PropertyTypeID] = 8,[Scope] = ''C''
		WHERE [PropertyKey] = ''EXPORTTOPASTELENABLED''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''FETCHLOGSFROM''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''FETCHLOGSFROM'', ''\\POS\PBKS_S'', ''Specify machine and shared folders (separated by commas if more than one) on POS stations from where POSErrors.txt files will be fetched and included in SEND uploads for support'', 7,''L'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Specify machine and shared folders (separated by commas if more than one) on POS stations from where POSErrors.txt files will be fetched and included in SEND uploads for support''  ,[PropertyTypeID] = 7,[Scope] = ''L''
		WHERE [PropertyKey] = ''FETCHLOGSFROM''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''FTPADDRESS''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''FTPADDRESS'', ''207.58.144.36'', ''Papyrus FTP support site'', 2,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Papyrus FTP support site''  ,[PropertyTypeID] = 2,[Scope] = ''C''
		WHERE [PropertyKey] = ''FTPADDRESS''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''FTPFOLDER''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''FTPFOLDER'', ''/public_ftp'', ''FTP folder on Papyrus FTP support site'', 2,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''FTP folder on Papyrus FTP support site''  ,[PropertyTypeID] = 2,[Scope] = ''C''
		WHERE [PropertyKey] = ''FTPFOLDER''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''FTPPASSWORD''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''FTPPASSWORD'', ''3t6h36f9'', ''Password for Papyrus FTP support site'', 2,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Password for Papyrus FTP support site''  ,[PropertyTypeID] = 2,[Scope] = ''C''
		WHERE [PropertyKey] = ''FTPPASSWORD''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''FTPUSERNAME''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''FTPUSERNAME'', ''papyruss'', ''User name for Papyrus FTP support site'', 2,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''User name for Papyrus FTP support site''  ,[PropertyTypeID] = 2,[Scope] = ''C''
		WHERE [PropertyKey] = ''FTPUSERNAME''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''GeneralPurposeYN''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''GeneralPurposeYN'', ''TRUE'', ''If TRUE then the system must not display Book-specific features'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE then the system must not display Book-specific features''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''GeneralPurposeYN''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''GenerateSeparateInvoicesForSeparateOrders''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''GenerateSeparateInvoicesForSeparateOrders'', ''Q'', ''In order fulfilment, if true, it generates separate invoices for separate orders'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''In order fulfilment, if true, it generates separate invoices for separate orders'' ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''GenerateSeparateInvoicesForSeparateOrders''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''GRNInputFormatFilePath''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''GRNInputFormatFilePath'', ''C:\PBKS\TEMPLATES\ImportGRNFormat.xml'', ''Path to the format file for GRN input'', 13,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Path to the format file for GRN input''  ,[PropertyTypeID] = 13,[Scope] = ''C''
		WHERE [PropertyKey] = ''StockInputFormatFilePath''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''GSPRINTLocation''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''GSPRINTLocation'', ''C:\Program Files\Ghostgum\gsview'', ''Location of folder containing GSPrint.EXE'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Location of folder containing GSPrint.EXE''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''GSPRINTLocation''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''HIDELOCALSKUONINV''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''HIDELOCALSKUONINV'', ''FALSE'', ''If TRUE, hides the #number code on an invoice (blanks it)'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE, hides the #number code on an invoice (blanks it)''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''HIDELOCALSKUONINV''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''HUB_ON''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''HUB_ON'', ''FALSE'', ''HUB_ON is either TRUE or FALSE'', 12,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''HUB_ON is either TRUE or FALSE''  ,[PropertyTypeID] = 12,[Scope] = ''C''
		WHERE [PropertyKey] = ''HUB_ON''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''INTERNETDIALUP''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''INTERNETDIALUP'', ''FALSE'', ''Must the computer dial out to get an internet connection'', 7,''L'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Must the computer dial out to get an internet connection''  ,[PropertyTypeID] = 7,[Scope] = ''L''
		WHERE [PropertyKey] = ''INTERNETDIALUP''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''InventoryAccountingModel''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''InventoryAccountingModel'', ''PERIODIC'', ''Either PERIODIC or PERPETUAL'', 8,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Either PERIODIC or PERPETUAL''  ,[PropertyTypeID] = 8,[Scope] = ''C''
		WHERE [PropertyKey] = ''InventoryAccountingModel''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''InventoryControlAccount''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
	VALUES (''InventoryControlAccount'', ''7700001'', ''The number of inventory control account in the accounting system'', 8,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''The number of inventory control account in the accounting system)''  ,[PropertyTypeID] = 8,[Scope] = ''C''
		WHERE [PropertyKey] = ''InventoryControlAccount''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''InvoiceDocCodeInsert''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''InvoiceDocCodeInsert'', ''IV'', ''Prefix to use on invoice number'', 7,''C'') 
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Prefix to use on invoice number'' ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''InvoiceDocCodeInsert''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''InvoiceQuoteCodeInsert''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''InvoiceQuoteCodeInsert'', ''Q'', ''Code for quotations'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Code for quotations'' ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''InvoiceQuoteCodeInsert''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''InvoiceSubject''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''InvoiceSubject'', ''Invoice'', ''Used on email heading'', 10,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Used on email heading''  ,[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''InvoiceSubject''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''INVTOTALSEQ''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''INVTOTALSEQ'', ''E,V,NV'', ''Controls how the totals on an invoice are presented'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Controls how the totals on an invoice are presented''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''INVTOTALSEQ''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''ISSUEBOOKCLUBRETURNDOCS''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''ISSUEBOOKCLUBRETURNDOCS'', ''TRUE'', ''Unknown'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Unknown''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''ISSUEBOOKCLUBRETURNDOCS''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''IssueQuickCOs''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''IssueQuickCOs'', ''FALSE'', ''If TRUE then quick COs are issued immediately, else they must be reviewed, edited and issued.'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE then quick COs are issued immediately, else they must be reviewed, edited and issued.''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''IssueQuickCOs''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''IssueQuickInvoices''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''IssueQuickInvoices'', ''FALSE'', ''If TRUE then quick invoices are issued immediately, else they must be reviewed, edited and issued.'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE then quick invoices are issued immediately, else they must be reviewed, edited and issued.''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''IssueQuickInvoices''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''IssueQuickPFs''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''IssueQuickPFs'', ''FALSE'', ''If TRUE then quick PFs are issued immediately, else they must be reviewed, edited and issued.'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE then quick PFs are issued immediately, else they must be reviewed, edited and issued.''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''IssueQuickPFs''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''KeepTemporaryfilesFor_n_Days''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''KeepTemporaryfilesFor_n_Days'', ''3'', ''Sets the number of days that temporary files in the PSF and TEMP folders are kept before deleting.'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Sets the number of days that temporary files in the PSF and TEMP folders are kept before deleting.''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''KeepTemporaryfilesFor_n_Days''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''LABELPRINTER''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''LABELPRINTER'', ''OKI'', ''Type of label printer being used'', NULL,''L'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Type of label printer being used'',[PropertyTypeID] = 5,[Scope] = ''L''
		WHERE [PropertyKey] = ''LABELPRINTER''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''LOYALTYtoCENTRAL_ON''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''LOYALTYtoCENTRAL_ON'', ''FALSE'', ''LOYALTYtoCENTRAL_ON is either TRUE or FALSE'', 12,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''LOYALTYtoCENTRAL_ON is either TRUE or FALSE''  ,[PropertyTypeID] = 12,[Scope] = ''C''
		WHERE [PropertyKey] = ''LOYALTYtoCENTRAL_ON''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''MarkPOLasFulfilledWhenShortSupplied''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''MarkPOLasFulfilledWhenShortSupplied'', ''FALSE'', ''If TRUE then POL is set as fulfilled when a claim is placed instead of a delivery being received'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE then POL is set as fulfilled when a claim is placed instead of a delivery being received''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''MarkPOLasFulfilledWhenShortSupplied''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''MAXBROWSE''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''MAXBROWSE'', ''1100'', ''The maximum number of rows to be returned when browsing items within Papyrus (to prevent long waits)'', 5,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''The maximum number of rows to be returned when browsing items within Papyrus (to prevent long waits)''  ,[PropertyTypeID] = 5,[Scope] = ''C''
		WHERE [PropertyKey] = ''MAXBROWSE''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''MECPath''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''MECPath'', ''PBKS_S\EXECUTABLES'', ''Path for MEC Executables'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Path for MEC Executables''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''MECPath''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''MultiStore''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''MultiStore'', ''FALSE'', ''If TRUE then DB is recording stock for more than one store'', 10,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE then DB is recording stock for more than one store''  ,[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''MultiStore''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''OnlyActiveAccounts''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''OnlyActiveAccounts'', ''TRUE'', ''If TRUE, only debtors with balances will print statements'', 11,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE, only debtors with balances will print statements''  ,[PropertyTypeID] = 11,[Scope] = ''C''
		WHERE [PropertyKey] = ''OnlyActiveAccounts''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''OrderDocCodeInsert''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''OrderDocCodeInsert'', ''C'', ''Prefix to use on customer order'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Prefix to use on customer order''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''OrderDocCodeInsert''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''OutlookCustomFolderForEmails''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''OutlookCustomFolderForEmails'', ''PapyrusEmailDrafts'', ''Folder where emails are placed if drafts folder not wanted'', 10,''L'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Folder where emails are placed if drafts folder not wanted'' ,[PropertyTypeID] = 10,[Scope] = ''L''
		WHERE [PropertyKey] = ''OutlookCustomFolderForEmails''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''OutlookParentOfCustomFolder''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''OutlookParentOfCustomFolder'', ''Personal folders'', ''The folders group inside which the custom folder is placed'', 10,''L'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''The folders group inside which the custom folder is placed'' ,[PropertyTypeID] = 10,[Scope] = ''L''
		WHERE [PropertyKey] = ''OutlookParentOfCustomFolder''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''PDFPrintTool''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''PDFPrintTool'', ''A'', ''X-XPDFPRINT,G-Ghostscript, O-Other (Foxit or Adobe Reader)'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''X-XPDFPRINT,G-Ghostscript, O-Other (Foxit or Adobe Reader)'' ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''PDFPrintTool''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''POSACTIVE''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''POSACTIVE'', ''TRUE'', ''Supports Point of Sale'', 3,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Supports Point of Sale''  ,[PropertyTypeID] = 3,[Scope] = ''C''
		WHERE [PropertyKey] = ''POSACTIVE''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''POSUsesSB''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''POSUsesSB'', ''FALSE'', ''If TRUE then POS messages are carried using service broker'', 10,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE then POS messages are carried using service broker''  ,[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''POSUsesSB''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''PrintPackingSlip''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''PrintPackingSlip'', ''FALSE'', ''If TRUE, a packing slip is printed after the invoice itself.'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE, a packing slip is printed after the invoice itself.''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''PrintPackingSlip''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''PRINTSERVERMACHINE''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''PRINTSERVERMACHINE'', '''', ''The network name of the computer running Dispatch'', 7,''L'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''The network name of the computer running Dispatch''  ,[PropertyTypeID] = 7,[Scope] = ''L''
		WHERE [PropertyKey] = ''PRINTSERVERMACHINE''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''PTInputFormatFilePath''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''PTInputFormatFilePath'', ''C:\PBKS\TEMPLATES\ImportPTFormat.xml'', ''Path to the format file for product types'', 13,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Path to the format file for for product types''  ,[PropertyTypeID] = 13,[Scope] = ''C''
		WHERE [PropertyKey] = ''PTInputFormatFilePath''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''PurchaseOrderInputFormatFilePath''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''PurchaseOrderInputFormatFilePath'', ''C:\PBKS\TEMPLATES\ImportPurchaseOrderFormat.xml'', ''Path to the format file for purchase order input'', 13,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Path to the format file for purchase order input''  ,[PropertyTypeID] = 13,[Scope] = ''C''
		WHERE [PropertyKey] = ''PurchaseOrderInputFormatFilePath''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''QuoteCOStaffNameOnInvoice''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''QuoteCOStaffNameOnInvoice'', ''FALSE'', ''If TRUE, Papyrus uses staff name of person handling CO on the invoice, else person signing invoice'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE, Papyrus uses staff name of person handling CO on the invoice, else person signing invoice''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''QuoteCOStaffNameOnInvoice''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''ROUNDPRICETO''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''ROUNDPRICETO'', ''0'', ''Not used'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Not used''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''ROUNDPRICETO''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''RunsAccountsTF''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''RunsAccountsTF'', ''FALSE'', ''If TRUE, debtors accounts will be supported'', 11,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE, debtors accounts will be supported''  ,[PropertyTypeID] = 11 ,[Scope] = ''C''
		WHERE [PropertyKey] = ''RunsAccountsTF''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SalesAccount''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
	VALUES (''SalesAccount'', ''1000000'', ''The number of the sales account in the accounting system'', 8,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''The number of the sales account in the accounting system)''  ,[PropertyTypeID] = 8,[Scope] = ''C''
		WHERE [PropertyKey] = ''SalesAccount''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SalesOrderInputFormatFilePath''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SalesOrderInputFormatFilePath'', ''C:\PBKS\TEMPLATES\ImportSalesOrderFormat.xml'', ''Path to the format file for sales order input'', 13,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Path to the format file for sales order input''  ,[PropertyTypeID] = 13,[Scope] = ''C''
		WHERE [PropertyKey] = ''SalesOrderInputFormatFilePath''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SALEStoCENTRAL_ON''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SALEStoCENTRAL_ON'', ''FALSE'', ''SALEStoCENTRAL_ON is either TRUE or FALSE'', 12,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''SALEStoCENTRAL_ON is either TRUE or FALSE''  ,[PropertyTypeID] = 12,[Scope] = ''C''
		WHERE [PropertyKey] = ''SALEStoCENTRAL_ON''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SecondaryEDIAddress''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SecondaryEDIAddress'', ''111.111.111.111'', ''BackupEDI'', 7,''L'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''BackupEDI''  ,[PropertyTypeID] = 7,[Scope] = ''L''
		WHERE [PropertyKey] = ''SecondaryEDIAddress''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SECURETF''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SECURETF'', ''FALSE'', ''If TRUE, price changes and discount changes will be tracked and need authorization'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE, price changes and discount changes will be tracked and need authorization''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''SECURETF''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SenderName''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SenderName'', ''Sue McNaught'', ''The senders name to reflect on the email. (Used for direct emailing)'', 10,''L'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''The senders name to reflect on the email. (Used for direct emailing)'' ,[PropertyTypeID] = 10,[Scope] = ''L'' 
		WHERE [PropertyKey] = ''SenderName''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SendsCR''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SendsCR'', ''False'', ''Used for FrontDesk application scanner'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Used for FrontDesk application scanner''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''SendsCR''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''Service_Broker_ON''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''Service_Broker_ON'', ''FALSE'', ''Service_Broker_ON is either TRUE or FALSE'', 12,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Service_Broker_ON is either TRUE or FALSE''  ,[PropertyTypeID] = 12,[Scope] = ''C''
		WHERE [PropertyKey] = ''Service_Broker_ON''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''ServiceBroker_Alerts_ON''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''ServiceBroker_Alerts_ON'', ''FALSE'', ''If TRUE then Service broker enabled Alerts are active'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE then Service broker enabled Alerts are active''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''ServiceBroker_Alerts_ON''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''ServiceBroker_IBTs_ON''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''ServiceBroker_IBTs_ON'', ''FALSE'', ''If TRUE then Service broker enabled IBTs are active'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE then Service broker enabled IBTs are active''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''ServiceBroker_IBTs_ON''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SetPricesInGRN''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SetPricesInGRN'', ''TRUE'', ''If TRUE, selling prices can be changed in the receiving process. (Use FALSE for distributors only)'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, selling prices can be changed in the receiving process. (Use FALSE for distributors only)''  ,
								[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''SetPricesInGRN''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SetSupplierIDFROMDEL''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SetSupplierIDFROMDEL'', ''TRUE'', ''Remember the supplier of a book when you receive it. '', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Remember the supplier of a book when you receive it. ''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''SetSupplierIDFROMDEL''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SetSupplierIDFROMPO''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SetSupplierIDFROMPO'', ''TRUE'', ''Remember the supplier of a book when you order it..'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Remember the supplier of a book when you order it.''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''SetSupplierIDFROMPO''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SHOWALLAPPROS''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SHOWALLAPPROS'', ''TRUE'', ''On a product record, if TRUE, shows all appros for the product, else irt shows only the o/s appros'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''On a product record, if TRUE, shows all appros for the product, else irt shows only the o/s appros''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''SHOWALLAPPROS''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''ShowCategoryOrPTInReorderSlate''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
	VALUES (''ShowCategoryOrPTInReorderSlate'', ''PT'', ''In the reorder slate show either Categories (CAT) or Product types (PT)'',7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''In the reorder slate show either Categories (CAT) or Product types (PT)''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''ShowCategoryOrPTInReorderSlate''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SHOWQUOTES''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SHOWQUOTES'', ''TRUE'', ''Display quotes on status bar'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Display quotes on status bar''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''SHOWQUOTES''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''ShowWordstockSales''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''ShowWordstockSales'', ''FALSE'', ''If TRUE then the system will display Wordstock sales from transferred data (Used after taking over from Wordstock). It can be turned off later when data is redundant'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE then the system will display Wordstock sales from transferred data (Used after taking over from Wordstock). It can be turned off later when data is redundant''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''ShowWordstockSales''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SMTP_Password''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SMTP_Password'', ''96989698'', ''Used for direct emailing'', 10,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Used for direct emailing''  ,[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''SMTP_Password''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SMTP_Username''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SMTP_Username'', ''bookcottage@whalemail.co.za'', ''Used for direct emailing'', 10,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Used for direct emailing''  ,[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''SMTP_Username''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SMTPServer''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SMTPServer'', ''mail.whalemail.co.za'', ''Used for direct emailing'', 10,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Used for direct emailing''  ,[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''SMTPServer''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SOH_ON''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SOH_ON'', ''FALSE'', ''SOH_ON is either TRUE or FALSE'', 12,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''SOH_ON is either TRUE or FALSE''  ,[PropertyTypeID] = 12,[Scope] = ''C''
		WHERE [PropertyKey] = ''SOH_ON''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''Cashup_Reporting_ON''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''Cashup_Reporting_ON'', ''FALSE'', ''Cashup_Reporting_ON is either TRUE or FALSE'', 12,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Cashup_Reporting_ON''  ,[PropertyTypeID] = 12,[Scope] = ''C''
		WHERE [PropertyKey] = ''Cashup_Reporting_ON''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''COLS_Reporting_ON''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''COLS_Reporting_ON'', ''FALSE'', ''COLS_Reporting_ON is either TRUE or FALSE'', 12,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''COLS_Reporting_ON''  ,[PropertyTypeID] = 12,[Scope] = ''C''
		WHERE [PropertyKey] = ''COLS_Reporting_ON''


SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''StockCategoryInputFormatFilePath''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''StockCategoryInputFormatFilePath'', ''C:\PBKS\TEMPLATES\ImportStockCategoryFormat.xml'', ''Path to the format file for stock category input'', 13,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Path to the format file for stock category input''  ,[PropertyTypeID] = 13,[Scope] = ''C''
		WHERE [PropertyKey] = ''StockCategoryInputFormatFilePath''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''StockInputFormatFilePath''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''StockInputFormatFilePath'', ''C:\PBKS\TEMPLATES\ImportStockFormat.xml'', ''Path to the format file for stock input'', 13,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Path to the format file for stock input''  ,[PropertyTypeID] = 13,[Scope] = ''C''
		WHERE [PropertyKey] = ''StockInputFormatFilePath''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''StopOverInvoicing''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''StopOverInvoicing'', ''FALSE'', ''Stop operator when invoicing takes the qty on hand negative. '', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Stop operator when invoicing takes the qty on hand negative.''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''StopOverInvoicing''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''StoreType''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''StoreType'', ''INDEPENDENT'', ''StoreType is either HO or INDEPENDENT OR BRANCH. This controls how data is exported to Pastel'', 8,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''StoreType is either HO or INDEPENDENT OR BRANCH. This controls how data is exported to Pastel''  ,[PropertyTypeID] = 8,[Scope] = ''C''
		WHERE [PropertyKey] = ''StoreType''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''Subject''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''Subject'', ''ORDER FROM Book Cottage Hermanus'', ''The subject to be displayed on a Purchase order email'', 10,''L'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''The subject to be displayed on a Purchase order email''  ,[PropertyTypeID] = 10,[Scope] = ''L''
		WHERE [PropertyKey] = ''Subject''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SubscriptionOrderSuffix''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
	VALUES (''SubscriptionOrderSuffix'', '''', ''Uses this single character to append to the document number of subscription purchase orders'',7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Uses this single character to append to the document number of subscription purchase orders''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''SubscriptionOrderSuffix''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SUPPLERINVOICETOLERANCE''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SUPPLERINVOICETOLERANCE'', ''.005'', ''The degree to which the Papyrus II calculated GRN may differ from the suppliers invoice total (to allow for rounding errors)'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''The degree to which the Papyrus II calculated GRN may differ from the suppliers invoice total (to allow for rounding errors)''  ,	
		[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''SUPPLERINVOICETOLERANCE''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SupplierBasedCurrencyConversion''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SupplierBasedCurrencyConversion'', ''TRUE'', ''If TRUE, local currency value is determined from the foreign currency..'', NULL,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Local currency value is determined from the foreign currency using this rate''   ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''SupplierBasedCurrencyConversion''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SupplierInputFormatFilePath''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SupplierInputFormatFilePath'', ''C:\PBKS\TEMPLATES\ImportSupplierFormat.xml'', ''Path to the format file for supplier input'',13,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Path to the format file for suppliers input''  ,[PropertyTypeID] = 13,[Scope] = ''C''
		WHERE [PropertyKey] = ''SupplierInputFormatFilePath''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SupportsBookclubs''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SupportsBookclubs'', ''TRUE'', ''If TRUE, book clubs are supported'', 3,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, book clubs are supported''  ,[PropertyTypeID] = 3,[Scope] = ''C''
		WHERE [PropertyKey] = ''SupportsBookclubs''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SupportsCatalogue''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SupportsCatalogue'', ''FALSE'', ''Supports the creation of printed catalogue with hierarchical headings using MS-WORD and VBA'', 3,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Supports the creation of printed catalogue with hierarchical headings using MS-WORD and VBA''  ,[PropertyTypeID] = 3,[Scope] = ''C''
		WHERE [PropertyKey] = ''SupportsCatalogue''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SupportsLoyaltyCustomers''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''SupportsLoyaltyCustomers'', ''TRUE'', ''If TRUE, loyalty customers are supported'', 3,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, loyalty customers are supported''  ,[PropertyTypeID] = 3,[Scope] = ''C''
		WHERE [PropertyKey] = ''SupportsLoyaltyCustomers''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''SUPPORTSMULTIBUY''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) VALUES (''SUPPORTSMULTIBUY'', ''FALSE'', ''Bargain Books style 3 for R99'', 3,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Bargain Books style 3 for R99''   ,[PropertyTypeID] = 3,[Scope] = ''C''
		WHERE [PropertyKey] = ''SUPPORTSMULTIBUY''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''TestMode''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
	VALUES (''TestMode'', ''TRUE'', ''If TRUE then EMails are sent to sender rather than to real trading partners.'', 10,''L'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE then EMails are sent to sender rather than to real trading partners.''  ,[PropertyTypeID] = 10,[Scope] = ''L''
		WHERE [PropertyKey] = ''TestMode''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''TIMERINTERVAL''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''TIMERINTERVAL'', ''3000'', ''Used for FrontDesk application scanner'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Used for FrontDesk application scanner''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''TIMERINTERVAL''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''TRANSFERCALC''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''TRANSFERCALC'', ''VATDISC'', ''Manages calculations of I.B.T.'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Manages calculations of I.B.T.''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''TRANSFERCALC''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''TransferIsExVAT''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''TransferIsExVAT'', ''TRUE'', ''If TRUE, I.B.T. does not reflect VAT'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, I.B.T. does not reflect VAT''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''TransferIsExVAT''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''Translation_offset''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''Translation_offset'', ''0'', ''Offset in string resource table'', 12,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Offset in string resource table''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''Translation_offset''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''UNISASupport''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''UNISASupport'', ''3'', ''Key for Unisa support'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Key for Unisa support''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''UNISASupport''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''UpdatePriceOnForeignDelivery''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''UpdatePriceOnForeignDelivery'', ''FALSE'', ''If TRUE then Foreign prices are converted to local prices and the local selling prices are updated'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE then Foreign prices are converted to local prices and the local selling prices are updated''  ,[PropertyTypeID]= 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''UpdatePriceOnForeignDelivery''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''UsePickingSlip''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''UsePickingSlip'', ''FALSE'', ''If TRUE, additional status of ''''Picking'''' is supported (not certain about this)'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, additional status of ''''Picking'''' is supported (not certain about this)''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''UsePickingSlip''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''UsesHUB''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''UsesHUB'', ''FALSE'', ''If TRUE, the tools to connect to the HUB are made visible on the forms as required '', 3,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE, the tools to connect to the HUB are made visible on the forms as required ''  ,[PropertyTypeID] = 3,[Scope] = ''C''
		WHERE [PropertyKey] = ''UsesHUB''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''UsesOutlookForCOEmail''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''UsesOutlookForCOEmail'', ''TRUE'', ''If TRUE, sales orders are emailed using Outlook'', 10,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE, sales orders are emailed using Outlook''  ,[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''UsesOutlookForCOEmail''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''UsesOutlookForINVEmail''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''UsesOutlookForINVEmail'', ''FALSE'', ''If TRUE, invoices are emailed using Outlook'', 10,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, invoices are emailed using Outlook''  ,[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''UsesOutlookForINVEmail''


SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''UsesOutlookForPOEmail''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''UsesOutlookForPOEmail'', ''TRUE'', ''If TRUE, purchase orders are emailed using Outlook'', 10,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, purchase orders are emailed using Outlook''  ,[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''UsesOutlookForPOEmail''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''UsesOutlookForQuoteEmail''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''UsesOutlookForQuoteEmail'', ''TRUE'', ''If TRUE, quotations are emailed using Outlook'', 10,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, quotations are emailed using Outlook''  ,[PropertyTypeID] = 10,[Scope] = ''C''
		WHERE [PropertyKey] = ''UsesOutlookForQuoteEmail''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''UseXalan''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''UseXalan'', ''FALSE'', ''If TRUE, we use XALAN (thereby enabling barcodes in stylesheets'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE, we use XALAN (thereby enabling barcodes in stylesheets''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''UseXalan''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''UseXMLPrintingForAPP''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''UseXMLPrintingForAPP'', ''FALSE'', ''If TRUE, Papyrus uses XML and Stylesheet template to print APP.'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE, Papyrus uses XML and Stylesheet template to print APP''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''UseXMLPrintingForAPP''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''UseXMLPrintingForCO''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''UseXMLPrintingForCO'', ''FALSE'', ''If TRUE, Papyrus uses XML and Stylesheet template to print CO.'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE, Papyrus uses XML and Stylesheet template to print CO''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''UseXMLPrintingForCO''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''UseXMLPrintingForGRN''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''UseXMLPrintingForGRN'', ''FALSE'', ''If TRUE, Papyrus uses XML and Stylesheet template to print GRN.'', 7,''L'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''If TRUE, Papyrus uses XML and Stylesheet template to print GRN''  ,[PropertyTypeID] = 7,[Scope] = ''L''
		WHERE [PropertyKey] = ''UseXMLPrintingForGRN''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''VATControlAccount''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
	VALUES (''VATControlAccount'', ''9500000'', ''The number of VAT Control Account in the accounting system'', 8,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''The number of VAT Control Account in the accounting system)''  ,[PropertyTypeID] = 8,[Scope] = ''C''
		WHERE [PropertyKey] = ''VATControlAccount''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''VOUCHERREPORTTOGETHER''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''VOUCHERREPORTTOGETHER'', '''', ''Not used'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''Not used''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''VOUCHERREPORTTOGETHER''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''WarnOverInvoicing''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''WarnOverInvoicing'', ''FALSE'', ''Warn operator when invoicing takes the qty on hand negative. '', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Warn operator when invoicing takes the qty on hand negative.''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''WarnOverInvoicing''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''WordstockSalesInputFormatFilePath''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''WordstockSalesInputFormatFilePath'', ''C:\PBKS\TEMPLATES\ImportWordstockSalesFormat.xml'', ''Path to the format file for Wordstock sales input'', 13,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Path to the format file for Wordstock sales input''  ,[PropertyTypeID] = 13,[Scope] = ''C''
		WHERE [PropertyKey] = ''WordstockSalesInputFormatFilePath''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''UniqueProducts''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''UniqueProducts'', ''FALSE'', ''UniqueProducts is either TRUE or FALSE'', 12,''C'')
ELSE
	UPDATE [dbo].[tProperty] 
	SET [PropertyDescription] = ''UniqueProducts is either TRUE or FALSE''  ,[PropertyTypeID] = 12,[Scope] = ''C''
		WHERE [PropertyKey] = ''UniqueProducts''


SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''DimensionMeasurementUnits''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''DimensionMeasurementUnits'', ''m'', ''Capture and format dimensions in this unit'', 7,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Capture and format dimensions in this unit''  ,[PropertyTypeID] = 7,[Scope] = ''C''
		WHERE [PropertyKey] = ''DimensionMeasurementUnits''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''AllowSupplierDetailsCaptureInCustomerOrder''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''AllowSupplierDetailsCaptureInCustomerOrder'', ''FALSE'', ''Provides for recording supplier  ID, Price and discount when capturing a title on a customer order'', 3,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Provides for recording supplier  ID, Price and discount when capturing a title on a customer order''  ,
			[PropertyTypeID] =3,[Scope] = ''C''
		WHERE [PropertyKey] = ''AllowSupplierDetailsCaptureInCustomerOrder''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''UsePapyrusExchangeRateWhenExportingToAccounting''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''UsePapyrusExchangeRateWhenExportingToAccounting'', ''TRUE'', ''Supply currency conversion rates stores in Papyrus to Pastel when exporting.'', 8,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''Supply currency conversion rates stores in Papyrus to Pastel when exporting.''  ,
			[PropertyTypeID] =8,[Scope] = ''C''
		WHERE [PropertyKey] = ''UsePapyrusExchangeRateWhenExportingToAccounting''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''NonTaxableCodeInAccountingApplication''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''NonTaxableCodeInAccountingApplication'', ''0'', ''The code that the accounting application (e.g. Pastel) uses to mark non VAT sales etc).'', 8,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''The code that the accounting application (e.g. Pastel) uses to mark non VAT sales etc).''  ,
			[PropertyTypeID] =8,[Scope] = ''C''
		WHERE [PropertyKey] = ''NonTaxableCodeInAccountingApplication''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''ShowCustomerAcnoLeftOfName''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription], [PropertyTypeID],[Scope]) 
		VALUES (''ShowCustomerAcnoLeftOfName'', ''TRUE'', ''In forms show the customer Acno in front of the customer name).'', 8,''C'')
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''In forms show the customer Acno in front of the customer name.''  ,
			[PropertyTypeID] =7,[Scope] = ''C''
		WHERE [PropertyKey] = ''ShowCustomerAcnoLeftOfName''

END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.InitializeProperties Procedure Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.InitializeProperties Procedure'
END
GO

--
-- Script To Update dbo.PBKS_INITIALIZEDATA Procedure In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.PBKS_INITIALIZEDATA Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER PROCEDURE [dbo].[PBKS_INITIALIZEDATA] AS
DECLARE @BC_ID INT
DECLARE @UNCT_ID INT
DECLARE @LOYCT1_ID INT
DECLARE @LOYCT2_ID INT
DECLARE @LOYCT3_ID INT
DECLARE @WEB_ID INT
DECLARE @MBC_ID INT
DECLARE @ID INT
DECLARE @MB INT

	SELECT [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''DS''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICTTYPES (DESCRIPTION,[System]) VALUES (''Dispatch types'',''DS'')
	ELSE
		UPDATE [dbo].tDICTTYPES SET [DESCRIPTION] = ''Dispatch types'' WHERE [SYSTEM] = ''DS''

	SELECT [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''IG''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICTTYPES (DESCRIPTION,[System]) VALUES (''Interest group'',''IG'')
	ELSE
		UPDATE [dbo].tDICTTYPES SET [DESCRIPTION] = ''Interest group'' WHERE [SYSTEM] = ''IG''

	SELECT [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''CT''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICTTYPES (DESCRIPTION,[System]) VALUES (''Customer type'',''CT'')
	ELSE
		UPDATE [dbo].tDICTTYPES SET [DESCRIPTION] = ''Customer type'' WHERE [SYSTEM] = ''CT''

	SELECT [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''ST''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICTTYPES (DESCRIPTION,[System]) VALUES (''Supplier type'',''ST'')
	ELSE
		UPDATE [dbo].tDICTTYPES SET [DESCRIPTION] = ''Supplier type'' WHERE [SYSTEM] = ''ST''

	SELECT [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''SE''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICTTYPES (DESCRIPTION,[System]) VALUES (''Category'',''SE'')
	ELSE
		UPDATE [dbo].tDICTTYPES SET [DESCRIPTION] = ''Category'' WHERE [SYSTEM] = ''SE''

	SELECT [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''DT''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICTTYPES (DESCRIPTION,[System]) VALUES (''Document type'',''DT'')
	ELSE
		UPDATE [dbo].tDICTTYPES SET [DESCRIPTION] = ''Document type'' WHERE [SYSTEM] = ''DT''

	SELECT [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''SR''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICTTYPES (DESCRIPTION,[System]) VALUES (''Supplier claim reasons'',''SR'')
	ELSE
		UPDATE [dbo].tDICTTYPES SET [DESCRIPTION] = ''Supplier claim reasons'' WHERE [SYSTEM] = ''SR''

	SELECT [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''TB''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICTTYPES (DESCRIPTION,[System]) VALUES (''Text bites'',''TB'')
	ELSE
		UPDATE [dbo].tDICTTYPES SET [DESCRIPTION] = ''Text bites'' WHERE [SYSTEM] = ''TB''

	SELECT [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''PS''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICTTYPES (DESCRIPTION,[System]) VALUES (''Product availablility status'',''PS'')
	ELSE
		UPDATE [dbo].tDICTTYPES SET [DESCRIPTION] = ''Product availablility status'' WHERE [SYSTEM] = ''PS''

	SELECT [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''OA''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICTTYPES (DESCRIPTION,[System]) VALUES (''Order action code'',''OA'')
	ELSE
		UPDATE [dbo].tDICTTYPES SET [DESCRIPTION] = ''Order action code'' WHERE [SYSTEM] = ''OA''

	SELECT [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''MB''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICTTYPES (DESCRIPTION,[System]) VALUES (''Multi-buy categories'',''MB'')
	ELSE
		UPDATE [dbo].tDICTTYPES SET [DESCRIPTION] = ''Multi-buy categories'' WHERE [SYSTEM] = ''MB''

	SELECT [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''AU''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICTTYPES (DESCRIPTION,[System]) VALUES (''Audit points'',''AU'')
	ELSE
		UPDATE [dbo].tDICTTYPES SET [DESCRIPTION] = ''Audit points'' WHERE [SYSTEM] = ''AU''
--Audit points

	SELECT @ID = [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''AU''

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''SP'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Sell Price'',''SP'',''SP'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''MB'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Multibuy'',''MB'',''MB'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''RRP'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''R.R.P.'',''RRP'',''RRP'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''COST'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Cost Price'',''COST'',''COST'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''SSP'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Special S.P.'',''SSP'',''SSP'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''CL'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Credit Lim.'',''CL'',''CL'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''CD'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Cust. Disc.'',''CD'',''CD'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''NDA'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''No Disc.Allowed'',''NDA'',''NDA'',1)


--Dictionary entries
	--Document types

	SELECT @ID = [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''DT''

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''IN''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Invoice'',''INV'',''IN'',1)
	
	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''PO''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Purchase order'',''PO'',''PO'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''AP''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Appro'',''APP'',''AP'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''AR''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Appro return'',''APR'',''AR'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''TF''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Transfer'',''TFR'',''TF'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''DE''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Delivery'',''DEL'',''DE'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''CO''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Customer order'',''CO'',''CO'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''R''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Return'',''R'',''R'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''AS''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Appro slip'',''AS'',''AS'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''CN''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Credit note'',''CN'',''CN'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''QU''
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Quotation'',''QU'',''QU'',1)
---------------------------------------------
---Set up product status types
--DECLARE @ID INT
	SELECT @ID = [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''PS''

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''OP'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(OP)Out of print'',''OP'',''OP'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(OP)Out of print'' WHERE DICT_SYSTEM = ''OP'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''NP'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(MP)Not yet published'',''NP'',''NP'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(MP)Not yet published'' WHERE DICT_SYSTEM = ''NP'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''BO'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(BO)On backorder'',''BO'',''BO'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(BO)On backorder'' WHERE DICT_SYSTEM = ''BO'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''RP'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(RP)Reprinting'',''RP'',''RP'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(RP)Reprinting'' WHERE DICT_SYSTEM = ''RP'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''IP'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(IP)In print'',''IP'',''IP'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(IP)In print'' WHERE DICT_SYSTEM = ''IP'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''AB'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(AB)Publication abandoned'',''AB'',''AB'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(AB)Publication abandoned'' WHERE DICT_SYSTEM = ''AB'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''AD'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(AD)Apply direct: item not available to trade'',''AD'',''AD'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(AD)Apply direct: item not available to trade'' WHERE DICT_SYSTEM = ''AD'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''AS'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(AS)Already supplied'',''AS'',''AS'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(AS)Already supplied'' WHERE DICT_SYSTEM = ''AS'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''CS'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(CS)Status uncertain: check with customer service'',''CS'',''CS'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(CS)Status uncertain: check with customer service'' WHERE DICT_SYSTEM = ''CS'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''DQ'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) 
						VALUES (@ID,''(DQ)Discount query: available, but discount claimed in order is unacceptable'',''DQ'',''DQ'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(DQ)Discount query: available, but discount claimed in order is unacceptable'' WHERE DICT_SYSTEM = ''DQ'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''HK'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(HK)Paperback out of print: hardback available'',''HK'',''HK'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(HK)Paperback out of print: hardback available'' WHERE DICT_SYSTEM = ''HK'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''MD'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(MD)Manufactured on demand'',''MD'',''MD'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(MD)Manufactured on demand'' WHERE DICT_SYSTEM = ''MD'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''NK'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(NK)Item not known (cannot be traced)'',''NK'',''NK'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(NK)Item not known (cannot be traced)'' WHERE DICT_SYSTEM = ''NK'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''NS'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(NS)Not sold separately'',''NS'',''NS'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(NS)Not sold separately'' WHERE DICT_SYSTEM = ''NS'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''OF'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(OF)This format out of print: other format available'',''OF'',''OF'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(OF)This format out of print: other format available'' WHERE DICT_SYSTEM = ''OF'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''RM'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(RM)Remaindered'',''RM'',''RM'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(RM)Remaindered'' WHERE DICT_SYSTEM = ''RM'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''RR'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(RR)Rights restricted: cannot supply'',''RR'',''RR'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(RR)Rights restricted: cannot supply'' WHERE DICT_SYSTEM = ''RR'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''RF'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(RF)Refer to other publisher or distributor'',''RF'',''RF'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(RF)Refer to other publisher or distributor'' WHERE DICT_SYSTEM = ''RF'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''PK'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(PK)Hardback out of print: paperback available'',''PK'',''PK'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(PK)Hardback out of print: paperback available'' WHERE DICT_SYSTEM = ''PK'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''PQ'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) 
					VALUES (@ID,''(PQ)Price query: available, but query whether price is acceptable'',''PQ'',''PQ'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(PQ)Price query: available, but query whether price is acceptable'' WHERE DICT_SYSTEM = ''PQ'' AND DICT_TYPE = @ID
	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''SO'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(SO)Pack not available: available as single items only'',''SO'',''SO'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(SO)Pack not available: available as single items only'' WHERE DICT_SYSTEM = ''SO'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''TO'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(TO)Only to order'',''TO'',''TO'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(TO)Only to order'' WHERE DICT_SYSTEM = ''TO''  AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''TU'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) 
						VALUES (@ID,''(TU)Temporarily unavailable, but expected to be available again (including “reprinting”)'',''TU'',''TU'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(TU)Temporarily unavailable, but expected to be available again (including “reprinting”)'' 
					WHERE DICT_SYSTEM = ''TU'' AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''TO'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(TO)Only to order'',''TO'',''TO'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(TO)Only to order'' WHERE DICT_SYSTEM = ''TO''  AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''UC'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(UC)Unavailable, and may or may not become available again'',''UC'',''UC'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(UC)Unavailable, and may or may not become available again'' WHERE DICT_SYSTEM = ''UC''  AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''TH'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(TH)Temporary hold: definitive response is delayed'',''TH'',''TH'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(TH)Temporary hold: definitive response is delayed'' WHERE DICT_SYSTEM = ''TH''  AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''ST'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(ST)Stocktaking: temporarily unavailable'',''ST'',''ST'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(ST)Stocktaking: temporarily unavailable'' WHERE DICT_SYSTEM = ''ST''  AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''OR'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(OP)Out of print: (to be) replaced by new edition'',''OR'',''OR'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(OP)Out of print: (to be) replaced by new edition'' WHERE DICT_SYSTEM = ''OR''  AND DICT_TYPE = @ID

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''RE'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''(RE)Awaiting reissue'',''RE'',''RE'',1)
	ELSE
		UPDATE [dbo].tDICT SET DICT_DESC = ''(RE)Awaiting reissue'' WHERE DICT_SYSTEM = ''RE''  AND DICT_TYPE = @ID
--------------------------------------
--Set up order action codes
	SELECT @ID = [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''OA''

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''2'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Cancelled'',''2'',''2'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''3'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Change requested by the supplier'',''3'',''3'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''4'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''No action'',''4'',''4'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''5'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Accepted without amendment'',''5'',''5'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''24'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Accepted with change'',''24'',''24'',1)

	SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''10'' AND DICT_TYPE = @ID
	IF @@ROWCOUNT = 0
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Not found: the supplier has no record of the order line'',''10'',''10'',1)
--------------------------------------

--SET UP CUSTOMER TYPES
SELECT @ID = [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''CT''
--Set up book club type
SELECT @BC_ID =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''bc'')
IF COALESCE(@BC_ID,0) = 0 
BEGIN
	INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System) VALUES (@ID,''Book club'',''BCsys'',''bc'')
	SELECT @BC_ID = ( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''bc'')
END
--Set up Loyalty customer type
SELECT @LOYCT1_ID =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''L1'')
IF COALESCE(@LOYCT1_ID,0) = 0 
BEGIN
	INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System,DICT_ACTIVE) VALUES (@ID,''*Loyalty club 1'',''LOY1sys'',''L1'',1)
	SELECT @LOYCT1_ID = ( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''L1'')
END

--Set up Business customer type type
SELECT @LOYCT1_ID =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''BUS'')
IF COALESCE(@LOYCT1_ID,0) = 0 
BEGIN
	INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System,DICT_ACTIVE) VALUES (@ID,''Business'',''Bus'',''BUS'',0)
	SELECT @LOYCT1_ID = ( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''BUS'')
END

--Set up Private customer type type
SELECT @LOYCT1_ID =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''PRV'')
IF COALESCE(@LOYCT1_ID,0) = 0 
BEGIN
	INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System,DICT_ACTIVE) VALUES (@ID,''Private'',''Prv.'',''PRV'',0)
	SELECT @LOYCT1_ID = ( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''PRV'')
END

--Set up Account customerType
SELECT @LOYCT1_ID =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''acc'')
IF COALESCE(@LOYCT1_ID,0) = 0 
BEGIN
	INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System,DICT_ACTIVE) VALUES (@ID,''Account'',''acc.'',''acc'',1)
	SELECT @LOYCT1_ID = ( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''acc'')
END

--Set up Cash customerType
SELECT @LOYCT1_ID =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''cas'')
IF COALESCE(@LOYCT1_ID,0) = 0 
BEGIN
	INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System,DICT_ACTIVE) VALUES (@ID,''Private'',''cas.'',''cas'',0)
	SELECT @LOYCT1_ID = ( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''cas'')
END

--Set up unallocated customer type type
SELECT @LOYCT1_ID =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''un'')
IF COALESCE(@LOYCT1_ID,0) = 0 
BEGIN
	INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System,DICT_ACTIVE) VALUES (@ID,''Unallocated'',''UN.'',''un'',0)
	SELECT @LOYCT1_ID = ( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''PRV'')
END

--Set up Category for Web export

SELECT @ID = ID FROM tDICTTYPES WHERE [SYSTEM] = ''SE''
SELECT @WEB_ID =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''WEB'')
IF COALESCE(@WEB_ID,0) = 0 
BEGIN
	INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System,DICT_ACTIVE) VALUES (@ID,''For Web export'',''Web'',''WEB'',1)
	SELECT @WEB_ID = ( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''WEB'')
END

IF dbo.GetProperty(''SUPPORTSMULTIBUY'') = ''TRUE''
BEGIN
	--Set up Category for Multibuy catchall
	SELECT @ID = ID FROM tDICTTYPES WHERE [SYSTEM] = ''SE''
	SELECT @MBC_ID =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''MBC'')
	IF COALESCE(@MBC_ID,0) = 0 
	BEGIN
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System,DICT_ACTIVE) VALUES (@ID,''Multibuys'',''MB'',''MBC'',1)
		SELECT @MBC_ID = ( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''MBC'')
	END

	--insert rows to tProductSection table for all multibuy products where there is not such a record already
	INSERT INTO tPRODUCTSECTION (PSEC_P_ID,PSEC_SEC_ID,PSEC_Priority)
	SELECT P_ID,@MBC_ID,0 FROM tPRODUCT a LEFT JOIN 
	(SELECT PSEC_P_ID  FROM tPRODUCTSECTION WHERE PSEC_SEC_ID = @MBC_ID) b ON a.P_ID= b.PSEC_P_ID  
	WHERE b.PSEC_P_ID IS NULL AND a.P_MultibuyCode > '''' 
	UPDATE tProduct SET P_NDA = 1 FROM tPRODUCT JOIN tProductSection ON P_ID = PSEC_P_ID WHERE PSEC_SEC_ID = @MBC_ID

END

--Insert customer interest groups into tDICT
SELECT @ID = [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''IG''
IF ISNULL((SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''LA''),0) =0
	INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Launches'',''LAU'',''LA'',1)
IF ISNULL((SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''SA''),0) =0
	INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Sales'',''SAL'',''SA'',1)
IF ISNULL((SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''PR''),0) =0
	INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Promotions'',''PRO'',''PR'',1)
IF ISNULL((SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''LL''),0) =0
	INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE) VALUES (@ID,''Literary lunches'',''LIT'',''LL'',1)

IF dbo.GetProperty(''SUPPORTSMULTIBUY'') = ''TRUE''
BEGIN
	SELECT @ID = [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''MB''
	SELECT @MB =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''mb1'')
	IF COALESCE(@MB,0) = 0 
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System,DICT_ACTIVE) VALUES (@ID,''R49.00 or 3 for R99'',''700'',''mb1'',1)
	SELECT @MB =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''mb2'')
	IF COALESCE(@MB,0) = 0 
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System,DICT_ACTIVE) VALUES (@ID,''R39.95 or 3 for R99'',''701'',''mb2'',1)
	SELECT @MB =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''mb3'')
	IF COALESCE(@MB,0) = 0 
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System,DICT_ACTIVE) VALUES (@ID,''R39.95 or 4 for R99'',''702'',''mb3'',1)
	SELECT @MB =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''mb4'')
	IF COALESCE(@MB,0) = 0 
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System,DICT_ACTIVE) VALUES (@ID,''R39.00 or 3 for R99'',''703'',''mb4'',1)
	SELECT @MB =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''mb5'')
	IF COALESCE(@MB,0) = 0 
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System,DICT_ACTIVE) VALUES (@ID,''R29.95 or 5 for R99 (romance)'',''704'',''mb5'',1)
	SELECT @MB =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''mb6'')
	IF COALESCE(@MB,0) = 0 
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System,DICT_ACTIVE) VALUES (@ID,''R29 or 5 for R99 (kids)'',''705'',''mb6'',1)
	SELECT @MB =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''mb7'')
	IF COALESCE(@MB,0) = 0 
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System,DICT_ACTIVE) VALUES (@ID,''R19.95 or 6 for R99'',''706'',''mb7'',1)
	SELECT @MB =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''mb8'')
	IF COALESCE(@MB,0) = 0 
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System,DICT_ACTIVE) VALUES (@ID,''R35.00 or 4 for R99'',''707'',''mb8'',1)
	SELECT @MB =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''mb9'')
	IF COALESCE(@MB,0) = 0 
		INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_Short,DICT_System,DICT_ACTIVE) VALUES (@ID,''R79.00 or 2 for R99'',''708'',''mb9'',1)

END


--Set up unallocatedproduct type
SELECT PT_ID FROM tPT WHERE PT_SYSTEM = ''un''
IF @@ROWCOUNT = 0
	INSERT INTO tPT (PT_CODE,PT_Active,PT_System,PT_NUMBER) VALUES (''*UNALLOCATED'',1,''un'',''X'')

SELECT @ID = [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''CT''
UPDATE tCONFIGURATION SET CF_BC_ID =  (SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''bc'')

SELECT @UNCT_ID = DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''un''
UPDATE tCONFIGURATION SET CF_UNALLOCATEDCT = @UNCT_ID--(SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @IG AND DICT_SYSTEM = ''un'')
UPDATE tCONFIGURATION SET CF_LOYALTYCLUBTYPE =  (SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''L1'')
UPDATE tCONFIGURATION SET CF_SectionDictID =  (SELECT [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''SE'')


UPDATE tCONFIGURATION SET CF_BusinessCustomerTypeID = (SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''BUS'')
UPDATE tCONFIGURATION SET CF_PrivateCustomerTypeID = (SELECT DICT_ID FROM tDICT WHERE DICT_SYSTEM = ''PRV'')
UPDATE tCONFIGURATION SET CF_LoyaltyClubType = (SELECT DICT_ID FROM tDICT WHERE [DICT_SYSTEM] = ''L1'')

UPDATE tCONFIGURATION SET CF_CustomerTypeDictID = (SELECT [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''CT'')
UPDATE tCONFIGURATION SET CF_CustomerIGDictID = (SELECT [ID] FROM tDICTTYPES WHERE [SYSTEM] = ''IG'')

--Set up CASH SALES ACCOUNT for POS operations
--Look for Existence of Cash sales account matching CF_CSCustomerID  - skip if found
SELECT TP_NAME FROM tTP WHERE TP_ID = (SELECT ISNULL(CF_CSCUSTOMERID,0) FROM tCONFIGURATION)
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tTP (TP_NAME,TP_ACNO,TP_ROLE,TP_CT_ID,TP_SYSTEM,TP_CanBeDeletedYN,TP_OnMailList,TP_VATABLE,
	TP_CUSTNOTIFYBOOKSALE,TP_CustNotifyBookPromotion,TP_CustNotifyBookLaunch,TP_BALANCE,TP_CREDITLIMIT,TP_TERMS,
	TP_BALANCE_CUR,TP_BALANCE_30,TP_BALANCE_60,TP_BALANCE_90,TP_BALANCE_120PLUS) VALUES (''Cash Sales (POS)'',''CASHPOS'',3,@UNCT_ID,''CS1'',0,0,1,0,0,0,0,0,0,0,0,0,0,0)
END
UPDATE tCONFIGURATION SET CF_CSCustomerID = (SELECT TP_ID FROM tTP WHERE [TP_SYSTEM] = ''CS1'')

--Set up customer with a/c no ''CASUAL''
SELECT TP_NAME FROM tTP WHERE TP_ID = (SELECT ISNULL(CF_CASUALCUSTOMERID,0) FROM tCONFIGURATION)
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tTP (TP_NAME,TP_ACNO,TP_ROLE,TP_CT_ID,TP_SYSTEM,TP_CanBeDeletedYN,TP_OnMailList,TP_VATABLE,
	TP_CUSTNOTIFYBOOKSALE,TP_CustNotifyBookPromotion,TP_CustNotifyBookLaunch,TP_BALANCE,TP_CREDITLIMIT,TP_TERMS,
	TP_BALANCE_CUR,TP_BALANCE_30,TP_BALANCE_60,TP_BALANCE_90,TP_BALANCE_120PLUS) VALUES (''Casual customers'',''CASUAL'',3,0,''CAS'',0,0,1,0,0,0,0,0,0,0,0,0,0,0)
END
UPDATE tCONFIGURATION SET CF_CasualCustomerID = (SELECT TP_ID FROM tTP WHERE [TP_SYSTEM] = ''CAS'')


DECLARE @PT_ID INT
SELECT @PT_ID = PT_ID FROM  tPT WHERE [PT_SYSTEM] = ''un''
UPDATE tPRODUCT SET P_PRODUCTTYPE_ID = @PT_ID WHERE COALESCE(P_PRODUCTTYPE_ID,0) = 0

UPDATE tCONFIGURATION SET CF_UNALLOCATEDPT =  (SELECT PT_ID FROM tPT WHERE PT_CODE = ''*UNALLOCATED'' OR PT_SYSTEM = ''un'')



alter table [dbo].[tTP] disable trigger [UNIQUECODETP1]
UPDATE tTP SET TP_CT_ID = @UNCT_ID FROM tTP LEFT OUTER JOIN tDICT on TP_CT_ID = DICT_ID  WHERE  COALESCE (DICT_ID,0) = 0 
alter table [dbo].[tTP] enable trigger [UNIQUECODETP1]')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.PBKS_INITIALIZEDATA Procedure Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.PBKS_INITIALIZEDATA Procedure'
END
GO

--
-- Script To Update dbo.UpdateProductRecsFromPO Procedure In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.UpdateProductRecsFromPO Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER PROCEDURE [dbo].[UpdateProductRecsFromPO] @POID INT,@FOREIGN BIT,@SUPPLIERID INT,@STATUS INT,@ERR INT OUTPUT,@ERRPOS VARCHAR(10) OUTPUT
AS
BEGIN
	SET NOCOUNT ON;

	BEGIN TRY
		IF @STATUS NOT IN (6,2) --  NOT CANCELLED or in process
		BEGIN
			UPDATE tPRODUCT SET P_LastQtyFirmOrdered = QTYFIRM,
								P_LastQtySSOrdered = QTYSS,
								P_QtyOnOrder = TOTOS,
								P_LastDateOrdered = FLOOR(CONVERT(FLOAT, GetDate())),
								P_LastPriceOrdered=PRICE,
								P_SUPPLIERID = @SUPPLIERID,
								P_DEALID = DEALID,
								P_ProductType_ID = CASE PRODUCTTYPE WHEN 0 THEN P_ProductType_ID ELSE PRODUCTTYPE END,
								P_RRP = CASE @FOREIGN WHEN 0 THEN PRICE ELSE P_RRP END
			FROM tPRODUCT JOIN vPOLs_Aggr ON PID = P_ID WHERE POID =  @POID
		END
		ELSE 
		IF @STATUS = 6 -- CANCELLED
			BEGIN
				UPDATE tPRODUCT SET P_LASTDATEORDERED = v.Dte,
									P_LastQtyFirmOrdered = v.LastQtyFirm,
									P_LastQtySSOrdered = v.LastQtySS,
									P_QtyOnOrder = v.TOTOS,
									P_LastPriceOrdered=v.PRICE
				FROM tPRODUCT a JOIN vLastPOLsperPO v ON a.P_ID = v.PID WHERE POID = @POID
			END
			ELSE
			IF @STATUS = 2  -- IN PROCESS
			BEGIN
				UPDATE tPRODUCT SET P_QtyOnOrder_Unissued = v.OO
				FROM tPRODUCT a JOIN vPOLSInProcessPerPID v ON a.P_ID = v.PID WHERE POID = @POID
			END
	END TRY
	BEGIN CATCH
		DECLARE @ErrorString NVARCHAR(4000)
		DECLARE @TMPPARAMETER VARCHAR(MAX)
		IF ERROR_PROCEDURE() <> OBJECT_NAME(@@PROCID)
			SELECT @ErrorString = ''Catch 1: '' + ERROR_MESSAGE();
		ELSE
			SELECT @ErrorString = ''Catch 2: Error Message: '' + ERROR_MESSAGE() + '', Error Procedure: '' + RTRIM(CONVERT(CHAR,ERROR_PROCEDURE())) + 
					'', Error Line: '' + RTRIM(CONVERT(Char,ERROR_LINE ()));

		if (XACT_STATE()) = -1
		begin
			rollback transaction;
			SELECT @TMPPARAMETER = ''Catch 1 (ROLLBACK): '' + @ErrorString
			EXEC dbo.SAVELOG @TMPPARAMETER, ''[PostPO]''
		end;
		-- Test whether the transaction is active and valid.
		if (XACT_STATE()) = 1
		BEGIN
			COMMIT TRANSACTION
			SELECT @TMPPARAMETER = ''Catch 2: (COMMIT)'' + @ErrorString
			EXEC dbo.SAVELOG @ErrorString,''[PostPO]''
		END
		RAISERROR (@ErrorString, 16,1)

	END CATCH
END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.UpdateProductRecsFromPO Procedure Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.UpdateProductRecsFromPO Procedure'
END
GO

--
-- Script To Update dbo.PostPO Procedure In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.PostPO Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
ALTER  PROCEDURE [dbo].[PostPO] @POID INT,@Status INT,@FOREIGN BIT,@SUPPLIERID INT,@ERR INT OUTPUT ,@ERRPOS VARCHAR(10) OUTPUT

AS
BEGIN
DECLARE @ERR2 INT
DECLARE @ERRPOS2 VARCHAR(10)

BEGIN TRY

	UPDATE tTR SET TR_STATUS = @STATUS WHERE TR_ID = @POID

	IF @STATUS = 2 
		EXECUTE UpdateProductRecsFromPO @POID,@FOREIGN,@SUPPLIERID,@STATUS,@ERR2 OUTPUT,@ERRPOS2 OUTPUT
	ELSE
	BEGIN
		BEGIN TRANSACTION
		EXECUTE UpdateProductRecsFromPO @POID,@FOREIGN,@SUPPLIERID,@STATUS,@ERR2 OUTPUT,@ERRPOS2 OUTPUT

		IF @STATUS = 6
			UPDATE tCOL SET COL_ACTIONTAKEN = 0 
			FROM dbo.tCOL INNER JOIN
				 dbo.tPOL ON dbo.tCOL.COL_P_ID = dbo.tPOL.POL_P_ID WHERE POL_TR_ID = @POID AND COL_ACTIONTAKEN = 1

		UPDATE tCOL SET COL_ACTIONTAKEN = 1 
		FROM         dbo.tCOL INNER JOIN
						  dbo.tPOL ON dbo.tCOL.COL_P_ID = dbo.tPOL.POL_P_ID INNER JOIN
						  dbo.tTR ON dbo.tCOL.COL_TR_ID = dbo.tTR.TR_ID 
		WHERE     (dbo.tCOL.COL_ActionTaken IN (0, 2)) AND (dbo.tTR.TR_Status = 3) AND POL_TR_ID = @POID
		COMMIT TRANSACTION
	END
END TRY
BEGIN CATCH
DECLARE @ErrorString NVARCHAR(4000)
DECLARE @TMPPARAMETER VARCHAR(MAX)
		IF ERROR_PROCEDURE() <> OBJECT_NAME(@@PROCID)
			SELECT @ErrorString = ''Catch 1: '' + ERROR_MESSAGE();
		ELSE
			SELECT @ErrorString = ''Catch 2: Error Message: '' + ERROR_MESSAGE() + '', Error Procedure: '' + RTRIM(CONVERT(CHAR,ERROR_PROCEDURE())) + 
					'', Error Line: '' + RTRIM(CONVERT(Char,ERROR_LINE ()));

		if (XACT_STATE()) = -1
		begin
			rollback transaction;
			SELECT @TMPPARAMETER = ''Catch 1 (ROLLBACK): '' + @ErrorString
			EXEC dbo.SAVELOG @TMPPARAMETER, ''[PostPO]''
		end;
		-- Test whether the transaction is active and valid.
		if (XACT_STATE()) = 1
		BEGIN
			COMMIT TRANSACTION
			SELECT @TMPPARAMETER = ''Catch 2: (COMMIT)'' + @ErrorString
			EXEC dbo.SAVELOG @ErrorString,''[PostPO]''
		END
		RAISERROR (@ErrorString, 16,1)

END CATCH

END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.PostPO Procedure Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.PostPO Procedure'
END
GO

--
-- Script To Update dbo.ReorderBrowsed Procedure In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.ReorderBrowsed Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER PROCEDURE [dbo].[ReorderBrowsed] @Slatename VARCHAR(50),@WSNAME VARCHAR(50),@FILTER_OO BIT = 0,@FILTER_OH BIT = 0,@SUPPLIERID INT = 0 AS
DELETE FROM  tREORDERGENERAL WHERE SLATENAME = ''Browsed'' AND WSNAME = Host_Name()
INSERT  INTO dbo.tREORDERGENERAL  (SLATENAME,WSNAME,STATUS,COLID,REF,QTYFIRM,QTYSS,PID,QTYCO,QTYPO,QTYPOUNISSUED,QTYAPP,PRCODE,DESCRIP,LASTSUPPLIERID,LASTSUPPLIERNAME,LASTDEALID,LASTDEALNAME,PT,
PUBLISHER,TOTALSOLD,PRICE,ONHAND,LASTRECEIVEDDATE,LASTORDEREDDATE,LASTRECEIVEDQTY,LASTORDEREDQTYFIRM,
LASTORDEREDQTYSS,LASTRECEIVEDPRICE ,LASTORDEREDPRICE,LASTSIXWEEKS,LASTSIXMONTHS,TitleForSort,FOreignPrice,Category) 
SELECT  DISTINCT
LEFT(ISNULL(@Slatename,''''),50),
LEFT(ISNULL(@WSNAME,''''),50),
'''',
0,
'''',
0,0,
P_ID,
ISNULL(P_QtyOnBackorder,0),
ISNULL(P_QtyOnOrder,0),
ISNULL(P_QtyOnOrder_Unissued,0),
ISNULL(P_QtyOnAPPro,0),
LEFT(ISNULL(P_EAN,''''),20),
--LEFT(P_Title,35) + '': '' + Left(P_MainAuthor,12),
LEFT(ISNULL(Description,''''),120) + '': '' + Left(ISNULL(P_MainAuthor,''''),78),
P_SupplierID,
Left(ISNULL(SupplierName,''''),55),
DefaultDeal,
Left(ISNULL(DL_Description,''''),40),
LEFT(ISNULL(PT,''''),50),
LEFT(ISNULL(P_Publisher,''''),50),
ISNULL(Qty,0),
ISNULL(P_RRP,0),
ISNULL(P_QTyOnHand,0),
P_LastDateDelivered,
P_LastDateOrdered,
ISNULL(P_LastQtyDelivered,0),
ISNULL(P_LastQtyFirmOrdered,0),
ISNULL(P_LastQtySSOrdered,0),
ISNULL(P_LastPriceDelivered,0),
ISNULL(P_LastPriceOrdered,0),
LEFT(ISNULL(LastSixWeeks,''''),60),
LEFT(ISNULL(LastSixMOnths,''''),60),
Left(ISNULL(TitleForSort,''''),50),
		ForeignPrice,
Left(Sections,30)
FROM vReorder_Browsed 

--The following code sets the deal to the only deal available on the supplier record where there is only one deal
UPDATE tREORDERGENERAL SET LastDealID = DL_ID FROM tREORDERGENERAL a JOIN 
	tDEAL b 
					ON a.LastSupplierID = b.DL_TP_ID JOIN 
	(SELECT COUNT(DL_ID) as CNT,DL_TP_ID FROM tDEAL GROUP BY DL_TP_ID) c 
					ON a.LastSUpplierID = c.DL_TP_ID

	WHERE c.CNT = 1	AND ISNULL(a.LastDealID,0) = 0')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.ReorderBrowsed Procedure Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.ReorderBrowsed Procedure'
END
GO

--
-- Script To Update dbo.REORDERCUST Procedure In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.REORDERCUST Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER PROCEDURE [dbo].[REORDERCUST] @SLATENAME VARCHAR(50),@WSNAME VARCHAR(50), @pDate DATETIME, @SUPPLIERID INT = 0,@ExclActioned BIT,@FILTER_OO BIT = 0,@FILTER_OH BIT = 0 AS
	
INSERT  INTO dbo.tREORDERGENERAL  (SLATENAME,WSNAME,STATUS,COLID,REF,QTYFIRM,QTYSS,PID,QTYCO,QTYPO,QTYPOUNISSUED,QTYAPP,PRCODE,DESCRIP,LASTSUPPLIERID,LASTSUPPLIERNAME,LASTDEALID,				LASTDEALNAME,PT,PUBLISHER,TOTALSOLD,PRICE,ONHAND,LASTRECEIVEDDATE,LASTORDEREDDATE,LASTRECEIVEDQTY,LASTORDEREDQTYFIRM,
		LASTORDEREDQTYSS,LASTRECEIVEDPRICE ,LASTORDEREDPRICE,TItleForSort,FOreignPrice,Category) 
		SELECT 
	LEFT(@SLATENAME,50),
	LEFT(@WSNAME,50),
		''C'',
		0,
		'''',
		0,0,
		P_ID,
		CAST(ISNULL(P_QtyOnBackorder,0) AS VARCHAR(15)),
		ISNULL(P_QtyOnOrder,0),
		ISNULL(P_QtyOnOrder_Unissued,0),
		ISNULL(P_QtyOnAppro,0),
	LEFT(ISNULL(P_EAN,''''),20),
		LEFT(ISNULL(Description,''''),120) + '': '' + Left(ISNULL(P_MainAuthor,''''),78),
		P_SupplierID,
		Left(ISNULL(SupplierName,''''),55),
		DEFAULTDEAL,
		Left(ISNULL(DL_Description,''''),40),
		LEFT(ISNULL(PT,''''),50),
		LEFT(ISNULL(P_Publisher,''''),50),
		ISNULL(P_QtyTotalSold,0),
		ISNULL(P_RRP,0),
		ISNULL(P_QTyOnHand,0),
		P_LastDateDelivered,
		P_LastDateOrdered,
		ISNULL(P_LastQtyDelivered,0),
		ISNULL(P_LastQtyFirmOrdered,0),
		ISNULL(P_LastQtySSOrdered,0),
		ISNULL(P_LastPriceDelivered,0),
		ISNULL(P_LastPriceOrdered,0),
		LEFT(ISNULL(P_TITLE,''''),50),
		ForeignPrice,
Left(ISNULL(Sections,''''),30)
		FROM vReorderCust
		WHERE ISNULL(P_SUPPLIERID,0) = CASE WHEN @SUPPLIERID = 0 THEN ISNULL(P_SUPPLIERID,0) ELSE @SUPPLIERID END AND ACTIONTAKEN Not in (3,1,9)
	AND ((@ExclActioned <> 0 AND ACTIONTAKEN NOT IN (3,1,9)) OR (@ExclActioned = 0 AND ACTIONTAKEN NOT IN (3)))
AND ISNULL(P_QTYONORDER,0) =  CASE WHEN @FILTER_OO = 1 THEN 0 ELSE ISNULL(P_QTYONORDER,0) END
AND ISNULL(P_QTYONHAND,0) =  CASE WHEN @FILTER_OH = 1 THEN 0 ELSE ISNULL(P_QTYONHAND,0)  END

--The following code sets the deal to the only deal available on the supplier record where there is only one deal
UPDATE tREORDERGENERAL SET LastDealID = DL_ID FROM tREORDERGENERAL a JOIN 
	tDEAL b 
					ON a.LastSupplierID = b.DL_TP_ID JOIN 
	(SELECT COUNT(DL_ID) as CNT,DL_TP_ID FROM tDEAL GROUP BY DL_TP_ID) c 
					ON a.LastSUpplierID = c.DL_TP_ID

	WHERE c.CNT = 1	AND ISNULL(a.LastDealID,0) = 0')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.REORDERCUST Procedure Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.REORDERCUST Procedure'
END
GO

--
-- Script To Update dbo.ReorderCust_ByCOL Procedure In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.ReorderCust_ByCOL Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER PROCEDURE [dbo].[ReorderCust_ByCOL] (@SLATENAME VARCHAR(50),@WSNAME VARCHAR(50),@STAFFID INT,@SUPPLIERID INT = 0,@ExclActioned BIT,@FILTER_OO BIT = 0,@FILTER_OH BIT = 0) AS
	
	If @STAFFID = 0 
		DELETE FROM tREORDERCUSTByCOL WHERE WSNAME = @WSNAME
	else
		DELETE FROM tREORDERCUSTByCOL WHERE STAFFID = @StaffID AND WSNAME = @WSNAME
	
	
INSERT  INTO dbo.tREORDERCUSTByCOL  (SLATENAME,WSNAME,STATUS,CODate,COLID,REF,QTYFIRM,QTYSS,PID,QTYCO,QTYPO,QtyPOUnissued,QTYAPP,PRCODE,DESCRIP,
	LASTSUPPLIERID,LASTSUPPLIERNAME,LASTDEALID,LASTDEALNAME,PT,
	PUBLISHER,TOTALSOLD,PRICE,ONHAND,LASTRECEIVEDDATE,LASTORDEREDDATE,LASTRECEIVEDQTY,LASTORDEREDQTYFIRM,
	LASTORDEREDQTYSS,LASTRECEIVEDPRICE ,LASTORDEREDPRICE,STAFFID,LASTSIXWEEKS,LASTSIXMONTHS,TITLEFORSORT,FOreignPrice,Category) 
	SELECT 
	LEFT(@SLATENAME,50),
	LEFT(@WSNAME,50),
	''C'',
	a.TR_DATE, 
	COL_ID,
	LEFT(ISNULL(a.TP_ACNO,'''') + '': '' + RTRIM(ISNULL(a.COL_REF,'''')),35) ,
	0,0,
	a.P_ID,
	LEFT(ISNULL(CAST(a.COL_QTY as VARCHAR),'''') + '';'' + ISNULL(CAST(ISNULL(a.COL_QTYDISPATCHED,0) as VARCHAR),''''),16) ,
	ISNULL(a.P_QtyOnOrder,0),
	ISNULL(a.P_QtyOnOrder_Unissued,0),
	ISNULL(a.P_QtyOnAPPRO,0),
	LEFT(ISNULL(a.P_EAN,''''),20),
	LEFT(ISNULL(a.P_Title,''''),120) + '': '' + Left(ISNULL(a.P_MainAuthor,''''),78),
	a.P_SupplierID,
	Left(ISNULL(SupplierName,''''),55),
	DefaultDeal,
	Left(ISNULL(DL_Description,''''),40),
	LEFT(ISNULL(PT,''''),50),
	LEFT(ISNULL(P_Publisher,''''),50),
	ISNULL(P_QtyTotalSold,0),
	ISNULL(P_RRP,0),
	ISNULL(P_QTyOnHand,0),
	P_LastDateDelivered,
	P_LastDateOrdered,
	ISNULL(P_LastQtyDelivered,0),
	ISNULL(P_LastQtyFirmOrdered,0),
	ISNULL(P_LastQtySSOrdered,0),
	ISNULL(P_LastPriceDelivered,0),
	ISNULL(P_LastPriceOrdered,0),
	TR_STAFFID,
	LEFT(ISNULL(a.LASTSIXWEEKS,''''),60),
	LEFT(ISNULL(a.LASTSIXMONTHS,''''),60),
	LEFT(ISNULL(P_TITLE,''''),50),
	ForeignPrice,
Left(ISNULL(Sections,''''),30)
	FROM vREORDERCUSTByCOL a WHERE CASE WHEN @StaffID = 0 THEN @StaffID ELSE a.TR_STAFFID END = @StaffID 
	AND ISNULL(a.P_SUPPLIERID,0) = CASE WHEN @SUPPLIERID = 0 THEN ISNULL(a.P_SUPPLIERID,0) ELSE @SUPPLIERID END 
	AND ((@ExclActioned <> 0 AND ACTIONTAKEN NOT IN (3,1,9)) OR (@ExclActioned = 0 AND ACTIONTAKEN NOT IN (3)))
AND ISNULL(P_QTYONORDER,0) =  CASE WHEN @FILTER_OO = 1 THEN 0 ELSE ISNULL(P_QTYONORDER,0) END
AND ISNULL(P_QTYONHAND,0) =  CASE WHEN @FILTER_OH = 1 THEN 0 ELSE ISNULL(P_QTYONHAND,0)  END

--The following code sets the deal to the only deal available on the supplier record where there is only one deal
UPDATE tREORDERCUSTByCOL SET LastDealID = DL_ID FROM tREORDERCUSTByCOL a JOIN 
	tDEAL b 
					ON a.LastSupplierID = b.DL_TP_ID JOIN 
	(SELECT COUNT(DL_ID) as CNT,DL_TP_ID FROM tDEAL GROUP BY DL_TP_ID) c 
					ON a.LastSUpplierID = c.DL_TP_ID

	WHERE c.CNT = 1	AND ISNULL(a.LastDealID,0) = 0')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.ReorderCust_ByCOL Procedure Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.ReorderCust_ByCOL Procedure'
END
GO

--
-- Script To Update dbo.ReorderSales Procedure In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.ReorderSales Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER PROCEDURE [dbo].[ReorderSales] @Slatename VARCHAR(50),@WSNAME VARCHAR(50),@pDate DATETIME,@FILTER_OO BIT = 0,@FILTER_OH BIT = 0,@SUPPLIERID INT = 0 AS

INSERT  INTO dbo.tREORDERGENERAL  (SLATENAME,WSNAME,STATUS,COLID,REF,QTYFIRM,QTYSS,PID,QTYCO,QTYPO,QTYPOUNISSUED,QTYAPP,PRCODE,DESCRIP,LASTSUPPLIERID,LASTSUPPLIERNAME,LASTDEALID,LASTDEALNAME,PT,
PUBLISHER,TOTALSOLD,PRICE,ONHAND,LASTRECEIVEDDATE,LASTORDEREDDATE,LASTRECEIVEDQTY,LASTORDEREDQTYFIRM,
LASTORDEREDQTYSS,LASTRECEIVEDPRICE ,LASTORDEREDPRICE,LASTSIXWEEKS,LASTSIXMONTHS,TitleForSort,Category) 
SELECT  DISTINCT
LEFT(ISNULL(@Slatename,''''),50),
LEFT(ISNULL(@WSNAME,''''),50),
'''',
0,
'''',
0,0,
P_ID,
ISNULL(P_QtyOnBackorder,0),
ISNULL(P_QtyOnOrder,0),
ISNULL(P_QtyOnOrder_Unissued,0),
ISNULL(P_QtyOnAPPro,0),
LEFT(ISNULL(P_EAN,''''),20),
--LEFT(P_Title,35) + '': '' + Left(P_MainAuthor,12),
LEFT(ISNULL(Description,''''),120) + '': '' + Left(ISNULL(P_MainAuthor,''''),78),
P_SupplierID,
Left(ISNULL(SupplierName,''''),55),
DefaultDeal,
Left(ISNULL(DL_Description,''''),40),
LEFT(ISNULL(PT,''''),50),
LEFT(ISNULL(P_Publisher,''''),50),
ISNULL(Qty,0),
ISNULL(P_RRP,0),
ISNULL(P_QTyOnHand,0),
P_LastDateDelivered,
P_LastDateOrdered,
ISNULL(P_LastQtyDelivered,0),
ISNULL(P_LastQtyFirmOrdered,0),
ISNULL(P_LastQtySSOrdered,0),
ISNULL(P_LastPriceDelivered,0),
ISNULL(P_LastPriceOrdered,0),
LEFT(ISNULL(LastSixWeeks,''''),60),
LEFT(ISNULL(LastSixMOnths,''''),60),
Left(ISNULL(TitleForSort,''''),50),
Left(ISNULL(Sections,''''),30)
FROM zReorder_SalesandTransfers WHERE DTE > @pDATE

AND ISNULL(P_QTYONORDER,0) =  CASE WHEN @FILTER_OO = 1 THEN 0 ELSE ISNULL(P_QTYONORDER,0) END
AND ISNULL(P_QTYONHAND,0) =  CASE WHEN @FILTER_OH = 1 THEN 0 ELSE ISNULL(P_QTYONHAND,0)  END
AND ISNULL(P_SUPPLIERID,0) = CASE WHEN @SUPPLIERID = 0 THEN ISNULL(P_SUPPLIERID,0) ELSE @SUPPLIERID END

--The following code sets the deal to the only deal available on the supplier record where there is only one deal
UPDATE tREORDERGENERAL SET LastDealID = DL_ID FROM tREORDERGENERAL a JOIN 
	tDEAL b 
					ON a.LastSupplierID = b.DL_TP_ID JOIN 
	(SELECT COUNT(DL_ID) as CNT,DL_TP_ID FROM tDEAL GROUP BY DL_TP_ID) c 
					ON a.LastSUpplierID = c.DL_TP_ID

	WHERE c.CNT = 1	AND ISNULL(a.LastDealID,0) = 0')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.ReorderSales Procedure Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.ReorderSales Procedure'
END
GO

--
-- Script To Create dbo.SaveExportToPastel Procedure In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Creating dbo.SaveExportToPastel Procedure'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('CREATE PROCEDURE [dbo].[SaveExportToPastel] (@Type VARCHAR(2))
AS
BEGIN
DECLARE @ID INT

	INSERT INTO tExportToAccountingMaster(ExportDate,ExportDebtorsOrCreditors) VALUES (GetDate(),@Type)
	SELECT @ID = SCOPE_IDENTITY()
	INSERT INTO tExportToAccountingLog(
	FKEY,
	[Period],
	[TransactionNominalDate],
	[GDC],
	[Acno],
	[Reference],
	[Description],
	[Amount],
	[TaxType],
	[TaxAmount],
	[Openitem],
	[Costcode],
	[ContraAccount],
	[ExchangeRate],
	[BankExchangeRate],
	[BatchID],
	[DiscountTax],
	[DiscountAmount],
	[HomeAmount],
	[TRGLOBlobalID],
	[SignedDate])

	SELECT @ID,[PERIOD],[DTE],[GDC],[ACNO],[REFERENCE],[DESCR],[AMT],[TAXTYPE],[TAXAMT],
	[OPENITEM],[COSTCODE],[CONTRAACCOUNT],[EXCHANGERATE],[BANKEXCHANGERATE],[BATCHID],
	[DISCOUNTTAX],[DISCOUNTAMT],[HOMEAMT],[TRGLOBALID],[ProcessingDate] FROM tPASTEL WHERE [ACTION] <> 0



END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.SaveExportToPastel Procedure Added Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Add dbo.SaveExportToPastel Procedure'
END
GO

--
-- Script To Update dbo.trigCOUpdate Trigger In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.trigCOUpdate Trigger'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER TRIGGER [dbo].[trigCOUpdate] ON [dbo].[tCOL]
FOR INSERT, UPDATE
AS
	
	IF dbo.GETPROPERTY(''POSACTIVE'') = ''TRUE''
	INSERT INTO tCOUpdate
		(COU_COLID,
		COU_TRID,
		COU_TPID,
		COU_DATE,
		COU_CODE,
		COU_PID,
		COU_QTY,
		COU_QTYDISPATCHED,
		COU_PRICE,
		COU_DISCOUNTRATE,
		COU_DEPOSIT,
		COU_DEPOSITSTATUS,
		COU_REF,
		COU_TriggerDate,
		COU_DOCSTATUS,
		COU_TRSTATUS)
		SELECT ins.COL_ID,
			ins.COL_TR_ID,
			tTP.TP_ID,
			tTR.TR_DATE,
			tTR.TR_CODE,
			ins.COL_P_ID,
			ins.COL_QTY,
			ins.COL_QTYDISPATCHED,
			ins.COL_PRICE,
			ins.COL_DISCOUNTPERCENT,
			ins.COL_DEPOSIT,
			ins.COL_DEPOSITSTATUS,
			ins.COL_REF,
			GetDate(),
			dbo.INACTIVECOL(TR_STATUS,ins.COL_FULFILLED),
			tTR.TR_STATUS
		FROM inserted ins JOIN tTR ON COL_TR_ID = TR_ID JOIN tTP ON TR_TP_ID = TP_ID
			Left JOIN deleted del ON ins.COL_ID = del.COL_ID
		WHERE		
					ISNULL(ins.COL_QTY,0) <> ISNULL(del.COL_QTY,0) or
					ISNULL(ins.COL_QTYDISPATCHED,0) <> ISNULL(del.COL_QTYDISPATCHED,0) or
					ISNULL(ins.COL_PRICE,0) <> ISNULL(del.COL_PRICE,0) or
					ISNULL(ins.COL_DISCOUNTPERCENT,0) <> ISNULL(del.COL_DISCOUNTPERCENT,0) or
					ISNULL(ins.COL_DEPOSIT,0) <> ISNULL(del.COL_DEPOSIT,0) or
					ISNULL(ins.COL_DEPOSITSTATUS,'''') <> ISNULL(del.COL_DEPOSITSTATUS,'''') or
					ISNULL(ins.COL_REF,'''') <> ISNULL(del.COL_REF,'''')')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.trigCOUpdate Trigger Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.trigCOUpdate Trigger'
END
GO

--
-- Script To Update dbo.UniqueCodeDICT1 Trigger In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.UniqueCodeDICT1 Trigger'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER TRIGGER dbo.UniqueCodeDICT1 ON [dbo].[tDict]
FOR INSERT, UPDATE
AS
BEGIN
DECLARE @TMP VARCHAR(50)

 IF (SELECT MAX(cnt) FROM (SELECT COUNT(i.DICT_SHORT) as cnt from tDICT,
  inserted i WHERE tDICT.DICT_SHORT=i.DICT_SHORT AND tDICT.DICT_TYPE = i.DICT_TYPE AND i.DICT_SHORT > '''' GROUP BY i.DICT_SHORT)x)>1
	BEGIN

		SELECT @TMP = Max(i.DICT_SHORT) 
		FROM tDICT,	inserted i 
		WHERE tDICT.DICT_SHORT=i.DICT_SHORT AND tDICT.DICT_TYPE = i.DICT_TYPE AND i.DICT_SHORT > '''' 
		GROUP BY i.DICT_SHORT

		RAISERROR (''DUPLICATE'', 16, 1,@TMP, '''')
	END
End')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.UniqueCodeDICT1 Trigger Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.UniqueCodeDICT1 Trigger'
END
GO

--
-- Script To Update dbo.trig_PriceChange Trigger In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.trig_PriceChange Trigger'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER TRIGGER [dbo].[trig_PriceChange] ON dbo.tProduct
FOR UPDATE
AS
BEGIN

BEGIN TRY

	IF update(P_SP) 
	BEGIN

		INSERT INTO tPriceChange 
		(PCH_PID,PCH_Date,PCH_Price)
		SELECT	ins.P_ID,GetDate(),
			ins.P_SP
		FROM inserted ins Left JOIN DELETED del ON ins.P_ID = del.P_ID
		WHERE ins.P_SP <> ISNULL(del.P_SP,0) AND ins.P_SP IS NOT NULL
		
	END
	SELECT 1
END TRY

BEGIN CATCH

	DECLARE @ErrorString NVARCHAR(4000)

	IF ERROR_PROCEDURE() <> OBJECT_NAME(@@PROCID)
	BEGIN
		PRINT ''X''
		SELECT @ErrorString = ERROR_MESSAGE() + '' X'';
		PRINT @ERRORSTRING
	END
	ELSE
	BEGIN
		PRINT ''Y''
		SELECT @ErrorString = ''Error Message: '' + ERROR_MESSAGE() + 
				'', Error Procedure: '' + RTRIM(CONVERT(CHAR,ERROR_PROCEDURE())) + '', Error Line: '' + RTRIM(CONVERT(Char,ERROR_LINE ()));
		SELECT @ERRORSTRING =  ERROR_MESSAGE() + '' X''
	END
	RAISERROR (@ERRORSTRING, 16,1)
END CATCH
END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.trig_PriceChange Trigger Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.trig_PriceChange Trigger'
END
GO

--
-- Script To Update dbo.trigProdDelete Trigger In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.trigProdDelete Trigger'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER TRIGGER [dbo].[trigProdDelete] ON dbo.tProduct 
FOR DELETE 

		
AS
		INSERT INTO tProdUpdates 
		(PRU_Log_Type,
		PRU_P_ID		)
		SELECT	''DEL'',
			del.P_ID
		FROM Deleted del

IF dbo.GetProperty(''SUPPORTSMULTIBUY'') = ''TRUE''
BEGIN
DECLARE @ID INT
DECLARE @MBC_ID INT
DECLARE @OLDVAL VARCHAR(10)
DECLARE @PID UNiqueidentifier
		SELECT @PID = del.P_ID from deleted del
	--Add to or remove from multibuy category if updated
		SELECT @ID = ID FROM tDICTTYPES WHERE [SYSTEM] = ''SE''
		SELECT @MBC_ID =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''MBC'')
		DELETE FROM tPRODUCTSECTION WHERE PSEC_P_ID = @PID AND PSEC_SEC_ID = @MBC_ID
END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.trigProdDelete Trigger Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.trigProdDelete Trigger'
END
GO

--
-- Script To Update dbo.trigProdUpdate Trigger In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.trigProdUpdate Trigger'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER TRIGGER [dbo].[trigProdUpdate] ON dbo.tProduct
FOR INSERT, UPDATE
AS


		INSERT INTO tProdUpdates 
		(PRU_Log_Type,
		PRU_P_ID,
		PRU_Code,
		PRU_EAN,
		PRU_Publisher,
		PRU_SeriesTitle,
		PRU_MainAuthor,
		PRU_Title,
		PRU_SP,
		PRU_SSP,
		PRU_VATRate,
		PRU_TriggerDate,
		PRU_PTID,
		PRU_SECID,
		PRU_NDA,
		PRU_MULTIBUYCODE)
		SELECT	''NEW'',
			ins.P_ID,
			ins.P_Code,
			ins.P_EAN,
			ins.P_Publisher,
			ins.P_SeriesTitle,
			ins.P_MainAuthor,
			LEFT(ins.P_Title,250),
			ins.P_SP,
			ins.P_SPecial,
			dbo.VATRATETOUSE(ins.P_SpecialVat,ins.P_VatRate),
			GetDate(),
			ins.P_ProductType_ID,
			vSectionMaster.PSEC_SEC_ID,
			ins.P_NDA,
			ins.P_MultibuyCode
		FROM inserted ins LEFT JOIN vSectionMaster ON ins.P_ID = vSectionMaster.PSEC_P_ID 
		Left JOIN deleted del ON ins.P_ID = del.P_ID 
		WHERE		ISNULL(ins.P_CODE,'''') <> ISNULL(del.P_Code,'''') or
					ISNULL(ins.P_EAN,'''') <> ISNULL(del.P_EAN,'''') or
					ISNULL(ins.P_Publisher,'''') <> ISNULL(del.P_Publisher,'''') or
					ISNULL(ins.P_SeriesTitle,'''') <> ISNULL(del.P_SeriesTitle,'''') or
					ISNULL(ins.P_MainAuthor,'''') <> ISNULL(del.P_MainAuthor,'''') or
					ISNULL(ins.P_Title,'''') <> ISNULL(del.P_Title,'''') or
					ISNULL(ins.P_SP,0) <> ISNULL(del.P_SP,0) or
					ISNULL(ins.P_SPecial,0) <> ISNULL(del.P_SPecial,0) or
					ISNULL(ins.P_VatRate,0) <> ISNULL(del.P_VatRate,0) or
					ISNULL(ins.P_SpecialVat,0) <> ISNULL(del.P_SpecialVat,0) or
					ISNULL(ins.P_ProductType_ID,0) <> ISNULL(del.P_ProductType_ID,0) or
					ISNULL(ins.P_NDA,'''') <> ISNULL(del.P_NDA,'''') or
					ISNULL(ins.P_MultibuyCode,'''') <> ISNULL(del.P_MultibuyCode,'''') or
					UPDATE (P_CODE)


IF dbo.GetProperty(''SUPPORTSMULTIBUY'') = ''TRUE''
BEGIN
DECLARE @ID INT
DECLARE @MBC_ID INT
DECLARE @OLDVAL VARCHAR(10)
DECLARE @NEWVAL VARCHAR(10)
DECLARE @PID UNiqueidentifier
		SELECT @OLDVAL = del.P_MultibuyCode from Deleted del
		SELECT @NEWVAL = ins.P_MultibuyCode,@PID = ins.P_ID from inserted ins
	--Add to or remove from multibuy category if updated
		IF ISNULL(@NEWVAL,'''') <> ISNULL(@OLDVAL,'''') 
		BEGIN
			SELECT @ID = ID FROM tDICTTYPES WHERE [SYSTEM] = ''SE''
			SELECT @MBC_ID =( SELECT DICT_ID FROM tDICT WHERE DICT_TYPE = @ID AND DICT_SYSTEM = ''MBC'')
			IF ISNULL(@NEWVAL,'''') > ''''
			BEGIN
				IF	NOT (SELECT COUNT( PSEC_P_ID) FROM tProductSection WHERE PSEC_P_ID = @PID and PSEC_SEC_ID = @MBC_ID) > 0
					INSERT INTO tProductSection (PSEC_P_ID,PSEC_SEC_ID,PSEC_Priority)
					Values	(@PID,@MBC_ID,0) 
				UPDATE tProduct SET P_NDA = 1 WHERE P_ID = @PID
			END
			ELSE
			BEGIN
				UPDATE tProduct SET P_NDA = 0 WHERE P_ID = @PID
				DELETE FROM tPRODUCTSECTION WHERE PSEC_P_ID = @PID AND PSEC_SEC_ID = @MBC_ID
			END
		END		

			


END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.trigProdUpdate Trigger Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.trigProdUpdate Trigger'
END
GO

--
-- Script To Update dbo.trigQtyOHChange Trigger In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.trigQtyOHChange Trigger'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


exec('ALTER TRIGGER [dbo].[trigQtyOHChange] ON dbo.tProduct
AFTER INSERT, UPDATE, DELETE 
AS
BEGIN
DECLARE @QtyOHBody XML
DECLARE @DMLType CHAR(1)	
DECLARE @INSTALLATIONCODE VARCHAR(10)
DECLARE @CMD VARCHAR(200)
DECLARE @TMP VARCHAR(MAX)
DECLARE @RES INT
DECLARE @ERRMESS VARCHAR(500)
DECLARE @SOH INT
DECLARE @SOO INT
DECLARE @LASTDELIVEREDPRICE INT
DECLARE @SP INT
DECLARE @LASTDATEORDERED DATETIME
DECLARE @LASTQTYFIRMORDERED INT
DECLARE @LASTQTYSSORDERED INT

DECLARE @TOTALQTYSOLD INT
DECLARE @LASTDELIVEREDDATE DATETIME
DECLARE @LASTDATESOLD DATETIME
DECLARE @LASTQTYDELIVERED INT

DECLARE @PEAN VARCHAR(20)
DECLARE @PID UNIQUEIDENTIFIER
DECLARE @STORECODE VARCHAR(5)
DECLARE @ErrorString NVARCHAR(4000)

--declare @$prog varchar(50), 
--	@$errno int, 
--	@$errmsg varchar(4000), 
--	@$proc_section_nm varchar(50),
--	@$row_cnt INT,
--	@$error_db_name varchar(50), 
--	@$CreateUserName varchar(128),   -- last user changed the data 
--	@$CreateMachineName varchar(128), -- last machine changes-procedure were run from
--	@$CreateSource varchar(128)		-- last process that made a changes
--
--select @$errno = NULL,  @$errmsg = NULL,  @$proc_section_nm = NULL
--	,  @$prog = LEFT(object_name(@@procid),50), @$row_cnt = NULL
--	, @$error_db_name = db_name();
--

BEGIN TRY
	IF dbo.GETPROPERTY(''SOH_ON'') <>''TRUE''
		RETURN
	
	SELECT @INSTALLATIONCODE =  CF_INSTALLATIONCODE FROM tCONFIGURATION
	IF NOT EXISTS (SELECT * FROM inserted)
	BEGIN	
		SELECT	@PID = NULL
		SELECT @SOH = 0
		SELECT @PEAN = ''''
		SELECT @DMLType = ''D''
	END 
	-- after update or insert statement
	ELSE
	BEGIN
		SELECT	@PID = ISNULL(P_ID,''''), 
				@SOH = ISNULL(P_QTYONHAND,0), 
				@SOO = ISNULL(P_QTYONORDER,0), 
				@LASTDELIVEREDPRICE = ISNULL(P_LastPriceDelivered,0), 
				@SP = ISNULL(P_SP,0), 
				@LASTDATEORDERED = ISNULL(P_LastDateOrdered,0), 
				@LASTQTYFIRMORDERED = ISNULL(P_LastQtyFirmOrdered,0), 
				@LASTQTYSSORDERED = ISNULL(P_LastQtySSOrdered,0), 

				@TOTALQTYSOLD = ISNULL(P_QtyTotalSold,0), 
				@LASTDELIVEREDDATE = ISNULL(P_LastDateDelivered,0), 
				@LASTDATESOLD = ISNULL(P_LastDateSold,0), 
				@LASTQTYDELIVERED = ISNULL(P_LastQtyDelivered,0), 

				@PEAN = ISNULL(P_EAN,'''') 
		FROM inserted
		-- after update statement
		IF EXISTS (SELECT * FROM deleted)
			SELECT 	@DMLType = ''U''
		ELSE
			SELECT	@DMLType = ''I''
	END
	SELECT @QtyOHBody = 
		''<SOHMsg>
			<DMLType>'' + @DMLType + ''</DMLType>
			<PID>'' + CAST(@PID AS VARCHAR(40)) + ''</PID>
			<SOH>'' + CAST(@SOH as VARCHAR(10)) + ''</SOH>
			<SOO>'' + CAST(@SOO as VARCHAR(10)) + ''</SOO>
			<LDP>'' + CAST(@LASTDELIVEREDPRICE as VARCHAR(10)) + ''</LDP>
			<SP>'' + CAST(@SP as VARCHAR(10)) + ''</SP>
			<LDO>'' + CONVERT(VARCHAR(20),@LASTDATEORDERED,120) + ''</LDO>
			<QTYFIRMORDERED>'' + CAST(@LASTQTYFIRMORDERED as VARCHAR(10)) + ''</QTYFIRMORDERED>
			<QTYSSORDERED>'' + CAST(@LASTQTYSSORDERED as VARCHAR(10)) + ''</QTYSSORDERED>

			<TOTALSOLD>'' + CAST(@TOTALQTYSOLD as VARCHAR(10)) + ''</TOTALSOLD>
			<LDREC>'' + CONVERT(VARCHAR(20),@LASTDELIVEREDDATE,120) + ''</LDREC>
			<LDSOLD>'' + CONVERT(VARCHAR(20),@LASTDATESOLD,120) + ''</LDSOLD>
			<QTYLASTDELIVERED>'' + CAST(@LASTQTYDELIVERED as VARCHAR(10)) + ''</QTYLASTDELIVERED>

			<EAN>'' + CAST(@PEAN as VARCHAR(20)) + ''</EAN>
			<STCODE>'' + @INSTALLATIONCODE + ''</STCODE>
		</SOHMsg>''
	--INSERT INTO _tSBLog (SBL_Msg,SBL_Proc) VALUES (ISNULL(CAST(@QtyOHBody as VARCHAR(580)),''NULL''),''trigQtyOHChange'')

	SELECT @CMD = ''SOHSOURCE_'' + @INSTALLATIONCODE + ''_SERVICE''

	IF NOT @QTYOHBODY IS NULL 
		EXEC dbo._usp_SendXML @CMD,''SOHCONSUMER_SERVICE'',''SOH_CONTRACT'', ''SOH_MSG'', @QtyOHBody,@RES,@ERRMESS
END TRY
BEGIN CATCH
		IF ERROR_PROCEDURE() <> OBJECT_NAME(@@PROCID)
			SELECT @ErrorString = ERROR_MESSAGE();
		ELSE
			SELECT @ErrorString = ''Error Message: '' + ERROR_MESSAGE() + '', Error Procedure: '' + 
						ERROR_PROCEDURE() + '', Error Line: '' + CAST(ERROR_LINE() as VARCHAR(10));
		
		RAISERROR (@ErrorString, 16,1)
END CATCH


SET NOCOUNT OFF; 


END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.trigQtyOHChange Trigger Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.trigQtyOHChange Trigger'
END
GO

--
-- Script To Update dbo.tProduct Table In (local).PBKS
-- Generated Monday, June 21, 2010, at 02:10 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.tProduct Table'
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   IF NOT EXISTS (SELECT name FROM sysobjects WHERE name = N'FK_tProduct_tPT')
      ALTER TABLE [dbo].[tProduct] WITH NOCHECK ADD CONSTRAINT [FK_tProduct_tPT] FOREIGN KEY ([P_ProductType_ID]) REFERENCES [dbo].[tPT] ([PT_ID]) NOT FOR REPLICATION
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.tProduct Table Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.tProduct Table'
END
GO
