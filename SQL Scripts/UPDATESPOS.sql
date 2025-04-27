--
-- Script To Update dbo.tExchange Table In 5.216.220.154\PBKSINSTANCE2.PBKSFD
-- Generated Thursday, March 4, 2010, at 11:32 AM
--
-- Please backup 5.216.220.154\PBKSINSTANCE2.PBKSFD before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.tExchange Table'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   ALTER TABLE [dbo].[tExchange]
      ALTER COLUMN [EXCH_Note] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.tExchange Table Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.tExchange Table'
END
GO

--
-- Script To Update dbo.tOpSession Table In 5.216.220.154\PBKSINSTANCE2.PBKSFD
-- Generated Thursday, March 4, 2010, at 11:32 AM
--
-- Please backup 5.216.220.154\PBKSINSTANCE2.PBKSFD before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.tOpSession Table'
GO

SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON
GO

SET NUMERIC_ROUNDABORT OFF
GO


IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE TABLE [dbo].[tmp_tOpSession] (
   [OPS_ID] [uniqueidentifier] NOT NULL,
   [OPS_Z_ID] [uniqueidentifier] NOT NULL,
   [OPS_StartTime] [datetime] NULL,
   [OPS_EndTime] [datetime] NULL,
   [OPS_OperatorID] [int] NULL,
   [OPS_Status] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
   [OPS_FloatValue] [numeric] (12, 2) NULL CONSTRAINT [DF_tOpSession_OPS_FloatValue] DEFAULT ((0)),
   [OPS_FloatBreakdown] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL
)
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   INSERT INTO [dbo].[tmp_tOpSession] ([OPS_ID], [OPS_Z_ID], [OPS_StartTime], [OPS_EndTime], [OPS_OperatorID], [OPS_Status], [OPS_FloatBreakdown])
   SELECT [OPS_ID], [OPS_Z_ID], [OPS_StartTime], [OPS_EndTime], [OPS_OperatorID], [OPS_Status], NULL
   FROM [dbo].[tOpSession]
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   DROP TABLE [dbo].[tOpSession]
GO

sp_rename N'[dbo].[tmp_tOpSession]', N'tOpSession'

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   ALTER TABLE [dbo].[tOpSession] ADD CONSTRAINT [PK_tOpSession] PRIMARY KEY CLUSTERED ([OPS_ID])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
   CREATE INDEX [iOPZID] ON [dbo].[tOpSession] ([OPS_Z_ID])
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.tOpSession Table Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.tOpSession Table'
END
GO

--
-- Script To Update dbo.UPDATE_DATA Procedure In 5.216.220.154\PBKSINSTANCE2.PBKSFD
-- Generated Thursday, March 4, 2010, at 11:32 AM
--
-- Please backup 5.216.220.154\PBKSINSTANCE2.PBKSFD before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.UPDATE_DATA Procedure'
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
ALTER PROCEDURE [dbo].[UPDATE_DATA]
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
TRUNCATE TABLE tMultibuys
SELECT [MB_SystemCode] FROM tMultibuys WHERE MB_SYSTEMCODE = ''mb1''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tMultiBuys]([MB_SystemCode], [MB_Description], [MB_Price], [MB_Qtygroup]) VALUES ( ''mb1'', ''R49.00 or 3 for R99'', 3300, 3)

SELECT [MB_SystemCode] FROM tMultibuys WHERE MB_SYSTEMCODE = ''mb2''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tMultiBuys]([MB_SystemCode], [MB_Description], [MB_Price], [MB_Qtygroup]) VALUES (''mb2'', ''R39.95 or 3 for R99'', 3300, 3)

SELECT [MB_SystemCode] FROM tMultibuys WHERE MB_SYSTEMCODE = ''mb3''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tMultiBuys]([MB_SystemCode], [MB_Description], [MB_Price], [MB_Qtygroup]) VALUES (''mb3'', ''R39.95 or 4 for R99'', 2475, 4)

SELECT [MB_SystemCode] FROM tMultibuys WHERE MB_SYSTEMCODE = ''mb4''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tMultiBuys]([MB_SystemCode], [MB_Description], [MB_Price], [MB_Qtygroup]) VALUES ( ''mb4'', ''R39.00 or 3 for R99'', 3300, 3)

SELECT [MB_SystemCode] FROM tMultibuys WHERE MB_SYSTEMCODE = ''mb5''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tMultiBuys]([MB_SystemCode], [MB_Description], [MB_Price], [MB_Qtygroup]) VALUES ( ''mb5'', ''R29.95 or 5 for R99 (romance)'', 1980, 5)

SELECT [MB_SystemCode] FROM tMultibuys WHERE MB_SYSTEMCODE = ''mb6''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tMultiBuys]([MB_SystemCode], [MB_Description], [MB_Price], [MB_Qtygroup]) VALUES ( ''mb6'', ''R29 or 5 for R99 (kids)'', 1980, 5)

SELECT [MB_SystemCode] FROM tMultibuys WHERE MB_SYSTEMCODE = ''mb7''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tMultiBuys]([MB_SystemCode], [MB_Description], [MB_Price], [MB_Qtygroup]) VALUES ( ''mb7'', ''R19.95 or 6 for R99'', 1650, 6)

SELECT [MB_SystemCode] FROM tMultibuys WHERE MB_SYSTEMCODE = ''mb8''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tMultiBuys]([MB_SystemCode], [MB_Description], [MB_Price], [MB_Qtygroup]) VALUES ( ''mb8'', ''R35.00 or 4 for R99'', 2475, 4)

SELECT [MB_SystemCode] FROM tMultibuys WHERE MB_SYSTEMCODE = ''mb9''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tMultiBuys]([MB_SystemCode], [MB_Description], [MB_Price], [MB_Qtygroup]) VALUES ( ''mb9'', ''R79.00 or 2 for R99'', 4950, 2)



SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''CaptureFloat''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription]) 
		VALUES (''CaptureFloat'', ''FALSE'', ''If TRUE then ask for float value at Log on'') 
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE then ask for float value at Log on'' 
		WHERE [PropertyKey] = ''CaptureFloat''

SELECT PropertyKey FROM tProperty WHERE PropertyKey = ''Secure''
IF @@ROWCOUNT = 0
	INSERT INTO [dbo].[tProperty]([PropertyKey], [PropertyValue], [PropertyDescription]) 
		VALUES (''Secure'', ''FALSE'', ''If TRUE then request signatures for special actions (e.g. refund'') 
ELSE
	UPDATE [dbo].[tProperty] SET [PropertyDescription] = ''If TRUE then request signatures for special actions (e.g. refund''
		WHERE [PropertyKey] = ''Secure''

END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.UPDATE_DATA Procedure Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.UPDATE_DATA Procedure'
END
GO