--
-- Script To Update dbo.spcu_GetTotalSalesOnAccount Procedure In (local).PBKS
-- Generated Thursday, April 8, 2010, at 10:00 PM
--
-- Please backup (local).PBKS before executing this script
--


BEGIN TRANSACTION
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

PRINT 'Updating dbo.spcu_GetTotalSalesOnAccount Procedure'
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
ALTER PROCEDURE dbo.spcu_GetTotalSalesOnAccount (@STARTDATE DATETIME,@ENDDATE DATETIME,@VAL REAL OUTPUT)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	SELECT  @VAL =   SUM(TotalRetail)
	FROM         vSalesSummary a Join tTR b ON a.TR_ID = b.TR_ID
	WHERE     a.SALETYPE IN (''INVOICE SALES'',''CREDIT RETURNS'') AND ISNULL(CAST(b.TR_EXCHANGEID as VARCHAR(40)),'''') =''''
		AND (a.dte BETWEEN dbo.startofDay(@STARTDATE) AND dbo.endofday(@ENDDATE)) 
END')
GO

IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION
GO

IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.spcu_GetTotalSalesOnAccount Procedure Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.spcu_GetTotalSalesOnAccount Procedure'
END
GO