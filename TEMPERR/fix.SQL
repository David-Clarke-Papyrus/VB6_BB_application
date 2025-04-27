--  
-- Script to Update dbo.vReorder_Browsed in (local).PBKS_Standard 
-- Generated Tuesday, November 15, 2011, at 02:47 PM 
--  
-- Please backup (local).PBKS_Standard before executing this script
--  
-- ** Script Begin **
BEGIN TRANSACTION

SET TRANSACTION ISOLATION LEVEL SERIALIZABLE


PRINT 'Updating dbo.vReorder_Browsed View'


SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON


SET NUMERIC_ROUNDABORT OFF



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
                      dbo.tProduct.P_QtyOnOrder_UnIssued, dbo.tProduct.P_QtyTotalSold
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
                      dbo.tProduct.P_QtyOnOrder_UnIssued, dbo.tProduct.P_QtyTotalSold
')


IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION


IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.vReorder_Browsed View Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.vReorder_Browsed View'
END



--  
-- Script to Update dbo.ReorderBrowsed in (local).PBKS_Standard 
-- Generated Tuesday, November 15, 2011, at 02:47 PM 
--  
-- Please backup (local).PBKS_Standard before executing this script
--  
-- ** Script Begin **
BEGIN TRANSACTION

SET TRANSACTION ISOLATION LEVEL SERIALIZABLE


PRINT 'Updating dbo.ReorderBrowsed Procedure'


SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, QUOTED_IDENTIFIER, CONCAT_NULL_YIELDS_NULL ON


SET NUMERIC_ROUNDABORT OFF



exec('ALTER PROCEDURE [dbo].[ReorderBrowsed] @Slatename VARCHAR(50),@WSNAME VARCHAR(50),@FILTER_OO BIT = 0,@FILTER_OH BIT = 0,@SUPPLIERID INT = 0 AS
DELETE FROM  tREORDERGENERAL WHERE SLATENAME = ''Browsed'' AND WSNAME = Host_Name()
INSERT  INTO dbo.tREORDERGENERAL  (SLATENAME,WSNAME,STATUS,COLID,REF,QTYFIRM,QTYSS,PID,QTYCO,QTYPO,QTYPOUNISSUED,QTYAPP,PRCODE,DESCRIP,LASTSUPPLIERID,LASTSUPPLIERNAME,LASTDEALID,LASTDEALNAME,PT,
PUBLISHER,TOTALSOLD,PRICE,ONHAND,LASTRECEIVEDDATE,LASTORDEREDDATE,LASTRECEIVEDQTY,LASTORDEREDQTYFIRM,
LASTORDEREDQTYSS,LASTRECEIVEDPRICE ,LASTORDEREDPRICE,LASTSIXWEEKS,LASTSIXMONTHS,TitleForSort,FOreignPrice,Catery) 
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

	WHERE c.CNT = 1	AND ISNULL(a.LastDealID,0) = 0
')


IF @@ERROR <> 0
   IF @@TRANCOUNT = 1 ROLLBACK TRANSACTION


IF @@TRANCOUNT = 1
BEGIN
   PRINT 'dbo.ReorderBrowsed Procedure Updated Successfully'
   COMMIT TRANSACTION
END ELSE
BEGIN
   PRINT 'Failed To Update dbo.ReorderBrowsed Procedure'
END

