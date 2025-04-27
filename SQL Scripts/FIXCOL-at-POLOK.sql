UPDATE tCOL SET COL_FULFILLED = 'FUL'
FROM tCOL  a JOIN
(SELECT     TOP (100) PERCENT dbo.tCOL.COL_ID, dbo.tCOL.COL_Fulfilled, dbo.COLFix3.[Column 3]
FROM         dbo.tTR INNER JOIN
                      dbo.tCO ON dbo.tTR.TR_ID = dbo.tCO.CO_ID INNER JOIN
                      dbo.tCOL ON dbo.tTR.TR_ID = dbo.tCOL.COL_TR_ID INNER JOIN
                      dbo.tProduct ON dbo.tCOL.COL_P_ID = dbo.tProduct.P_ID LEFT OUTER JOIN
                      dbo.COLFix3 ON dbo.tProduct.P_EAN = dbo.COLFix3.[Column 3] AND dbo.tCOL.COL_Ref = dbo.COLFix3.[Column 9]
WHERE     (dbo.tTR.TR_CaptureDate < CONVERT(DATETIME, '2010-03-21 00:00:00', 102)) AND (dbo.COLFix3.[Column 3] IS NULL)
) b ON a.COL_ID = b.COL_ID