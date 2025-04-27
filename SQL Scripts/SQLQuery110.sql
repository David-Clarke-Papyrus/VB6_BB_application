
UPDATE tPRODUCT SET P_COST = P_SP * 0.37
FROM         dbo.tProduct
WHERE     (P_SupplierID = 26) AND (P_Cost < 1000)
--UPDATE tPRODUCT SET P_COST = P_SP * 0.37
--FROM         dbo.tProduct
--WHERE     (P_SupplierID = 25) AND (P_Cost < 1000)
--UPDATE tPRODUCT SET P_COST = P_SP * 0.35
--FROM         dbo.tProduct
--WHERE     (P_SupplierID = 12) AND (P_Cost < 1000)

