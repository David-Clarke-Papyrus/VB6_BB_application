--Notes: inmport with ID autoincrement

--Step one -clean up leading/trailing spaces
UPDATE [__Inventory] SET [CODE] = RTRIM(LTRIM([CODE])),[PRICE] = RTRIM(LTRIM([PRICE])),[Title] = RTRIM(LTRIM([Title]))
UPDATE [__Inventory] SET [PRICE] = RTRIM(LTRIM([PRICE]))

--zeroise excessive(Error) prices
UPDATE [__Inventory] SET PRICE = '0'  WHERE CAST(Price as NUMERIC(18,2)) > 100000.00


--INSERT rows with valid ISBN10 values
EXEC SwitchTriggers NULL,'disable'
INSERT INTO tPRODUCT  (P_Code, P_Title, P_RRP, P_SP, P_Cost, P_QtyOnHand, P_QtyReserved, P_QtyCopiesOnHand, P_QtyExpectedBack, P_QtyOnAppro, P_QtyOnOrder, P_QtyOnBackorder, P_QtyTotalSold, P_SpecialVAT, P_ProductType_ID, P_ProductType, P_EAN, OLDPID )
SELECT _IMPORT_INV_ISBN10.Code, _IMPORT_INV_ISBN10.Title, ROUND(CAST(_IMPORT_INV_ISBN10.Price as NUMERIC(18,2))*1.14,2) *100, 0, 0, 0, 0, 0 , 0, 0, 0, 0, 0, 0, 1, 'B', '', _IMPORT_INV_ISBN10.ID
FROM _IMPORT_INV_ISBN10
EXEC SwitchTriggers NULL,'enable'
--INSERT rows with valid ISBN13 values


EXEC SwitchTriggers NULL,'disable'
INSERT INTO tPRODUCT  (P_EAN, P_Title, P_RRP, P_SP, P_Cost, P_QtyOnHand, P_QtyReserved, P_QtyCopiesOnHand, P_QtyExpectedBack, P_QtyOnAppro, P_QtyOnOrder, P_QtyOnBackorder, P_QtyTotalSold, P_SpecialVAT, P_ProductType_ID, P_ProductType, P_CODE, OLDPID )
SELECT LEFT(_IMPORT_INV_ISBN13.Code,13), _IMPORT_INV_ISBN13.Title, ROUND(CAST(_IMPORT_INV_ISBN13.Price as NUMERIC(18,2))*1.14,2) *100, 0, 0, 0, 0, 0 , 0, 0, 0, 0, 0, 0, 1, 'B', '', _IMPORT_INV_ISBN13.ID
FROM _IMPORT_INV_ISBN13  LEFT OUTER JOIN tPRODUCT ON _IMPORT_INV_ISBN13.ID = CAST(tPRODUCT.OLDPID as INT) WHERE tPRODUCT.OLDPID IS NULL ORDER BY CAST(_IMPORT_INV_ISBN13.Price as NUMERIC(18,2)) DESC
EXEC SwitchTriggers NULL,'enable'

--all records with alpha codes with '#' in front
EXEC SwitchTriggers NULL,'disable'
INSERT INTO tPRODUCT  (P_CODE, P_Title, P_RRP, P_SP, P_Cost, P_QtyOnHand, P_QtyReserved, P_QtyCopiesOnHand, P_QtyExpectedBack, P_QtyOnAppro, P_QtyOnOrder, P_QtyOnBackorder, P_QtyTotalSold, P_SpecialVAT, P_ProductType_ID, P_ProductType, P_EAN, OLDPID )
SELECT LEFT(_IMPORT_INV_IRREGULAR.Code,20), _IMPORT_INV_IRREGULAR.Title, ROUND(CAST(_IMPORT_INV_IRREGULAR.Price as NUMERIC(18,2))*1.14,2) *100, 0, 0, 0, 0, 0 , 0, 0, 0, 0, 0, 0, 1, 'B', '', _IMPORT_INV_IRREGULAR.ID
FROM _IMPORT_INV_IRREGULAR  LEFT OUTER JOIN tPRODUCT ON _IMPORT_INV_IRREGULAR.ID = CAST(tPRODUCT.OLDPID as INT) WHERE tPRODUCT.OLDPID IS NULL ORDER BY CAST(_IMPORT_INV_IRREGULAR.Price as NUMERIC(18,2)) DESC
EXEC SwitchTriggers NULL,'enable'

--here we dont put the hash in yet cause it will break the link to the POL and COL tables
EXEC SwitchTriggers NULL,'disable'
INSERT INTO tPRODUCT  (P_CODE, P_Title, P_RRP, P_SP, P_Cost, P_QtyOnHand, P_QtyReserved, P_QtyCopiesOnHand, P_QtyExpectedBack, P_QtyOnAppro, P_QtyOnOrder, P_QtyOnBackorder, P_QtyTotalSold, P_SpecialVAT, P_ProductType_ID, P_ProductType, P_EAN, OLDPID )
SELECT Left(_IMPORT_INV_IRREGULAR2.Code,19), _IMPORT_INV_IRREGULAR2.Title, ROUND(CAST(_IMPORT_INV_IRREGULAR2.Price as NUMERIC(18,2))*1.14,2) *100, 0, 0, 0, 0, 0 , 0, 0, 0, 0, 0, 0, 1, 'B', '', _IMPORT_INV_IRREGULAR2.ID
FROM _IMPORT_INV_IRREGULAR2  LEFT OUTER JOIN tPRODUCT ON _IMPORT_INV_IRREGULAR2.ID = CAST(tPRODUCT.OLDPID as INT) WHERE tPRODUCT.OLDPID IS NULL ORDER BY CAST(_IMPORT_INV_IRREGULAR2.Price as NUMERIC(18,2)) DESC
EXEC SwitchTriggers NULL,'enable'

UPDATE tPRODUCT SET p_SP = P_RRP WHERE ISNULL(p_SP,0) = 0