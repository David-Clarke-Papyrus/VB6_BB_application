UPDATE tProductHistory Set PH_PreviousCost = OldCostPrice
FROM tPRODUCTHistory ph Join __GetPreviousCost_b pc ON ph.PH_P_ID = pc.PH_P_ID and ph.PH_Date = pc.PH_DATE

UPDATE tProductHistory Set PH_Cost = CostPrice,PH_Qty = DELL_QTYTOTAL
FROM tPRODUCTHistory ph Join __DeliveredCost dc ON ph.PH_P_ID = dc.DELL_P_ID and ph.PH_Date = dc.TR_CaptureDATE