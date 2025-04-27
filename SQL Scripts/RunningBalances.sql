
TRUNCATE table tPRODUCTHISTORY
INSERT INTO tPRODUCTHISTORY (PH_P_ID,PH_Date,PH_PreviousQTY,PH_PreviousCost) 
SELECT O.PID,O.Dte, 
    (SELECT sum(QTYSUM) from __RunningBalances where Dte < O.Dte AND PID = o.PID ),0
FROM __RunningBalances O WHERE Typ = 'DEL'
 ORDER BY PID,Dte

