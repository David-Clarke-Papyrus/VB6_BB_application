'********************************************************************************************************* 
' Program:  BASIC Sample Program -- 201TEST_1.BAS      * 
' Target Machine:  201              * 
' Description:  The terminal reads data from the reader ports and show the data and * 
'  the code type on the LCD.           * 
'********************************************************************************************************* 
  ‘VERSION ("201TEST_1.BAS") 
  ‘START_DEBUG(1,1,1,2,1) 
 
'**********    Constants    ********** 
  BUZ_VOL% = 2 
  BACK_LIT% = 3 
  PROMPT1$ = "Scan" 
  PROMPT2$ = "Download" 
 

'**********  Initialization  ********** 
  ENABLE READER(1) 
  ON READER(1) GOSUB BcrData_1 
'************************************ 
Main: 
  'VOL(BUZ_VOL%) 
  BACKLIT(BACK_LIT%) 
  BEEP(4750,15,0,15,4750,15,0,15,4750,15) 
  CLS 
  PRINT PROMPT1$ 
  LOCATE 6,1 
  PRINT PROMPT2$ 
  LOCATE 2,31 
  WAIT(400) 
  CURSOR(0) 
  LED(2,1,100) 
EOF$ = "FALSE"
SOUGHT$ = ""
ATTEMPT$ = ""

Loop: 
 	IF EOF$ = "TRUE" THEN
 		BACKLIT(BACK_LIT%) 
 		LED(2,1,800)
		LOCATE 6,1
		PRINT "**FOUND**"
 		BEEP(4750,5,2000,10,5000,75) 
 		LOCATE 8,1
		PRINT "1:new 2:cont"
		Locate 8,16
		INPUT NUM%
		IF NUM% = 1 THEN
  			ON READER(1) GOSUB BcrData_1 
			GOTO Main
		ELSE
			IF NUM% = 2 THEN
				LOCATE 4,1
				PRINT STRING$(16, " ") 
				LOCATE 5,1
				PRINT STRING$(16, " ")
 				LOCATE 6,1
				PRINT STRING$(16, " ")
 				LOCATE 7,1
				PRINT STRING$(16, " ") 				
				LOCATE 8,1
				PRINT STRING$(16, " ") 				
				EOF$ = "FALSE"
			END IF
		END IF
	END IF
	GOTO loop
 
'********************************************************************************************************* 
' Routine: BcrData_1              * 
' Purpose:To get sought code.       * 
'********************************************************************************************************* 
BcrData_1: 
  BEEP(4750,5,2000,10,5000,15) 
   BACKLIT(BACK_LIT%) 
   LED(1,1,100) 
  SOUGHT$ = GET_READER_DATA$(1) 
  ON READER(1) GOSUB BcrData_2 
 
  LOCATE 3,3
  PRINT SOUGHT$ 
  RETURN 

'********************************************************************************************************* 
' Routine: BcrData_2              * 
' Purpose:To get attempted matches to master code.       * 
'********************************************************************************************************* 
BcrData_2: 
  BEEP(2000,5,5000,10,2000,15) 
 
  ATTEMPT$ = GET_READER_DATA$(1) 
 		LOCATE 5,3
		PRINT ATTEMPT$

	If SOUGHT$ = ATTEMPT$ THEN
		EOF$ = "TRUE"
	END IF

  RETURN 
 
 
 
'********************************************************************************************************* 
' Routine: CheckType              * 
' Purpose:To check the type of the barcode.          * 
' Return:  BcrType$                * 
' Call:                  * 
'********************************************************************************************************* 
CheckType: 
  IF (CODE_TYPE = "A") THEN  
    BcrType$ = "Code 39" 
  ELSE IF (CODE_TYPE = "B") THEN  
    BcrType$ = "Italy Pharma-code" 
  ELSE IF (CODE_TYPE = "C") THEN  
    BcrType$ = "CIP 39" 
  ELSE IF (CODE_TYPE = "D") THEN  
    BcrType$ = "Industrial 25" 
  ELSE IF (CODE_TYPE = "E") THEN  
    BcrType$ = "Interleave 25" 
  ELSE IF (CODE_TYPE = "F") THEN  
    BcrType$ = "Matrix 25" 
  ELSE IF (CODE_TYPE = "G") THEN  
    BcrType$ = "Codabar (NW7)" 
  ELSE IF (CODE_TYPE = "H") THEN  
    BcrType$ = "Code 93" 
  ELSE IF (CODE_TYPE = "I ") THEN  
    BcrType$ = "Code 128" 
  ELSE IF (CODE_TYPE = "J") THEN  
    BcrType$ = "UPCE no Addon" 
  ELSE IF (CODE_TYPE = "K") THEN  
    BcrType$ = "UPCE w/ Addon 2" 
  ELSE IF (CODE_TYPE = "L") THEN  
    BcrType$ = "UPCE w/ Addon 5" 
  ELSE IF (CODE_TYPE = "M") THEN  
    BcrType$ = "EAN8 no Addon" 
  ELSE IF (CODE_TYPE = "N") THEN  
    BcrType$ = "EAN8 w/ Addon 2" 
  ELSE IF (CODE_TYPE = "O") THEN  
    BcrType$ = "EAN8 w/ Addon 5" 
  ELSE IF (CODE_TYPE = "P") THEN  
    BcrType$ = "UPCA/EAN13 no Addon" 
  ELSE IF (CODE_TYPE = "Q") THEN  
    BcrType$ = "EAN13 w/ Addon 2" 
  ELSE IF (CODE_TYPE = "R") THEN  
    BcrType$ = "EAN13 w/ Addon 5" 
  ELSE IF (CODE_TYPE = "S") THEN  
    BcrType$ = "MSI" 
  ELSE IF (CODE_TYPE = "T") THEN  
    BcrType$ = "Plessey" 
  ELSE IF (CODE_TYPE = "U") THEN  
    BcrType$ = "Code ABC" 
  ELSE IF (CODE_TYPE = "a") THEN  
    BcrType$ = "ISO Track 1" 
  ELSE IF (CODE_TYPE = "b") THEN  
    BcrType$ = "ISO Track 2" 
  ELSE IF (CODE_TYPE = "c") THEN  
    BcrType$ = "ISO Track 1 and 2" 
  ELSE IF (CODE_TYPE = "d") THEN  
    BcrType$ = "ISO Track 2 and 3" 
  ELSE  
    BcrType$ = "" 
  END IF 
  RETURN 
 
'##########  End of Program  ########## 