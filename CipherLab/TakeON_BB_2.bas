VERSION ("711 Demo 1.02")


MainStart:
	TotalRecord& = TRANSACTION_COUNT_EX(1)					'get transaction file #2 total record
	LOCATE 8,14										'show record
	PRINT "0000"; TotalRecord&		
'**********  Initialization  ********** 
  ENABLE READER(1) 
  ON READER(1) GOSUB BcrData_1 
'************************************ 

	ShowMainMenuFlag% = 1
MainLoop:
	IF ShowMainMenuFlag% = 1 THEN
		SELECT_FONT(2)
		LOCATE 2,1
		PRINT "1.Collect  "
		LOCATE 4,1
		PRINT "2.Upload  "
		LOCATE 6,1
		PRINT "3.Utilities "

		ShowMainMenuFlag% = 0
	END IF

	KeyData$ = INKEY$									'select main menu
	IF KeyData$ <> "" THEN 
		IntKeyData% = ASC(KeyData$) 
		IF IntKeyData% = 13 THEN							'press enter-key to select
			IF SelectMainMenu$ = "1" THEN
				GOSUB CollectData						'collect data
			ELSE IF SelectMainMenu$ = "2" THEN
				GOSUB UploadData    					'upload data
			ELSE
			'	GOSUB Utilities						'utilities
			END IF	
			GOTO MainStart
		END IF
		ShowMainMenuFlag% = 1
	END IF
	GOTO MainLoop

'----------------------------------------------------------
'	Collect Data
'----------------------------------------------------------
CollectData:
	CLS
	CLR_KBD
	font% = 2										'GetInputData's parameter
	mark$ = " "	
	mode% = 1

	ESCToMainMenuFlag% = 1
	WHILE ESCToMainMenuFlag% > 0
	GetItem:										'Item		
		x% = 6
		y% = 2
		length% = 20
		alpha$ = "text"
		inputsource$ = "both"
		InputData$ = ""
'		GOSUB GetInputData

		IF InputData$ = "@KEY_ESC" THEN GOTO CDReturnToMainMenu		'press ESC key => return to main menu
		IF LEFT$(InputData$, 1) = "@" OR InputData$ = "" THEN GOTO GetItem	'press special key, data is space => goto GetItem

		Item$ = InputData$	
		Desc$ = ""
		Qty$ = ""	
		IF FIND_RECORD(1, 1, Item$) = 1 THEN					'data is found in lookup file #1
			StrData$ = GET_RECORD$ (1, 1)
			Desc$ = MID$(StrData$, 22,10)
			Qty$ = MID$(StrData$, 33, 5)	
		END IF	

		LOCATE 6,6									'show Desc data
		PRINT Desc$		
		CLS
		LOCATE 2,1
		PRINT "Item:"

	WEND
CDReturnToMainMenu:
	SELECT_FONT(1)
	RETURN
ReturnToGetInputData:
	CURSOR(0)										'cursor off
	DISABLE READER(1)									'disable reader
	RETURN
SCANDATA:
 	BCODE$ = GET_READER_DATA$(1) 
	SAVE_TRANSACTION(BCODE$)
	IF GET_FILE_ERROR <> 0 THEN PRINT "TR not saved"
	RETURN

'----------------------------------------------------------
'	Upload Data
'----------------------------------------------------------
UploadData:
	CLR_KBD
	ON ESC GOSUB ESCToMainMenu
	TotalRecord& = TRANSACTION_COUNT_EX(1)					'get transaction #2 total record
	IF TotalRecord& = 0 THEN								'if record = 0, it can't upload data
		GOSUB NoDataMsg
		RETURN						
	END IF

	ComStatus% = 1
	GOSUB SetComToOpen								'set COM port, and open COM port
	GOSUB ConnectingMsg								'show connecting message

	Record& = 0
	RecvData$ = ""									'clear string
	SendData$ = ""	
	RecvComData$ = ""
	ESCToMainMenuFlag% = 1	
	ON COM(Comport%) GOSUB ReadComData						'ON COM(N%) GOSUB SubLabel

	WHILE ESCToMainMenuFlag% > 0
		IF RecvComData$ <> "" THEN						'read data from COM port
			RecvData$ = RecvComData$						
			RecvComData$ = ""
			SendComData$ = ""

			IF RecvData$ = "CIPHER-UP" THEN					'receive "CIPHER-UP", reply with "ACK"
				SendComData$ = "ACK"
			ELSE IF RecvData$ = "CIPHER-DN" THEN				'download lookup file
				WRITE_COM(Comport%, "LOADNG")
				GOSUB UploadError
				GOTO UDReturnToMainMenu	
			ELSE IF RecvData$ = "RECORD" THEN				'send records
				SendComData$ = "RECORD:" + STR$(TotalRecord&)
			ELSE IF RecvData$ = "ACK" THEN					'receive "ACK", and save transaction file # 2
				Record& = Record& + 1
				IF Record& > TotalRecord& THEN 				'end of transaction file #2 => send "OVER"
					SendComData$ = "OVER"
				ELSE
					LOCATE 4,1						'show records 
					PRINT Record&				
					SendComData$ = "#" + GET_TRANSACTION_DATA_EX$(2, Record&) + "#"				
				END IF
			ELSE IF RecvData$ = "DONE" THEN					'receive "DONE" => show upload OK
				OFF COM(Comport%)
				SendComData$ = ""		  		
				LOCATE 2,1
				PRINT "Upload OK!"
				BEEP(3000,10)
				WAIT(10)
				ShowFlag% = 2
				GOTO UDReturnToMainMenu		
			ELSE IF  RecvData$ = "FAIL" THEN					'stop communication
				GOSUB UploadError
				GOTO UDReturnToMainMenu
			ELSE IF RecvData$ = "NAK" THEN					'add NAK command
				SendComData$ = SendData$
		  	ELSE IF RecvData$ <> "" THEN					
				SendComData$ = "NAK"
			END IF		
			
			IF SendComData$ <> "" THEN 					'write data to COM port
				WRITE_COM(Comport%, SendComData$)	
				SendData$ = SendComData$
			END IF		
		END IF
	WEND

UDReturnToMainMenu:
	SELECT_FONT(1)
	CLOSE_COM(Comport%)								'close COM

	IF ShowFlag% = 2 THEN								'upload ok
		SettingData$ = GET_TRANSACTION_DATA_EX$ (1, 1)			'get setting from transaction file #1
		StrData$ = MID$(SettingData$, 6, 1)						'get Erase setting   
		IF StrData$ = "1" THEN 							'manually
			GOTO DeleteData							'delete data menu
		ELSE										'automatically
			EMPTY_TRANSACTION_EX(2)					'empty transaction file #2
			GOSUB DeleteDataOk							'delete data ok
		END IF
	END IF
	RETURN

