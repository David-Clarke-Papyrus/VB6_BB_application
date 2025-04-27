VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum enCodeType
    enCustomerOrder = 19
    enAppro = 13
    enApproReturn = 14
    enCashSale = 20
    enSupplierOrder = 16
    enStockTake = 17
    enOOS = 25
    enInvoice = 21
    enProduct = 23
    enCReditnote = 26
    enDelivery = 18
    enTransfer = 22
    enReturn = 28
    en230 = 230
    en240 = 240
    en250 = 250
    en260 = 260
    en270 = 270
    en280 = 280
End Enum


Public Enum enStatus
    stCOMPLETE = 4
    stInProcess = 2
    stVOID = 1
    stPROFORMA = 7
    stISSUED = 3
    stCANCELLED = 6
End Enum

Public Enum enSortField
    enSequence = 1
    enTitle = 2
    enAuthor = 3
    enCode = 4
End Enum

Public Enum enumCommandType
    enTable = adCmdTable
    enQuery = adCmdStoredProc
    enText = adCmdText
End Enum

Public Enum enumWildcard
    enNone = 0
    enEnd = 1
    enStart = 2
End Enum
Public Enum enumActionTaken
    enFulfilled = 9
    enPOIssued = 1
    enNoAction = 0
End Enum

Public Enum enumCOType
    enNormalCO = 1
    enWant = 2
    enStanding = 3
End Enum
'Public Enum enumTRType
'    enOrder = 6
'    enAppro = 9
'    enInvoice = 3
'End Enum

Public Enum enCatalogue
    enGetsCatalogueYes = 1
    enGetsCatalogueNo = 2
    enGetsCatalogueEither = 3
End Enum
Public Enum enOverseas
    enOverseasYes = 1
    enOverseasNo = 2
    enOverseasEither = 3
End Enum
Public Enum enMailType
    enAirmail = 2
    enSurfaceMail = 3
    enEitherMail = 1
    enAll = 4
End Enum

Public Enum enumSince
    enAny = 1
    enWeek = 2
    enMonth = 3
    enQuarter = 4
    enYear = 5
End Enum
Public Enum enProductSTatus
    enInPrint = 0
    enOutOfPrint = 1
    enAwaitingReprint = 2
End Enum

Public Enum enUseStatus
    enActive = 0
    enInActive = 1
End Enum
Public Enum enProductDataSource
    enLocalDB = 0
    enBF = 1
End Enum

