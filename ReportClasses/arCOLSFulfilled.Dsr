VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arCOLSFulfilled 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   16425
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   28972
   _ExtentY        =   13996
   SectionData     =   "arCOLSFulfilled.dsx":0000
End
Attribute VB_Name = "arCOLSFulfilled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long
Sub component(pRs As ADODB.Recordset)
    Set DC1.Recordset = pRs
    i = 1
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 12600
    Me.Height = 6000
    LogSaveToFile "Records of fulfilment: " & CStr(pRs.RecordCount)
End Sub



Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
   App.LogEvent "Error Event hit " & Number & " " & Description 'logs messages
   LogSaveToFile "Error Event hit " & Number & " " & Description 'logs messages
   CancelDisplay = True 'cancels message box display
    MsgBox "Error occurred in activeReports module. Please inform support. Error is logged in errors.txt", vbInformation + vbOKOnly, "Error"
End Sub

Private Sub ActiveReport_Initialize()
  ' Me.Printer.DeviceName = ""
End Sub
