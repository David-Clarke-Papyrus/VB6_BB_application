VERSION 5.00
Object = "{CCB90150-B81E-11D2-AB74-0040054C3719}#1.0#0"; "OposPOSPrinter.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Posiflex POS Printer OPOS Control Demo"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLDN 
      Height          =   360
      Left            =   3270
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   90
      Width           =   1755
   End
   Begin VB.CommandButton Command9 
      Caption         =   "OpEpson"
      Height          =   495
      Left            =   5955
      TabIndex        =   13
      Top             =   5055
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox List2 
      BackColor       =   &H8000000F&
      Height          =   4740
      ItemData        =   "Form1.frx":0000
      Left            =   240
      List            =   "Form1.frx":0002
      TabIndex        =   12
      Top             =   495
      Width           =   2775
   End
   Begin VB.Frame Frame5 
      Caption         =   "Device Statistics"
      Height          =   2025
      Left            =   120
      TabIndex        =   9
      Top             =   5400
      Width           =   6840
      Begin VB.CommandButton cmdRetrieveSt 
         Caption         =   "Retrieve Statistics"
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtRetrieveSt 
         Height          =   1245
         Left            =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PrintBitmap"
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   3120
      TabIndex        =   7
      Top             =   510
      Width           =   2535
   End
   Begin VB.CommandButton Command8 
      Caption         =   "CutPaper"
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "PrintImmediate"
      Height          =   495
      Left            =   5760
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SetBitmap"
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CheckHealth"
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin OposPOSPrinter_1_8_LibCtl.OPOSPOSPrinter Posiflex_PP1 
      Left            =   5760
      OleObjectBlob   =   "Form1.frx":0004
      Top             =   4200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idx1, idx2 As Integer
Dim ob As Object

Private Sub Command1_Click()
Set ob = Posiflex_PP1
Dim buf As String

 Cls
 idx1 = 0
 idx2 = 0
 List1.Clear
 AddList1 ("ST--> " + Str(ob.State))
 
 AddList2 ("OpenPf: " + Str(ob.Open(txtLDN)))
' AddList2 ("OpenPf: " + Str(ob.Open("PP Demo")))
'AddList2 ("OpenTm: " + Str(ob.Open("TM-T88II")))
 
 'Posiflex_PP1.FreezeEvents = True
 'Addlist2("FzEv--> "+str(ob.FreezeEvents
 'Addlist2("CapJrnPrs--> "; Posiflex_PP1.CapJrnPresent
 
 AddList1 ("ST--> " + Str(ob.State))
 AddList2 ("Claim: " + Str(ob.Claim(1000)))
 AddList1 ("ST--> " + Str(ob.State))
 ob.DeviceEnabled = True
 AddList2 ("DeviceEnabled: " + Str(ob.DeviceEnabled))
 AddList1 ("ST--> " + Str(ob.State))
 
 'ob.MapMode = PtrMmTwips
 AddList2 (ob.RecLineHeight)
 AddList2 (ob.RecLineSpacing)
 AddList2 (ob.RecLineWidth)
 AddList2 (ob.CharacterSetList)
 
 AddList2 (ob.CharacterSet)
 
 AddList2 (ob.CapMapCharacterSet)
 ob.MapCharacterSet = True
 AddList2 (ob.MapCharacterSet)
 AddList2 (ob.RecBitmapRotationList)
 
' ob.CharacterSet = 852
' AddList2 (ob.CharacterSet)
' ob.CharacterSet = 437
' AddList2 ("ob.Release(): " + Str(ob.Release()))
 
 cmdRetrieveSt.Enabled = Posiflex_PP1.CapStatisticsReporting
 DoEvents
 
 ' the demonstration device name "PP Demo"
 ' can be changed in the registery by using
 ' Posiflex OPOS Manager
End Sub

Private Sub Command2_Click()
Set ob = Posiflex_PP1

 AddList1 "ST--> " + Str(ob.State)
 ob.AsyncMode = True
 AddList2 ("ob.AsyncMode" + Str(ob.AsyncMode))
 AddList2 ("ob.CheckHealth(Ex)" + Str(ob.CheckHealth(OposChExternal)))
 DoEvents
 AddList2 ("ob.PrintBarCode(1)" + Str(ob.PrintBarCode(PtrSReceipt, "11111", PtrBcsCode39, 30, 3, PtrBcCenter, PtrBcTextAbove)))
 DoEvents
 AddList2 ("ob.PrintNormal(n)" + Str(ob.PrintNormal(PtrSReceipt, "  normal" + Chr(10))))
 DoEvents
 AddList2 ("ob.Transaction(begin)" + Str(ob.TransactionPrint(PtrSReceipt, PtrTpTransaction)))
 DoEvents
 AddList2 ("ob.RotatePrint(L90)" + Str(ob.RotatePrint(PtrSReceipt, PtrRpLeft90)))
 DoEvents
 AddList2 ("ob.PrintBarCode(2)" + Str(ob.PrintBarCode(PtrSReceipt, "222222", PtrBcsCode39, 30, 3, PtrBcLeft, PtrBcTextBelow)))
 DoEvents
 AddList2 ("ob.PrintNormal(L90)" + Str(ob.PrintNormal(PtrSReceipt, "  L90" + Chr(10))))
 DoEvents
 AddList2 ("ob.RotatePrint(180)" + Str(ob.RotatePrint(PtrSReceipt, PtrRpRotate180)))
 DoEvents
 AddList2 ("ob.PrintBarCode(3)" + Str(ob.PrintBarCode(PtrSReceipt, "3333333", PtrBcsCode39, 30, 3, PtrBcRight, PtrBcTextBelow)))
 DoEvents
 AddList2 ("ob.PrintNormal(u180)" + Str(ob.PrintNormal(PtrSReceipt, "  u180" + Chr(10))))
 DoEvents
 AddList2 ("ob.RotatePrint(0)" + Str(ob.RotatePrint(PtrSReceipt, PtrRpNormal)))
 DoEvents
 AddList2 ("ob.PrintBarCode(4)" + Str(ob.PrintBarCode(PtrSReceipt, "44444444", PtrBcsCode39, 30, 3, PtrBcLeft, PtrBcTextBelow)))
 DoEvents
 AddList2 ("ob.PrintNormal(n)" + Str(ob.PrintNormal(PtrSReceipt, "  normal" + Chr(10))))
 DoEvents
 AddList2 ("ob.RotatePrint(r90)" + Str(ob.RotatePrint(PtrSReceipt, PtrRpRight90)))
 DoEvents
 AddList2 ("ob.PrintBarCode(5)" + Str(ob.PrintBarCode(PtrSReceipt, "555555555", PtrBcsCode39, 30, 3, PtrBcLeft, PtrBcTextBelow)))
 DoEvents
 AddList2 ("ob.PrintNormal(r90)" + Str(ob.PrintNormal(PtrSReceipt, "  r90" + Chr(10))))
 DoEvents
 AddList2 ("ob.RotatePrint(L90)" + Str(ob.RotatePrint(PtrSReceipt, PtrRpLeft90)))
 DoEvents
 AddList2 ("ob.PrintBarCode(6)" + Str(ob.PrintBarCode(PtrSReceipt, "6666666666", PtrBcsCode39, 30, 3, PtrBcLeft, PtrBcTextBelow)))
 DoEvents
 AddList2 ("ob.PrintNormal(L90)" + Str(ob.PrintNormal(PtrSReceipt, "  L90" + Chr(10))))
 DoEvents
 AddList2 ("ob.RotatePrint(0)" + Str(ob.RotatePrint(PtrSReceipt, PtrRpNormal)))
 DoEvents
 AddList2 ("ob.Transaction(end)" + Str(ob.TransactionPrint(PtrSReceipt, PtrTpNormal)))
 DoEvents
 AddList2 ("ob.CheckHealth(In)" + Str(ob.CheckHealth(OposChInternal)))
 DoEvents
 ob.AsyncMode = False
 AddList1 "ST--> " + Str(ob.State)
End Sub

Private Sub Command3_Click()
Set ob = Posiflex_PP1
 
 Cls
 AddList2 ("ob.Close()" + Str(ob.Close()))
 DoEvents
 Beep
End Sub
'exit
Private Sub Command4_Click()
 Command3_Click
 End
End Sub

Private Sub Command5_Click()
Set ob = Posiflex_PP1
Dim ss As String

'---------------------------------------------
    AddList2 ("state: " + Str(ob.State))
    ob.AsyncMode = True
    ob.FlagWhenIdle = True
    AddList2 ("state: " + Str(ob.State))
    AddList2 ("flagWI: " + Str(ob.FlagWhenIdle))
    AddList2 ("SetBitmap(2,tt): " + Str(ob.SetBitmap(2, PtrSReceipt, "tt.bmp", 300, PtrBmLeft)))
    AddList2 ("state: " + Str(ob.State))
    AddList2 ("flagWI: " + Str(ob.FlagWhenIdle))
    AddList2 ("PrintNormal(): " + Str(ob.PrintNormal(PtrSReceipt, _
        "Posiflex POS Printer" + Chr(10) + Chr(27) + "|2B" + Chr(10))))
    AddList2 ("state: " + Str(ob.State))
    AddList2 ("flagWI: " + Str(ob.FlagWhenIdle))

'---------------------------------------------
'    ss = "aaaaaaaaaaa" + Chr(27) + "|300uF" + "bbbbbbb"
'    AddList2 ("ob.PrintNormal() " _
'        + Str(ob.PrintNormal(PtrSReceipt, ss)))

' ---------------------------------------------
'    ss = "FirmwareRevision=-33"
'    'AddList2 ("RetrieveStatistics: " + Str(ob.RetrieveStatistics(ss)))
'    AddList2 ("ResetStatistics: " + Str(ob.ResetStatistics(ss)))
'    AddList2 ("UpdateStatistics: " + Str(ob.UpdateStatistics(ss)))
'    AddList1 ss
'    DoEvents
    
'---------------------------------------------
'    ss = Chr(27) + "|21B 1xxxxxxxxxx" + Chr(10)
'    ss = Chr(27) + "|50P" + Chr(10)
'    ss = Chr(27) + "|2fC 1xxxxxxxxxx" + Chr(10) + Chr(27) + "|90P" + Chr(10)
'    ss = Chr(27) + "|2E 2xxxxxxxxxx" + Chr(10) + Chr(27) + "|90P" + Chr(10)
'    ss = Chr(27) + "|tbC 3xxxxxxxxxx" + Chr(10) + Chr(27) + "|90P" + Chr(10)
'    ss = Chr(27) + "|2uC 3xxxxxxxxxx" + Chr(10) + Chr(27) + "|90P" + Chr(10)
'    AddList2 ("ob.validateData " _
'        + Str(ob.ValidateData(PtrSReceipt, ss)))
'    AddList2 ("ob.PrintNormal() " _
'        + Str(ob.PrintNormal(PtrSReceipt, ss)))

'---------------------------------------------
'    AddList2 ("ob.PrintNormal() " _
'        + Str(ob.PrintNormal(PtrSReceipt, Chr(27) + "|2fT 2aaaaaaaax" + Chr(10))))
'    AddList2 ("ob.PrintNormal() " _
'        + Str(ob.PrintNormal(PtrSReceipt, Chr(27) + "|0fT 0aaaaaaaax" + Chr(10))))
'    AddList2 ("ob.PrintNormal() " _
'        + Str(ob.PrintNormal(PtrSReceipt, Chr(27) + "|2fT 2aaaaaaaax" + Chr(10))))
'    AddList2 ("ob.PrintNormal() " _
'        + Str(ob.PrintNormal(PtrSReceipt, Chr(27) + "|1fT 1aaaaaaaax" + Chr(10))))
'    AddList2 ("ob.PrintNormal() " _
'        + Str(ob.PrintNormal(PtrSReceipt, Chr(27) + "|2fT 2aaaaaaaax" + Chr(10))))
'    AddList2 ("ob.PrintNormal() " _
'        + Str(ob.PrintNormal(PtrSReceipt, Chr(27) + "|3fT 3aaaaaaaax" + Chr(10))))

'---------------------------------------------
'    AddList2 (Str(Posiflex_PP1.RecLineChars))
'    AddList2 (Str(Posiflex_PP1.RecLineWidth))
'    AddList2 (Str(Posiflex_PP1.RecLineHeight))
'    Posiflex_PP1.RecLineChars = 56
'    AddList2 (Str(Posiflex_PP1.RecLineChars))
'    AddList2 (Str(Posiflex_PP1.RecLineWidth))
'    AddList2 (Str(Posiflex_PP1.RecLineHeight))
'    AddList2 ("PrintNormal() " + Str(ob.PrintNormal(PtrSReceipt, "xaaaaaaaax" + Chr(10))))
'    AddList2 ("PrintNormal() " + Str(ob.PrintNormal(PtrSReceipt, "ybbbbbbbbby" + Chr(10))))

'---------------------------------------------
'    AddList2 (Str(Posiflex_PP1.RecLineSpacing))
'    Posiflex_PP1.RecLineSpacing = -50
'    AddList2 (Str(Posiflex_PP1.RecLineSpacing))
'    AddList2 ("PrintNormal() " + Str(ob.PrintNormal(PtrSReceipt, "xaaaaaaaax" + Chr(10))))
'    AddList2 ("PrintNormal() " + Str(ob.PrintNormal(PtrSReceipt, "ybbbbbbbbby" + Chr(10))))

'---------------------------------------------
' AddList2 ("ValidateData " + Str(ob.ValidateData(PtrSReceipt, Chr(27) + "|5hC" + Chr(27) + "|2vC" + "xaaaaaaaax" + Chr(10))))
 
'---------------------------------------------
'    AddList2 ("PrintBitmap()" + Str(ob.PrintBitmap(PtrSReceipt, "ss.bmp", PtrBmAsis, PtrBmLeft)))
'    AddList2 ("PrintBitmap()" + Str(ob.PrintBitmap(PtrSReceipt, "tt.bmp", 50, PtrBmCenter)))
'    AddList2 ("PrintBitmap()" + Str(ob.PrintBitmap(PtrSReceipt, "tt.bmp", 100, PtrBmCenter)))
 
'---------------------------------------------
' AddList2 ("TP: " + Str(Posiflex_PP1.TransactionPrint( _
' PtrSReceipt, PtrTpTransaction)))
' AddList2 ("PN: " + Str(ob.PrintNormal(PtrSReceipt, _
' "yaaaaaaaay" + Chr(27) + Chr(100) + Chr(2))))
' AddList2 ("PB: " + Str(Posiflex_PP1.PrintBarCode( _
' PtrSReceipt, "012345678901", PtrBcsUpce, 50, 50, _
' PtrBcCenter, PtrBcTextBelow)))
' AddList2 ("PN: " + Str(ob.PrintNormal(PtrSReceipt, _
' "xaaaaaaaax" + Chr(27) + Chr(100) + Chr(2))))
' 'AddList2 ("CLOP: " + Str(Posiflex_PP1.ClearOutput))
' AddList2 ("TPN: " + Str(Posiflex_PP1.TransactionPrint( _
' PtrSReceipt, PtrTpNormal)))
 
'---------------------------------------------
' ob.AsyncMode = True
' AddList2 ("ob.AsyncMode" + Str(ob.AsyncMode))
' AddList2 ("ob.RotatePrint(L90)" _
' + Str(ob.RotatePrint(PtrSReceipt, PtrRpLeft90)))
' AddList2 ("ob.PrintNormal(L90)" _
' + Str(ob.PrintNormal(PtrSReceipt, "xaaaaaaaax")))
' AddList2 ("ob.PrintNormal(L90)" + _
' Str(ob.PrintNormal(PtrSReceipt, "yaaaaaaaay")))
' AddList2 ("ob.RotatePrint(L90)" + _
' Str(ob.RotatePrint(PtrSReceipt, PtrRpNormal)))
' ob.AsyncMode = False

'---------------------------------------------
' ob.RotateSpecial = PtrRpRotate180
' ob.RotateSpecial = PtrRpRight90
' AddList2 ("ob.RotatePrint(180)" + Str(ob.RotatePrint(PtrSReceipt, PtrRpRotate180)))
' AddList2 ("ob.RotatePrint(L90)" + Str(ob.RotatePrint(PtrSReceipt, PtrRpLeft90)))
' DoEvents
 
 'AddList2 ("ob.PrintNormal(L90)" + Str(ob.PrintNormal(PtrSReceipt, "xaaaaaaaax" + Chr(27) + "|1B" + Chr(10))))
 'AddList2 ("ob.PrintNormal(L90)" + Str(ob.PrintNormal(PtrSReceipt, "yaaaaaaaay" + Chr(27) + "|3B" + Chr(10))))
 'Print Posiflex_PP1.PrintImmediate(PtrSReceipt, "abcdefghij" + Chr(10) + Chr(13))
 
' Print Posiflex_PP1.PrintBarCode(PtrSReceipt, "{C123456", PtrBcsCode128, 50, 50, PtrBcRight, PtrBcTextNone)
' Print Posiflex_PP1.PrintBarCode(PtrSReceipt, "{C1234567", PtrBcsCode128, 10, 10, PtrBcCenter, PtrBcTextBelow)
' Print Posiflex_PP1.PrintBarCode(PtrSReceipt, "{C12345678", PtrBcsCode128, 100, 100, PtrBcLeft, PtrBcTextAbove)
' Print Posiflex_PP1.PrintBarCode(PtrSReceipt, "{C123456789", PtrBcsCode128, 50, 50, PtrBcRight, PtrBcTextNone)
 
 'AddList2 ("ob.PrintNormal(L90)" + Str(ob.PrintNormal(PtrSReceipt, "abcdefghi" + Chr(10) + "xxxyyy")))
 'DoEvents
' AddList2 ("ob.PrintNormal(L90)" + Str(ob.PrintNormal(PtrSReceipt, " 123456789")))
' DoEvents
'
' AddList2 ("ob.printimmediate " + Str(Posiflex_PP1.PrintImmediate(PtrSReceipt, " 1qaz2wsx3edc")))
' DoEvents
'
' AddList2 ("ob.PrintNormal(L90)" + Str(ob.PrintNormal(PtrSReceipt, " -=[];',./" + Chr(10))))
' DoEvents
 
 'AddList2 ("ob.RotatePrint(L90)" + Str(ob.RotatePrint(PtrSReceipt, PtrRpNormal)))
 'DoEvents
 
 'Print Posiflex_PP1.PrintNormal(PtrSReceipt, "end of printing1")
 'Print Posiflex_PP1.PrintImmediate(PtrSReceipt, "end of printing2")
 
'---------------------------------------------
 'buf = """""""=2"
 'buf = "U_=3"
 'buf = "M_=4"
 'Print Posiflex_PP1.UpdateStatistics(buf)
 
'---------------------------------------------
' AddList2 ("state= " + Str(ob.State) + ", flag= " + Str(ob.FlagWhenIdle))
' DoEvents

' ob.AsyncMode = True
' AddList2 ("AsyncMode= " + Str(ob.AsyncMode))

' AddList2 ("ob.SetBitmap(tt)" + Str(ob.SetBitmap(2, PtrSReceipt, "tt.bmp", 300, PtrBmLeft)))
' DoEvents
' AddList2 ("ob.SetBitmap(ss)" + Str(ob.SetBitmap(1, PtrSReceipt, "ss.bmp", 300, PtrBmLeft)))
' DoEvents
'
' AddList2 ("ob.PrintNormal(|1B)" + Str(ob.PrintNormal(PtrSReceipt, "xaaaaaaaax" + Chr(10) + Chr(27) + "|1B" + Chr(10))))
' DoEvents
' AddList2 ("ob.PrintNormal(|2B)" + Str(ob.PrintNormal(PtrSReceipt, "yaaaaaaaay" + Chr(10) + Chr(27) + "|2B" + Chr(10))))
' DoEvents
'
' AddList2 ("state= " + Str(ob.State) + ", flag= " + Str(ob.FlagWhenIdle))
' DoEvents
'
' ob.AsyncMode = False
' AddList2 ("AsyncMode= " + Str(ob.AsyncMode))
'
' AddList2 ("state= " + Str(ob.State) + ", flag= " + Str(ob.FlagWhenIdle))
' DoEvents
' Posiflex_PP1.FlagWhenIdle = True
' DoEvents
' AddList2 ("state= " + Str(ob.State) + ", flag= " + Str(ob.FlagWhenIdle))
'' For i = 1 To 100000
'    DoEvents
'' Next
'' AddList2 ("ob.PrintNormal(|1B)" + Str(ob.PrintNormal(PtrSReceipt, "xaaaaaaaax" + Chr(10) + Chr(27) + "|1B" + Chr(10))))
'' DoEvents
' AddList2 ("state= " + Str(ob.State) + ", flag= " + Str(ob.FlagWhenIdle))
' DoEvents
 
End Sub

Private Sub Command6_Click()
Set ob = Posiflex_PP1
 
' ----------------------------------------
 AddList2 ("PrintBitmap= " + Str(ob.PrintBitmap(PtrSReceipt, "tt.bmp", PtrBmAsis, PtrBmLeft)))
 DoEvents
 
' ----------------------------------------
' ob.AsyncMode = True
' ob.RotateSpecial = PtrRpRotate180
' ob.RotateSpecial = PtrRpRight90
' ob.RotateSpecial = PtrRpLeft90

' AddList2 ("RotatePrint(180)" + Str(ob.RotatePrint(PtrSReceipt, PtrRpRotate180)))
' AddList2 ("RotatePrint(180)" + Str(ob.RotatePrint(PtrSReceipt, PtrRpRotate180 + &H3000)))
' AddList2 ("RotatePrint(R90)= " + Str(ob.RotatePrint(PtrSReceipt, PtrRpRight90)))
' AddList2 ("RotatePrint(L90)= " + Str(ob.RotatePrint(PtrSReceipt, PtrRpLeft90)))
' DoEvents
 
' AddList2 ("PrintNormal= " + Str(ob.PrintNormal(PtrSReceipt, "xaaaaaaaax" + Chr(27) + "|1B" + Chr(10))))
' AddList2 ("PrintNormal= " + Str(ob.PrintNormal(PtrSReceipt, "yaaaaaaaay" + Chr(27) + "|2B" + Chr(10))))
' Print Posiflex_PP1.PrintImmediate(PtrSReceipt, "abcdefghij" + Chr(10) + Chr(13))
 
' AddList2 ("printbarcode= " + Str(Posiflex_PP1.PrintBarCode(PtrSReceipt, "{C123456", PtrBcsCode128, 50, 50, PtrBcRight, PtrBcTextAbove)))
' AddList2 ("printbarcode= " + Str(Posiflex_PP1.PrintBarCode(PtrSReceipt, "{C1234567", PtrBcsCode128, 10, 10, PtrBcCenter, PtrBcTextBelow)))
 
' AddList2 ("PrintBitmap= " + Str(ob.PrintBitmap(PtrSReceipt, "tt.bmp", PtrBmAsis, PtrBmLeft)))
' AddList2 ("printbarcode= " + Str(Posiflex_PP1.PrintBarCode(PtrSReceipt, "{C123456", PtrBcsCode128, 50, 50, PtrBcRight, PtrBcTextAbove)))

' AddList2 ("PrintNormal= " + Str(ob.PrintNormal(PtrSReceipt, "xaaaaaaaax" + Chr(27) + "|1B" + Chr(10))))
' AddList2 ("PrintNormal= " + Str(ob.PrintNormal(PtrSReceipt, "BaaaaaaaaA")))
 
' AddList2 ("RotatePrint(N)= " + Str(ob.RotatePrint(PtrSReceipt, PtrRpNormal)))
' DoEvents

' ----------------------------------------
' Posiflex_PP1.AsyncMode = True
' Posiflex_PP1.TransactionPrint PtrSReceipt, PtrTpNormal
' Posiflex_PP1.TransactionPrint PtrSReceipt, PtrTpTransaction
'
' For i = 1 To 5
'    Posiflex_PP1.PrintNormal PtrSReceipt, "@@@@@@ " + Str(i) + " @@" + Chr(10)
'    DoEvents
' Next
' Posiflex_PP1.CutPaper (90)
'
' Posiflex_PP1.TransactionPrint PtrSReceipt, PtrTpNormal
' Posiflex_PP1.AsyncMode = False
 DoEvents
End Sub

Private Sub Command7_Click()
Dim ss As String
Set ob = Posiflex_PP1
 
     ss = Chr(27) + "|cA" + Chr(27) + "|2fT"
     ss = ss + "3. This is Font B."
     ss = ss + Chr(27) + "|N" + Chr(10)
     AddList2 ("ob.PrintNormal()" + Str(ob.PrintNormal(PtrSReceipt, ss)))
     DoEvents

     ss = Chr(27) + "|cA" + Chr(27) + "|2fT"
     ss = ss + Chr(27) + "|bC"
     ss = ss + "4. This is Bold Font B."
     ss = ss + Chr(27) + "|N" + Chr(10)
     AddList2 ("ob.PrintNormal()" + Str(ob.PrintNormal(PtrSReceipt, ss)))
     DoEvents

     ss = Chr(27) + "|cA" + Chr(27) + "|2fT"
     ss = ss + Chr(27) + "|2C"
     ss = ss + "5. This is Wide Font B."
     ss = ss + Chr(27) + "|N" + Chr(10)
     AddList2 ("ob.PrintNormal()" + Str(ob.PrintNormal(PtrSReceipt, ss)))
     DoEvents

     ss = Chr(27) + "|cA" + Chr(27) + "|bC"
     ss = ss + Chr(27) + "|2C" + Chr(27) + "|2fT"
     ss = ss + "6. Bold Wide Font B."
     ss = ss + Chr(27) + "|N" + Chr(10)
     AddList2 ("ob.PrintNormal()" + Str(ob.PrintNormal(PtrSReceipt, ss)))
     DoEvents

     ss = Chr(27) + "|cA"
     ss = ss + "7. This is Font A."
     ss = ss + Chr(27) + "|N" + Chr(10)
     AddList2 ("ob.PrintNormal()" + Str(ob.PrintNormal(PtrSReceipt, ss)))
     DoEvents

     ss = Chr(27) + "|cA" + Chr(27) + "|0fT" + Chr(27) + "|bC"
     ss = ss + "8. This is Bold Font A."
     ss = ss + Chr(27) + "|N" + Chr(10)
     AddList2 ("ob.PrintNormal()" + Str(ob.PrintNormal(PtrSReceipt, ss)))
     DoEvents

     ss = Chr(27) + "|cA" + Chr(27) + "|2C"
     ss = ss + "9. Wide Font A"
     ss = ss + Chr(27) + "|N" + Chr(10)
     AddList2 ("ob.PrintNormal()" + Str(ob.PrintNormal(PtrSReceipt, ss)))
     DoEvents

     ss = Chr(27) + "|cA" + Chr(27) + "|2C" + Chr(27) + "|bC"
     ss = ss + "0. Bold Wide Font A"
     ss = ss + Chr(27) + "|N" + Chr(10)
     AddList2 ("ob.PrintNormal()" + Str(ob.PrintNormal(PtrSReceipt, ss)))
     DoEvents

     ss = "" + Chr(28) + Chr(112) + Chr(1) + Chr(48) + Chr(10)
     AddList2 ("PrintImmediate()" + Str(ob.PrintImmediate(PtrSReceipt, ss)))
     DoEvents

     ss = "" + Chr(28) + Chr(112) + Chr(1) + Chr(48) + Chr(10)
     AddList2 ("PrintImmediate()" + Str(ob.PrintImmediate(PtrSReceipt, ss)))
     DoEvents
 
' ---------------------------------------------
' ss = Chr(27) + "|cA" + Chr(27) + "|2fT"
' ss = ss + "1. Font C." + Chr(10)
'' ss = ss + Chr(27) + "|N" + Chr(10)
' Addlist2("ob.PrintNormal()"+str(ob.PrintNormal(PtrSReceipt, ss)
' DoEvents
' Addlist2("ob.PrintNormal()"+str(ob.PrintNormal(PtrSReceipt, "test" + Chr(10))
'
' ss = Chr(27) + "|cA" + Chr(27) + "|3fT"
' ss = ss + "2. Font D." + Chr(10)
'' ss = ss + Chr(27) + "|N" + Chr(10)
' Addlist2("ob.PrintNormal()"+str(ob.PrintNormal(PtrSReceipt, ss)
' DoEvents
' Addlist2("ob.PrintNormal()"+str(ob.PrintNormal(PtrSReceipt, "test" + Chr(10))
 
End Sub

Private Sub Command8_Click()
Set ob = Posiflex_PP1

 ob.AsyncMode = False
 AddList2 ("ob.CutPaper() = " + Str(ob.CutPaper(90)))
 DoEvents
 AddList2 ("state= " + Str(ob.State) + ", flag= " + Str(ob.FlagWhenIdle))
 DoEvents
End Sub

Private Sub Command9_Click()    ' a test button
Dim i, j As Integer
Set ob = Posiflex_PP1
Dim buf As String

 Cls
 idx1 = 0
 idx2 = 0
 List1.Clear
 AddList1 ("ST--> " + Str(ob.State))
 
 AddList2 ("ob.Open()" + Str(ob.Open("TM-T88II")))
 
 'Posiflex_PP1.FreezeEvents = True
 'Addlist2(" FzEv--> "+str(ob.FreezeEvents
 'Addlist2(" CapJrnPrs--> "; Posiflex_PP1.CapJrnPresent
 
 AddList1 ("ST--> " + Str(ob.State))
 AddList2 ("ob.Claim()" + Str(ob.Claim(1000)))
 AddList1 ("ST--> " + Str(ob.State))
 ob.DeviceEnabled = True
 AddList2 ("ob.DeviceEnabled" + Str(ob.DeviceEnabled))
 AddList1 ("ST--> " + Str(ob.State))
 
 'ob.MapMode = PtrMmTwips
 AddList2 (ob.RecLineHeight)
 AddList2 (ob.RecLineSpacing)
 AddList2 (ob.RecLineWidth)
 AddList2 (ob.CharacterSetList)
 
 AddList2 (ob.CharacterSet)
 
 AddList2 (ob.CapMapCharacterSet)
 ob.MapCharacterSet = True
 AddList2 (ob.MapCharacterSet)
 
' ob.CharacterSet = 852
' AddList2 (ob.CharacterSet)
' ob.CharacterSet = 437
' AddList2 ("ob.Release(): " + Str(ob.Release()))
 
 cmdRetrieveSt.Enabled = Posiflex_PP1.CapStatisticsReporting
 DoEvents
 
 ' the demonstration device name "PP Demo"
 ' can be changed in the registery by using
 ' Posiflex OPOS Manager

' Open "LPT1:" For Output As #1
' DoEvents
' Print #1, Chr(29) + Chr(118) + Chr(48) + Chr(0);
' Print #1, Chr(80) + Chr(0) + Chr(127) + Chr(1);
' DoEvents
' For i = 1 To 256 + 127
' For j = 1 To 80
' Print #1, Chr(240);
' DoEvents
' Next
' Next
' Print #1, Chr(29) + Chr(86) + Chr(66) + Chr(200)
' Close #1
End Sub

Private Sub Form_Load()
   txtRetrieveSt.Text = "UnifiedPOSVersion,DeviceCategory,ManufacturerName,ModelName,SerialNumber,ManufacturerDate,MechanicalRevision,FirmwareRevision,Interface,InstallationDate,HoursPoweredCount,CommunicationErrorCount"
   txtRetrieveSt.Text = txtRetrieveSt.Text + _
                        ",BarcodePrintedCount,MaximumTempReachedCount,NVRAMWriteCount,PaperCutCount,FailedPaperCutCount,PrinterFaultCount,ReceiptCharacterPrintedCount,ReceiptCoverOpenCount,ReceiptLineFeedCount,ReceiptLinePrintedCount"
End Sub

Private Sub Posiflex_PP1_ErrorEvent(ByVal rc As Long, ByVal rce As Long, ByVal el As Long, er As Long)
 AddList1 ("Error-->" + Str(rc) + Str(rce) + Str(el) + Str(er))
 DoEvents
End Sub

Private Sub Posiflex_PP1_OutputComplete(ByVal id As Long)
 AddList1 ("OPC-->" + Str(id))
 DoEvents
End Sub

Private Sub Posiflex_PP1_OutputCompleteEvent(ByVal id As Long)
 AddList1 ("OPCEv-->" + Str(id))
' Addlist2("OpCEv --> ", id
 DoEvents
End Sub

Private Sub Posiflex_PP1_StatusUpdate(ByVal st As Long)
 AddList1 ("SU-->" + Str(st))
 DoEvents
End Sub

Private Sub Posiflex_PP1_StatusUpdateEvent(ByVal st As Long)
 AddList1 ("SUEv --> " + Str(st))
 AddList2 ("SUEv --> " + Str(st))
 DoEvents
End Sub

Private Sub AddList1(ByVal ss As String)
    DoEvents
    List1.List(idx1) = ss
    idx1 = idx1 + 1
    DoEvents
End Sub

Private Sub AddList2(ByVal ss As String)
    DoEvents
    List2.List(idx2) = ss
    idx2 = idx2 + 1
    DoEvents
End Sub

Private Sub cmdRetrieveSt_Click()

    Dim strParam As String
    Dim lLen As Long
    Dim strErrMsg As String
    Dim strXMLPath As String
    Dim strFindXMLPath As String

    DoEvents
    strParam = txtRetrieveSt.Text
    strErrMsg = ""
    strFindXMLPath = ""

    With Posiflex_PP1
        .RetrieveStatistics strParam
        If (.ResultCode <> OPOS_SUCCESS) Then
            strErrMsg = "RetrieveStatistics method error." + vbCrLf + vbCrLf
            strErrMsg = strErrMsg + "ResultCode = " + CStr(.ResultCode) + vbCrLf
            strErrMsg = strErrMsg + "ResultCodeExtended = " + CStr(.ResultCodeExtended)
            MsgBox strErrMsg, vbOKOnly + vbExclamation, "CashDrawer"
            Exit Sub
        End If
    End With
    
    strXMLPath = App.Path + "\demo.xml"
    'Delete XML file.
    strFindXMLPath = Dir(strXMLPath)
    If strFindXMLPath <> "" Then
        Kill (strXMLPath)
    End If
    'Create XML file.
    Open strXMLPath For Binary Access Write As #1
        Put #1, , strParam
    Close #1
    
    'Opens another window and indicates the information of the XML file.
    RetrieveStBrowser.Show
    DoEvents
End Sub

