VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmPrintLabels 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Print labels"
   ClientHeight    =   3975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSortByTitle 
      Caption         =   "Sort by title "
      Height          =   255
      Left            =   9840
      TabIndex        =   17
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton btnSortInReceivedOrder 
      Caption         =   "Sort in order received"
      Height          =   255
      Left            =   7920
      TabIndex        =   16
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrintReport 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6870
      Picture         =   "frmPrintLabels.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3270
      Width           =   1230
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5100
      Top             =   3540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrinter 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Printer "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3705
      Picture         =   "frmPrintLabels.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3270
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10710
      Picture         =   "frmPrintLabels.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3255
      Width           =   1000
   End
   Begin VB.CommandButton cmdADDProd 
      BackColor       =   &H00CCC8BB&
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2055
      Picture         =   "frmPrintLabels.frx":0A9E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3420
      Width           =   510
   End
   Begin VB.CommandButton cmdAddTFR 
      BackColor       =   &H00CCC8BB&
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3795
      Picture         =   "frmPrintLabels.frx":0E28
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2790
      Width           =   510
   End
   Begin VB.ListBox lstTransfers 
      BackColor       =   &H00DDF9FB&
      ForeColor       =   &H8000000D&
      Height          =   1230
      Left            =   105
      TabIndex        =   8
      Top             =   1905
      Width           =   3660
   End
   Begin VB.CommandButton cmdAddDels 
      BackColor       =   &H00CCC8BB&
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3795
      Picture         =   "frmPrintLabels.frx":11B2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   510
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   105
      TabIndex        =   3
      Top             =   3480
      Width           =   1920
   End
   Begin VB.ListBox lstDels 
      BackColor       =   &H00DDF9FB&
      ForeColor       =   &H8000000D&
      Height          =   1230
      Left            =   120
      TabIndex        =   2
      Top             =   315
      Width           =   3660
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print labels"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9390
      Picture         =   "frmPrintLabels.frx":153C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   1290
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   2730
      Left            =   4575
      OleObjectBlob   =   "frmPrintLabels.frx":18C6
      TabIndex        =   0
      Top             =   420
      Width           =   7125
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   5475
      TabIndex        =   13
      Top             =   3225
      Width           =   2220
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Transfers"
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   180
      TabIndex        =   9
      Top             =   1695
      Width           =   1410
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Products for label printing"
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   4635
      TabIndex        =   6
      Top             =   225
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Deliveries"
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   180
      TabIndex        =   5
      Top             =   90
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "ISBN-13"
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   165
      TabIndex        =   4
      Top             =   3255
      Width           =   1860
   End
End
Attribute VB_Name = "frmPrintLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oDel As a_Delivery
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim lngArrayRows As Long
Dim oSM As New z_StockManager
Dim tlDeliveries As New z_TextList
Dim tlTransfers As New z_TextList
Dim Row0 As String * 132
Dim Row1 As String * 132
Dim Row2 As String * 132
Dim Row3 As String * 132
Dim Row4 As String * 132
Dim Row5 As String * 132
Dim BarcodeOn As String
Dim BarcodePrefix As String
Dim iFileNo As Integer
Dim x As Printer
Dim strDevicename As String
Dim arPrinter() As String
Dim strCodeSet As String
Dim Oldprinter As Object

Public Sub component(pType As String, Optional oObj As Object, Optional x As XArrayDB)
    On Error GoTo errHandler
Dim i As Integer
    If pType = "D" Then
        Set XA = oSM.GetLabelsToPrint("D", True, oObj.TRID)
    ElseIf pType = "T" Then
        Set XA = oSM.GetLabelsToPrint("T", True, oObj.TRID)
    ElseIf pType = "S" Then
        For i = 1 To x.UpperBound(1)
            strCodeSet = strCodeSet & IIf(strCodeSet > "", ",", "") & x(i, 6)
        Next
        Set XA = oSM.GetLabelsToPrint("S", True, , , strCodeSet)
    End If
    LoadGrid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.component(pType,oObj,x)", Array(pType, oObj, x)
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim lngIndex As Long
Dim oDELL As a_DeliveryLine

    XA.QuickSort 1, XA.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
    Set Grid1.Array = XA
    Grid1.ReBind
    
    tlDeliveries.Load ltDeliveries, Format(DateAdd("ww", -2, Date), "yyyy-mm-dd")
    tlTransfers.Load ltTransfers, Format(DateAdd("ww", -2, Date), "yyyy-mm-dd")
    
    LoadListbox lstDels, tlDeliveries
    LoadListbox lstTransfers, tlTransfers
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.LoadGrid"
End Sub


Private Sub cmdAdd_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.cmdAdd_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub btnSortByTitle_Click()
    XA.QuickSort 1, XA.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
    Set Grid1.Array = XA
    Grid1.ReBind
End Sub

Private Sub btnSortInReceivedOrder_Click()
    XA.QuickSort 1, XA.UpperBound(1), 22, XORDER_ASCEND, XTYPE_LONG
    Set Grid1.Array = XA
    Grid1.ReBind
    
End Sub

Private Sub cmdAddDels_Click()
    On Error GoTo errHandler
    Set XA = oSM.GetLabelsToPrint("D", False, tlDeliveries.Key(lstDels.text))
    XA.QuickSort 1, XA.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
    Set Grid1.Array = XA
    Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.cmdAddDels_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdADDProd_Click()
    On Error GoTo errHandler
    Set XA = oSM.GetLabelsToPrint("P", False, , txtCode)
    XA.QuickSort 1, XA.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
    Set Grid1.Array = XA
    Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.cmdADDProd_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAddTFR_Click()
    On Error GoTo errHandler
    Set XA = oSM.GetLabelsToPrint("T", False, tlTransfers.Key(lstTransfers.text))
    XA.QuickSort 1, XA.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
    Set Grid1.Array = XA
    Grid1.ReBind
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.cmdAddTFR_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.cmdAddTFR_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.cmdClose_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdPrint_Click()
    On Error GoTo errHandler

Dim Label() As String
Dim i As Long
Dim j As Long
Dim k As Integer
Dim tmp As String
        
    arPrinter = Split(oPC.LabelPrinter, ":")
    If UBound(arPrinter, 1) > 0 Then
        strDevicename = arPrinter(1)
    Else
        strDevicename = ""
    End If
    
    If UCase(arPrinter(0)) = "ZEBRA" Then
        PrintToZebra
        Exit Sub
    End If
    If UCase(arPrinter(0)) = "ZEBRA600" Then
        PrintToZebra600
        Exit Sub
    End If
    If UCase(arPrinter(0)) = "ZEBRATLP" Then
        PrintLabelsToZebraTLP
        Exit Sub
    End If
    BarcodeOn = Chr(27) & Chr(16) & Chr(65) & Chr(4) & Chr(0) & Chr(2) & Chr(0) & Chr(2)
    BarcodePrefix = BarcodeOn & Chr(27) & Chr(16) & Chr(66) & Chr(13)
    
    If MsgBox("Have you lined up the printer and checked the label stationery?", vbYesNo + vbQuestion, "Confirm") = vbNo Then Exit Sub
    
    iFileNo = FreeFile
On Error Resume Next

    Open arPrinter(1) For Output As #iFileNo
    If Err Then
        MsgBox "No printer is set for labels, The labels printer must be called LABELS in Windows and the PBKS.INI file must be correctly completed."
        Exit Sub
    End If
On Error GoTo errHandler
    Print #iFileNo, Chr(27) & Chr(32) & Chr(0)
    k = 0
    Row0 = ""
    Row1 = ""
    Row2 = ""
    Row3 = ""
    Row4 = ""
    Row5 = ""
    
    XA.QuickSort 1, XA.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING

''This prints the non-barcode labels
    For i = 1 To XA.UpperBound(1)
        If XA(i, 7) = False Then
            'loop to iterate number of copies to print
            For j = 1 To FNN(XA(i, 5)) + FNN(XA(i, 4))
                AddtoRow1 XA(i, 1), k Mod 5
                AddtoRow2 Space(19 - Len(XA(i, 3))) & XA(i, 3), k Mod 5
                AddtoRow3 Left(XA(i, 2), 19), k Mod 5
                AddtoRow4 Mid(XA(i, 2), 20, 19), k Mod 5
                tmp = Left(oPC.Configuration.DefaultStore.Description, 13)
                tmp = tmp & Space(14 - Len(tmp))
                AddtoRow5 tmp & Format(Date, "MM/YY"), k Mod 5
                k = k + 1
                If k Mod 5 = 0 Then
                    PrintRow_NonBarcode
                End If
            Next j
        End If
    Next
    If Trim(Row1) > "" Then
        PrintRow_NonBarcode
    End If
''And this prints the barcode labels
    k = 0
    Row0 = ""
    Row1 = ""
    Row2 = ""
    Row3 = ""
    Row4 = ""
    Row5 = ""
    For i = 1 To XA.UpperBound(1)
        If XA(i, 7) = True Then
            'loop to iterate number of copies to print
            For j = 1 To FNN(XA(i, 5)) + FNN(XA(i, 4))
                AddtoRow0 XA(i, 8), k Mod 5
                AddtoRow1 XA(i, 1), k Mod 5
                AddtoRow2 Space(19 - Len(XA(i, 3))) & XA(i, 3), k Mod 5
                AddtoRow3 Left(XA(i, 2), 19), k Mod 5
                tmp = Left(oPC.Configuration.DefaultStore.Description, 13)
                tmp = tmp & Space(14 - Len(tmp))
                AddtoRow5 tmp & Format(Date, "MM/YY"), k Mod 5
                k = k + 1
                If k Mod 5 = 0 Then
                    PrintRow_Barcode
                End If
            Next j
        End If
    Next
    
    If Row1 > "" Then
        PrintRow_Barcode
    End If
    
Close #iFileNo
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.cmdPrint_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.cmdPrint_Click", , EA_NORERAISE, , "error line", Array(Erl())
    HandleError
End Sub
Private Sub PrintToZebra()
    On Error GoTo errHandler
Dim i, j, k As Integer
Dim tmp
Dim ar As arZebraLabels
Dim arB As arZebraLabelswithBarcode
Dim Dname As String
Const iMaxLineChars = 50
Const iMax2 = 40
Dim strDensity As String

    
    strDensity = Trim(oPC.GetProperty("LabelDensity"))
    If IsNumeric(strDensity) Then
        strDensity = "D" & strDensity
    Else
        strDensity = "D7"    'average
    End If
    For i = 1 To XA.UpperBound(1)
        If XA(i, 5) = False And FNN(XA(i, 4)) > 0 Then
            iFileNo = FreeFile
            Open arPrinter(1) For Output As #iFileNo
            Print #iFileNo, ""
            Print #iFileNo, "N"
            Print #iFileNo, "q400"
            Print #iFileNo, strDensity
            Print #iFileNo, "A10,5,0,4,1,1,N,""Code: " & IIf(Len(FNS(XA(i, 1))) < 16, FNS(XA(i, 8)), FNS(XA(i, 1))) & """"
            tmp = Trim(XA(i, 3))
            tmp = Space(23 - Len(tmp)) & tmp
            Print #iFileNo, "A10,40,0,4,1,1,N,""" & tmp & """"
            Row3 = ""
            Row4 = ""
            AddtoRow3 Left(XA(i, 2), 23), 0
            AddtoRow4 Mid(XA(i, 2), 24, 23), 0
            Print #iFileNo, "A10,75,0,4,1,1,N,""" & Row3 & """"
            Print #iFileNo, "A10,110,0,4,1,1,N,""" & Row4 & """"
            tmp = oPC.Configuration.DefaultStore.Description
            tmp = Left(tmp, 20) & Space(7) & Format(Date, "MM/YY")
'            tmp = oPC.Configuration.DefaultStore.Description
'            tmp = Left(tmp, 20) & Space(7) & Format(Date, "MM/YY")

            Print #iFileNo, "A10,145,0,4,1,1,N,""" & tmp & """"
            Print #iFileNo, "P" & CStr(FNN(XA(i, 4)))
            Close #iFileNo
        End If
    Next
    For i = 1 To XA.UpperBound(1)
        If XA(i, 5) = True Then
            iFileNo = FreeFile
            Open "LPT2" For Output As #iFileNo
            Print #iFileNo, ""
            Print #iFileNo, "N"
            Print #iFileNo, "q400"
            Print #iFileNo, strDensity
            
            Print #iFileNo, "B5,0,0,E30,3,10,60,B,""" & XA(i, 6) & """"
            
            tmp = Trim(XA(i, 3))
            Print #iFileNo, "A270,85,0,4,1,1,N,""" & tmp & """"
            Row3 = ""
            AddtoRow3 Left(XA(i, 2), 23), 0
            Print #iFileNo, "A5,120,0,4,1,1,N,""" & Row3 & """"
            tmp = oPC.Configuration.DefaultStore.Description
            tmp = Left(tmp, 20) & Space(7) & Format(Date, "MM/YY")
            tmp = oPC.Configuration.DefaultStore.Description
            tmp = Left(tmp, 20) & Space(7) & Format(Date, "MM/YY")

            Print #iFileNo, "A10,155,0,4,1,1,N,""" & tmp & """"
            Print #iFileNo, "P" & CStr(FNN(XA(i, 4)))
            Close #iFileNo
        End If
    Next
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.PrintToZebra"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.PrintToZebra"
End Sub
Private Sub AddtoRow0(pText As String, pPos As Integer)
    On Error GoTo errHandler
    Row0 = ReplaceEx(Row0, (13 * (pPos)), Len(pText), pText)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRow0(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRow0(pText,pPos)", Array(pText, pPos)
End Sub
Private Sub AddtoRow1(pText As String, pPos As Integer)
    On Error GoTo errHandler
    Row1 = ReplaceEx(Row1, (25 * (pPos)), Len(pText), pText)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRow1(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRow1(pText,pPos)", Array(pText, pPos)
End Sub
Private Sub AddtoRow2(pText As String, pPos As Integer)
    On Error GoTo errHandler
    Row2 = ReplaceEx(Row2, (25 * (pPos)), Len(pText), pText)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRow2(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRow2(pText,pPos)", Array(pText, pPos)
End Sub
Private Sub AddtoRow3(pText As String, pPos As Integer)
    On Error GoTo errHandler
    Row3 = ReplaceEx(Row3, (25 * (pPos)), Len(pText), pText)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRow3(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRow3(pText,pPos)", Array(pText, pPos)
End Sub
Private Sub AddtoRow4(pText As String, pPos As Integer)
    On Error GoTo errHandler
    Row4 = ReplaceEx(Row4, (25 * (pPos)), Len(pText), pText)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRow4(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRow4(pText,pPos)", Array(pText, pPos)
End Sub
Private Sub AddtoRow5(pText As String, pPos As Integer)
    On Error GoTo errHandler
    Row5 = ReplaceEx(Row5, (25 * (pPos)), Len(pText), pText)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRow5(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRow5(pText,pPos)", Array(pText, pPos)
End Sub

Private Sub PrintRow_NonBarcode()
    On Error GoTo errHandler
    Print #iFileNo, Row1
    Row1 = ""
    Print #iFileNo, Row2
    Row2 = ""
    Print #iFileNo, Row3
    Row3 = ""
    Print #iFileNo, Row4
    Row4 = ""
    Print #iFileNo, Row5
    Row5 = ""
    Print #iFileNo, ""
    Print #iFileNo, ""
    Print #iFileNo, ""
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.PrintRow_NonBarcode"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.PrintRow_NonBarcode", , EA_NORERAISE
    HandleError
End Sub
Private Sub PrintRow_Barcode()
    On Error GoTo errHandler
Dim strBarcode As String
Dim tmp As String

    strBarcode = BarcodePrefix & Mid(Row0, 1, 13) & "        "
    tmp = Mid(Row0, 14, 13)
    If tmp > "" Then
        strBarcode = strBarcode & BarcodePrefix & tmp & "        "
    End If
    tmp = Mid(Row0, 27, 13)
    If tmp > "" Then
        strBarcode = strBarcode & BarcodePrefix & tmp & "        "
    End If
    tmp = Mid(Row0, 40, 13)
    If tmp > "" Then
        strBarcode = strBarcode & BarcodePrefix & tmp & "        "
    End If
    tmp = Mid(Row0, 53, 13)
    If tmp > "" Then
        strBarcode = strBarcode & BarcodePrefix & tmp & "        "
    End If
    Print #iFileNo, strBarcode
    Print #iFileNo, ""
    Row0 = ""
    Print #iFileNo, Row1
    Row1 = ""
    Print #iFileNo, Row2
    Row2 = ""
    Print #iFileNo, Row3
    Row3 = ""
    Print #iFileNo, Row5
    Row5 = ""
    Print #iFileNo, ""
    Print #iFileNo, ""
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.PrintRow_Barcode"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.PrintRow_Barcode", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdPrintReport_Click()
Dim col0 As Long
Dim col1 As Long
Dim col2 As Long
Dim col3 As Long

    Grid1.PrintInfo.PageHeader = "Label print report"
    Grid1.PrintInfo.PageFooter = "\tPage:  \p of page \P"
    Grid1.PrintInfo.PreviewCaption = "Label print report"
    Grid1.PrintInfo.SettingsOrientation = 1
    col0 = Grid1.Columns(0).Width
    Grid1.Columns(0).Width = 2000
    col1 = Grid1.Columns(1).Width
    Grid1.Columns(1).Width = 4000
    col2 = Grid1.Columns(2).Width
    Grid1.Columns(2).Width = 1000
    col3 = Grid1.Columns(3).Width
    Grid1.Columns(3).Width = 1000
    
    Grid1.PrintInfo.PrintPreview 0
    
    Grid1.Columns(0).Width = col0
    Grid1.Columns(1).Width = col1
    Grid1.Columns(2).Width = col2
    Grid1.Columns(3).Width = col3
End Sub



Private Sub Form_Load()
    If oPC.LabelPrinter <> "OKI" Then
        Grid1.Columns(6).Visible = False
    End If
End Sub

Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
    If ColIndex = 4 Then
        If IsNumeric(Grid1.text) Then
            XA.Value(Grid1.Bookmark, ColIndex + 1) = CDbl(Grid1.text)
        Else
            Cancel = True
        End If
    End If
    Exit Sub
Errh:
    MsgBox Error
    Exit Sub
    Resume
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
'         OldValue, Cancel)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
         OldValue, Cancel), EA_NORERAISE
    HandleError
End Sub


Private Sub Grid1_LostFocus()
    On Error GoTo errHandler
    Grid1.Update
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.Grid1_LostFocus"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.Grid1_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub PrintToZebra600()
    On Error GoTo errHandler
Dim i, j, k As Integer
Dim tmp
Dim ar As arZebraLabels
Dim arB As arZebraLabelswithBarcode
Dim Dname As String
Const iMaxLineChars = 50
Const iMax2 = 40
Dim strPos As String
Dim strFileno As String
                    iFileNo = FreeFile
                    Open arPrinter(1) For Output As #iFileNo

''This prints the non-barcode labels
    k = 0
    For i = 1 To XA.UpperBound(1)
        If XA(i, 5) = False Then
            'loop to iterate number of copies to print
            For j = 1 To XA(i, 4)
                If k Mod 2 = 0 Then
                    strFileno = strFileno & CStr(iFileNo) & ","
                    Print #iFileNo, "^MTT"
                    Print #iFileNo, "^MFF,F"
                    Print #iFileNo, "^XA"
                    Print #iFileNo, "^PW900"
                    Print #iFileNo, "^LH30,30"
                End If
                
                AddtoRowZ1n IIf(Len(FNS(XA(i, 1))) < 16, FNS(XA(i, 8)), FNS(XA(i, 1))), k Mod 2
                AddtoRowZ2n FNS(XA(i, 3)), k Mod 2
                AddtoRowZ3n Left(XA(i, 2), 32), k Mod 2
                AddtoRowZ4n Mid(XA(i, 2), 32, 32), k Mod 2
                'tmp =
                AddtoRowZ5n FNS(XA(i, 8)), k Mod 2
                tmp = Left(oPC.Configuration.DefaultStore.Description, 32)
                AddtoRowZ5n FNS(XA(i, 10)), k Mod 2
                tmp = Left(oPC.Configuration.DefaultStore.Description, 32)
                tmp = tmp & Space(32 - Len(tmp))
                AddtoRowZ6n tmp & "  " & Format(Date, "MM/YY"), k Mod 2
                k = k + 1
                If k Mod 2 = 0 Then
                    Print #iFileNo, "^XZ"
                End If
            Next j
        End If
    Next
    If k Mod 2 = 1 Then
        Print #iFileNo, "^XZ"
    End If
''And this prints the barcode labels
    k = 0
    For i = 1 To XA.UpperBound(1)
        If XA(i, 5) = True Then
            'loop to iterate number of copies to print
            For j = 1 To XA(i, 4)
                If k Mod 2 = 0 Then
                    Print #iFileNo, "^MTT"
                    Print #iFileNo, "^MFF,F"
                    Print #iFileNo, "^XA"
                    Print #iFileNo, "^PW900"
                    Print #iFileNo, "^LH30,25"
                End If
                AddtoRowZ0 XA(i, 6), k Mod 2
                AddtoRowZ1 "Code: " & IIf(Len(FNS(XA(i, 1))) < 16, FNS(XA(i, 8)), FNS(XA(i, 1))), k Mod 2
                AddtoRowZ2 Space(19 - Len(XA(i, 3))) & XA(i, 3), k Mod 2
                AddtoRowZ3 Left(XA(i, 2), 24), k Mod 2
               ' tmp = FNS(XA(i, 10))
                AddtoRowZ5 FNS(XA(i, 10)), k Mod 2
                tmp = Left(oPC.Configuration.DefaultStore.Description, 32)
                tmp = tmp & Space(32 - Len(tmp))
                AddtoRowZ6 tmp & "  " & Format(Date, "MM/YY"), k Mod 2
                k = k + 1
                If k Mod 2 = 0 Then
                    Print #iFileNo, "^XZ"
                End If
            Next j
        End If
    Next
    If k Mod 2 = 1 Then
        Print #iFileNo, "^XZ"
    End If
    Close #iFileNo

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.PrintToZebra600", , , , "strPOS,strFileNo", Array(strPos, strFileno)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.PrintToZebra600"
End Sub


Private Sub AddtoRowZ0(pText As String, pPos As Integer)
    On Error GoTo errHandler
    If pPos = 0 Then
        Print #iFileNo, "^FO60,40^BEN,40,Y,N,Y^FD" & pText & "^FS"
    Else
        Print #iFileNo, "^FO480,40^BEN,40,Y,N,Y^FD" & pText & "^FS"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRowZ0(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRowZ0(pText,pPos)", Array(pText, pPos)
End Sub
Private Sub AddtoRowZ1(pText As String, pPos As Integer)
    On Error GoTo errHandler
    If pPos = 0 Then
        Print #iFileNo, "^FO30,10^A0,,25^FD" & pText & "^FS"
    Else
        Print #iFileNo, "^FO450,10^A0,,25^FD" & pText & "^FS"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRowZ1(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRowZ1(pText,pPos)", Array(pText, pPos)
End Sub
Private Sub AddtoRowZ2(pText As String, pPos As Integer)
    On Error GoTo errHandler
    If pPos = 0 Then
        Print #iFileNo, "^FO320,60^A0,,25^FD" & Trim(pText) & "^FS"
    Else
        Print #iFileNo, "^FO700,60^A0,,25^FD" & Trim(pText) & "^FS"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRowZ2(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRowZ2(pText,pPos)", Array(pText, pPos)
End Sub
Private Sub AddtoRowZ3(pText As String, pPos As Integer)
    On Error GoTo errHandler
    If pPos = 0 Then
        Print #iFileNo, "^FO30,110^A0,,25^FD" & pText & "^FS"
    Else
        Print #iFileNo, "^FO450,110^A0,,25^FD" & pText & "^FS"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRowZ3(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRowZ3(pText,pPos)", Array(pText, pPos)
End Sub
Private Sub AddtoRowZ4(pText As String, pPos As Integer)
    On Error GoTo errHandler
    If pPos = 0 Then
        Print #iFileNo, "^FO30,140^A0,,25^FD" & pText & "^FS"
    Else
        Print #iFileNo, "^FO450,140^A0,,25^FD" & pText & "^FS"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRowZ4(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRowZ4(pText,pPos)", Array(pText, pPos)
End Sub
Private Sub AddtoRowZ5(pText As String, pPos As Integer)
    On Error GoTo errHandler
    If pPos = 0 Then
        Print #iFileNo, "^FO30,137^A0,,25^FD" & pText & "^FS"
    Else
        Print #iFileNo, "^FO450,137^A0,,25^FD" & pText & "^FS"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRowZ5(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRowZ5(pText,pPos)", Array(pText, pPos)
End Sub
Private Sub AddtoRowZ6(pText As String, pPos As Integer)
    On Error GoTo errHandler
    If pPos = 0 Then
        Print #iFileNo, "^FO30,163^A0,,25^FD" & pText & "^FS"
    Else
        Print #iFileNo, "^FO450,163^A0,,25^FD" & pText & "^FS"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRowZ6(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRowZ6(pText,pPos)", Array(pText, pPos)
End Sub



Private Sub AddtoRowZ1n(pText As String, pPos As Integer)
    On Error GoTo errHandler
    If pPos = 0 Then
        Print #iFileNo, "^FO30,10^A0,,25^FD" & pText & "^FS"   '^AFN,22,15
    Else
        Print #iFileNo, "^FO450,10^A0,,25^FD" & pText & "^FS"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRowZ1n(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRowZ1n(pText,pPos)", Array(pText, pPos)
End Sub
Private Sub AddtoRowZ2n(pText As String, pPos As Integer)
    On Error GoTo errHandler
    If pPos = 0 Then
        Print #iFileNo, "^FO320,45^A0,,25^FD" & Trim(pText) & "^FS"
    Else
        Print #iFileNo, "^FO700,45^A0,,25^FD" & Trim(pText) & "^FS"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRowZ2n(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRowZ2n(pText,pPos)", Array(pText, pPos)
End Sub
Private Sub AddtoRowZ3n(pText As String, pPos As Integer)
    On Error GoTo errHandler
    If pPos = 0 Then
        Print #iFileNo, "^FO30,85^A0,,25^FD" & pText & "^FS"
    Else
        Print #iFileNo, "^FO450,85^A0,,25^FD" & pText & "^FS"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRowZ3n(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRowZ3n(pText,pPos)", Array(pText, pPos)
End Sub
Private Sub AddtoRowZ4n(pText As String, pPos As Integer)
    On Error GoTo errHandler
    If pPos = 0 Then
        Print #iFileNo, "^FO30,115^A0,,25^FD" & pText & "^FS"
    Else
        Print #iFileNo, "^FO450,115^A0,,25^FD" & pText & "^FS"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRowZ4n(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRowZ4n(pText,pPos)", Array(pText, pPos)
End Sub
Private Sub AddtoRowZ5n(pText As String, pPos As Integer)
    On Error GoTo errHandler
    If pPos = 0 Then
        Print #iFileNo, "^FO30,125^A0,,25^FD" & pText & "^FS"
    Else
        Print #iFileNo, "^FO450,125^A0,,25^FD" & pText & "^FS"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRowZ5n(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRowZ5n(pText,pPos)", Array(pText, pPos)
End Sub
Private Sub AddtoRowZ6n(pText As String, pPos As Integer)
    On Error GoTo errHandler
    If pPos = 0 Then
        Print #iFileNo, "^FO30,155^A0,,25^FD" & pText & "^FS"
    Else
        Print #iFileNo, "^FO450,155^A0,,25^FD" & pText & "^FS"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.AddtoRowZ5n(pText,pPos)", Array(pText, pPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.AddtoRowZ6n(pText,pPos)", Array(pText, pPos)
End Sub

Private Sub PrintLabelsToZebraTLP()
    On Error GoTo errHandler
Dim i, j, k, m As Integer
Dim oSQL As New z_SQL
Dim oMD As New z_ReportMetadata
Dim strDensity As String
Dim rpt As DDActiveReports2.ActiveReport
Dim s As String
Dim tf As New z_TextFile
Dim x As Printer
Dim Oldprinter As Object
Dim sTableName As String
Dim sec As DDActiveReports2.Section
Dim ctl As DDActiveReports2.DataControl

  '  MsgBox "Start of label printing"

    oSQL.RunSQL ("DELETE FROM tLABEL WHERE L_Workstation = '" & oPC.WorkstationName & "'")
    k = 0
    For i = 1 To XA.UpperBound(1)
        If FNN(XA(i, 4)) > 0 Then
            For j = 1 To FNN(XA(i, 4))
                k = k + 1
                oSQL.RunSQL ("INSERT INTO tLABEL (L_Workstation,L_EANF,L_EAN,L_Price,L_Description,L_DefaultStore, " _
                    & "L_WithBarcode,L_MainCategoryCode,L_ProductTypeCode,L_DateRecordAdded,L_MainAuthor,L_DateLabelPrinted, " _
                    & "L_LastDateDelivered,L_AllCategories,L_Length,L_Width,L_Dimensions,L_Code) " _
                    & " VALUES ('" & oPC.WorkstationName & "','" & IIf(Len(FNS(XA(i, 1))) < 16, FNS(XA(i, 8)), FNS(XA(i, 1))) & "','" & FNS(XA(i, 8)) & "','" & FNS(XA(i, 3)) & "','" _
                        & Replace(Left(XA(i, 2), 40), "'", "''") & "','" & oPC.Configuration.DefaultStore.Description & "'," _
                        & IIf(XA(i, 7) = True, 1, 0) & ",'" & FNS(XA(i, 11)) & "','" & FNS(XA(i, 14)) & "','" _
                        & Format(FNS(XA(i, 16)), "MM/YY") & "','" & Replace(Left(XA(i, 13), 30), "'", "''") & "','" _
                        & Format(Date, "MM/YY") & "','" & FNS(XA(i, 17)) & "','" & Left(FNS(XA(i, 18)), 20) & "','" _
                        & FNS(XA(i, 19)) & "','" & FNS(XA(i, 20)) & "','" & FNS(XA(i, 22)) & "','" & FNS(XA(i, 21)) & "')")
            Next j
        End If
    Next
   
    m = 0
    For i = 1 To XA.UpperBound(1)
        m = m + FNN(XA(i, 5))
    Next

   
    If m > 0 Then
        oSQL.RunSQL ("INSERT INTO tLABEL (L_Workstation,L_Description) VALUES ('" & oPC.WorkstationName & "','END OF NEW STOCK - REPRICING FOLLOWS BELOW')")
        If k Mod 2 = 0 Then
            oSQL.RunSQL ("INSERT INTO tLABEL (L_Workstation,L_Description) VALUES ('" & oPC.WorkstationName & "','END OF NEW STOCK - REPRICING FOLLOWS BELOW')")
        End If
    End If
    
    For i = 1 To XA.UpperBound(1)
        If FNN(XA(i, 5)) > 0 Then
            For j = 1 To FNN(XA(i, 5))
                oSQL.RunSQL ("INSERT INTO tLABEL (L_Workstation,L_EANF,L_EAN,L_Price,L_Description,L_DefaultStore, " _
                    & "L_WithBarcode,L_MainCategoryCode,L_ProductTypeCode,L_DateRecordAdded,L_MainAuthor,L_DateLabelPrinted, " _
                    & "L_LastDateDelivered,L_AllCategories,L_Length,L_Width,L_Dimensions,L_Code) " _
                    & " VALUES ('" & oPC.WorkstationName & "','" & IIf(Len(FNS(XA(i, 1))) < 16, FNS(XA(i, 8)), FNS(XA(i, 1))) & "','" & FNS(XA(i, 8)) & "','" & FNS(XA(i, 3)) & "','" _
                        & Replace(Left(XA(i, 2), 40), "'", "''") & "','" & oPC.Configuration.DefaultStore.Description & "'," _
                        & IIf(XA(i, 7) = True, 1, 0) & ",'" & FNS(XA(i, 11)) & "','" & FNS(XA(i, 14)) & "','" _
                        & Format(FNS(XA(i, 16)), "MM/YY") & "','" & Replace(Left(XA(i, 13), 30), "'", "''") & "','" _
                        & Format(Date, "MM/YY") & "','" & FNS(XA(i, 17)) & "','" & Left(FNS(XA(i, 18)), 20) & "','" _
                        & FNS(XA(i, 19)) & "','" & FNS(XA(i, 20)) & "','" & FNS(XA(i, 22)) & "','" & FNS(XA(i, 21)) & "')")
            Next j
        End If
    Next
    
    Set rpt = New DDActiveReports2.ActiveReport
    tf.OpenTextFileToRead oPC.SharedFolderRoot & "\Templates\Label_BB.XML"
    s = tf.ReadWholeFile
    tf.CloseTextFile
    oMD.LoadMetadataToXML s
    
    oMD.ConnectionString = oPC.ConnectionString
    oMD.RecordSource = "SELECT *,ISNULL(L_Description, '')  + ' - ' + ISNULL(L_MainAuthor, '') as DescriptionCombo FROM tLabel WHERE L_WORKSTATION = '" & oPC.WorkstationName & "' ORDER BY L_SEQUENCE "
  '  oMD.AddSortedVolumn "L_Sequence", "ASC"
 
  
    rpt.LoadLayout StringToByteArray(oMD.Layout_fromXML, False, True)
''remember existing default printer so we can reset it afterwards : (Printer.DeviceName here refers to the default printer device)
'    For Each x In Printers
'        If x.DeviceName = Printer.DeviceName Then
'            Set Oldprinter = x
'            Exit For
'        End If
'    Next

    
    For Each x In Printers
       If InStr(1, UCase(x.DeviceName), "LABEL") > 0 Then
          Set Printer = x
          Exit For
       End If
    Next
   
    On Error Resume Next
    rpt.Printer.DeviceName = Printer.DeviceName
    rpt.Printer = Printer
    
    rpt.PrintReport False
'    If Err.Number <> 0 Then
'        Err.Clear
'        Dim n As String
'        n = ParseDeviceName(x.DeviceName)
'        rpt.Printer.DeviceName = n
'    End If
'    On Error GoTo errHandler
'
'    rpt.PrintReport False
'
'    Set rpt = Nothing
'    Set Printer = Oldprinter
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.PrintLabelsToZebraTLP", , , , "line number", Array(Erl())
End Sub
'Private Sub PrintLabelsToZebraTLP()
'    On Error GoTo errHandler
'Dim i, j, k, m As Integer
'Dim oSQL As New z_SQL
'Dim oMD As New z_ReportMetadata
'Dim strDensity As String
'Dim rpt As DDActiveReports2.ActiveReport
'Dim s As String
'Dim tf As New z_TextFile
'Dim x As Printer
'Dim Oldprinter As Object
'Dim sTableName As String
'Dim sec As DDActiveReports2.Section
'Dim ctl As DDActiveReports2.DataControl
'
'    sTableName = "dbo.tLabels_" & Replace(oPC.WorkstationName, "-", "")
'
'    oSQL.RunProc "PrepareLabelTable", Array(sTableName), ""
'    oSQL.RunSQL ("DELETE FROM " & sTableName & " WHERE L_Workstation = '" & oPC.WorkstationName & "'")
'    k = 0
'    For i = 1 To XA.UpperBound(1)
'        If FNN(XA(i, 4)) > 0 Then
'            For j = 1 To FNN(XA(i, 4))
'                k = k + 1
'                oSQL.RunSQL ("INSERT INTO " & sTableName & "(L_Workstation,L_EANF,L_EAN,L_Price,L_Description,L_DefaultStore,L_WithBarcode,L_MainCategoryCode,L_ProductTypeCode,L_DateRecordAdded,L_MainAuthor,L_DateLabelPrinted,L_LastDateDelivered,L_AllCategories) VALUES ('" & oPC.WorkstationName & "','" & FNS(XA(i, 1)) & "','" & FNS(XA(i, 8)) & "','" & FNS(XA(i, 3)) & "','" & Replace(Left(XA(i, 2), 40), "'", "''") & "','" & oPC.Configuration.DefaultStore.Description & "'," & IIf(XA(i, 7) = True, 1, 0) & ",'" & FNS(XA(i, 11)) & "','" & FNS(XA(i, 14)) & "','" & Format(FNS(XA(i, 16)), "MM/YY") & "','" & Left(FNS(XA(i, 13)), 30) & "','" & Format(Date, "MM/YY") & "','" & FNS(XA(i, 17)) & "','" & Left(FNS(XA(i, 18)), 20) & "')")
'            Next j
'        End If
'    Next
'
'    m = 0
'    For i = 1 To XA.UpperBound(1)
'        m = m + FNN(XA(i, 5))
'    Next
'
'
'    If m > 0 Then
'        oSQL.RunSQL ("INSERT INTO " & sTableName & "(L_Workstation,L_Description) VALUES ('" & oPC.WorkstationName & "','END OF NEW STOCK - REPRICING FOLLOWS BELOW')")
'        If k Mod 2 = 0 Then
'            oSQL.RunSQL ("INSERT INTO " & sTableName & "(L_Workstation,L_Description) VALUES ('" & oPC.WorkstationName & "','END OF NEW STOCK - REPRICING FOLLOWS BELOW')")
'        End If
'    End If
'
'    For i = 1 To XA.UpperBound(1)
'        If FNN(XA(i, 5)) > 0 Then
'            For j = 1 To FNN(XA(i, 5))
'                oSQL.RunSQL ("INSERT INTO " & sTableName & "(L_Workstation,L_EANF,L_EAN,L_Price,L_Description,L_DefaultStore,L_WithBarcode,L_MainCategoryCode,L_ProductTypeCode,L_DateRecordAdded,L_MainAuthor,L_DateLabelPrinted,L_LastDateDelivered,L_AllCategories) VALUES ('" & oPC.WorkstationName & "','" & FNS(XA(i, 1)) & "','" & FNS(XA(i, 8)) & "','" & FNS(XA(i, 3)) & "','" & Replace(Left(XA(i, 2), 40), "'", "''") & "','" & oPC.Configuration.DefaultStore.Description & "'," & IIf(XA(i, 7) = True, 1, 0) & ",'" & FNS(XA(i, 11)) & "','" & FNS(XA(i, 14)) & "','" & Format(FNS(XA(i, 16)), "MM/YY") & "','" & Left(FNS(XA(i, 13)), 30) & "','" & Format(Date, "MM/YY") & "','" & FNS(XA(i, 17)) & "','" & Left(FNS(XA(i, 18)), 20) & "')")
'            Next j
'        End If
'    Next
'
'    Set rpt = New DDActiveReports2.ActiveReport
'    tf.OpenTextFileToRead oPC.SharedFolderRoot & "\Templates\Label_BB.XML"
'    s = tf.ReadWholeFile
'    tf.CloseTextFile
'    oMD.LoadMetadataToXML s
'
'    oMD.ConnectionString = oPC.ConnectionString
'    oMD.RecordSource = "(SELECT *,ISNULL(L_Description, '') + ' - ' + ISNULL(L_MainAuthor, '') as DescriptionCombo FROM " & sTableName & ")"
'
'    rpt.LoadLayout StringToByteArray(oMD.Layout_fromXML, False, True)
'
'    strDevicename = GetSetting("PBKS", "Printers", "Labels", "")
'
''remember existing default printer so we can reset it afterwards
'        For Each x In Printers
'        If x.DeviceName = Printer.DeviceName Then
'            Set Oldprinter = x
'            Exit For
'        End If
'    Next
''
'    For Each x In Printers
'       If InStr(1, x.DeviceName, "Label") > 0 Then
'          Set Printer = x
'          Exit For
'       End If
'    Next
'    rpt.Printer.DeviceName = Printer.DeviceName
'  '  rpt.PageSettings.Gutter = 40
'    rpt.PrintReport False
'    Set rpt = Nothing
'    Set Printer = Oldprinter
'
'    Exit Sub
'errHandler:
'    ErrPreserve
'    If err = 5707 Then
'        err.Clear
'        rpt.Printer.SetupDialog
'        strDevicename = rpt.Printer.DeviceName
'        SaveSetting "PBKS", "Printers", "Labels", strDevicename
'        Resume
'    End If
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintLabels.PrintLabelsToZebraTLP", , , , "Devicename:", Array(strDevicename)
'End Sub

Private Sub cmdPrinter_Click()
    On Error GoTo errHandler
'Dim f As frmSelectPrinter

'f.Show vbModal
'Dim X As Printer
'    For Each X In Printers
'        If X.DeviceName = Printer.DeviceName Then
'            Set Oldprinter = X
'            Exit For
'        End If
'    Next
'
'    For Each X In Printers
'       If InStr(1, X.DeviceName, "Label") > 0 Then
'          Set Printer = X
'          Exit For
'       End If
'    Next
'    CD1.FileName = GetSetting("PBKS", "Printers", "Labels", "")
'    CD1.DialogTitle = "Select labels printer"
'    CD1.ShowPrinter
'    strDevicename = Printer.DeviceName
'    SaveSetting "PBKS", "Printers", "Labels", strDevicename
'
'
'    Set Printer = Oldprinter

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintLabels.cmdPrinter_Click", , EA_NORERAISE
    HandleError
End Sub

