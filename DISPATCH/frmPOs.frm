VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{A19332D7-D707-4A30-9F38-796D120AF5B3}#1.1#0"; "BtnPlus1.ocx"
Begin VB.Form frmPOs 
   BackColor       =   &H00C8B9B3&
   Caption         =   "Recent purchase orders"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin TrueOleDBGrid60.TDBGrid POGrid 
      Height          =   3270
      Left            =   120
      OleObjectBlob   =   "frmPOs.frx":0000
      TabIndex        =   0
      Top             =   540
      Width           =   7800
   End
   Begin ButtonPlusCtl.ButtonPlus cbSince 
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
      BackStyle       =   0
      FocusStyle      =   0
      BorderStyle     =   4
      Caption         =   "&Since: Last week"
      MaskColor       =   12632256
      BackColor       =   12632256
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmPOs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Purchase orders ----
Dim XA As XArrayDB
Dim cSO As c_POs
Dim enSince As enumSince
Dim strPeriod As String
Dim flgLoading As Boolean

Private Sub cbSince_Click()
    If flgLoading Then Exit Sub
    enSince = OptionLoop(enSince, 5)
    cbSince.Caption = TranslateSince(CInt(enSince))
End Sub
Public Function TranslateSince(pIn As Integer) As String
    Select Case pIn
    Case 2
        TranslateSince = "&Since: Last week"
    Case 3
        TranslateSince = "&Since: Last month"
    Case 4
        TranslateSince = "&Since: Last quarter"
    Case 5
        TranslateSince = "&Since: Last year"
    Case 1
        TranslateSince = "&Since: <any date>"
    End Select
End Function

Private Sub RefreshRecs()
    XA.Clear
    GetRecs
    POGrid.ReBind
End Sub

Private Sub GetRecs()
Dim bNoRecs As Boolean
Dim bMoreToCome As Boolean
Dim strResult As String

    Set cSO = Nothing
    Set cSO = New c_POs
    If strPeriod = "week" Then
        cSO.Load bNoRecs, , , , DateAdd("w", -1, Date), DateAdd("d", 1, Date)
    ElseIf strPeriod = "two weeks" Then
        cSO.Load bNoRecs, , , , DateAdd("w", -2, Date), DateAdd("d", 1, Date)
    ElseIf strPeriod = "month" Then
        cSO.Load bNoRecs, , , , DateAdd("m", -1, Date), DateAdd("d", 1, Date)
    ElseIf strPeriod = "three months" Then
        cSO.Load bNoRecs, , , , DateAdd("m", -3, Date), DateAdd("d", 1, Date)
    End If

    LoadArray

    strResult = XA.UpperBound(1) & " records"
    
  '  lblResult.Caption = strResult & " (" & strPeriod & ")"
End Sub

Private Sub LoadArray()
Dim objItem As d_PO
Dim itmList As ListItem
Dim lngIndex As Long
    
    XA.ReDim 1, cSO.Count, 1, 7
    For lngIndex = 1 To cSO.Count
        With objItem
            Set objItem = cSO.Item(lngIndex)
            XA.Value(lngIndex, 1) = objItem.TPName
            XA.Value(lngIndex, 2) = objItem.DocCode
            XA.Value(lngIndex, 3) = objItem.DocDateF
            'XA.Value(lngIndex, 4) = objItem.di
            XA.Value(lngIndex, 5) = 0
            XA.Value(lngIndex, 6) = objItem.DocDate
            XA.Value(lngIndex, 7) = objItem.TRID
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 6, XORDER_DESCEND, XTYPE_DATE
    POGrid.Array = XA

End Sub

