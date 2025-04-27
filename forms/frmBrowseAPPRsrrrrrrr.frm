VERSION 5.00
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmBrowseAPPRs 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Browse appro returns"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrowseAPPRsrrrrrrr.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   570
      Left            =   3990
      Picture         =   "frmBrowseAPPRsrrrrrrr.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4950
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   90
      TabIndex        =   2
      Top             =   -75
      Width           =   4920
      Begin VB.CommandButton cmdFind1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Height          =   405
         Left            =   4290
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Click to find all customers matching the retrictions entered."
         Top             =   810
         UseMaskColor    =   -1  'True
         Width           =   570
      End
      Begin VB.TextBox txtArg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   450
         Left            =   2310
         TabIndex        =   0
         ToolTipText     =   "Enter product code, reference A/C/ no. or start of customer name. Hit ENTER to fetch."
         Top             =   465
         Width           =   1785
      End
      Begin CoolButtonControl.CoolButton cbSince 
         Height          =   585
         Left            =   165
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "You can find orders by their issue dates."
         Top             =   390
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1032
         BackColor       =   14737632
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
         Caption         =   "&Since: Last week"
         Style           =   1
         BackStyle       =   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Search for . . ."
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   2310
         TabIndex        =   3
         Top             =   210
         Width           =   1980
      End
   End
   Begin VB.PictureBox Grid 
      Height          =   3345
      Left            =   135
      ScaleHeight     =   3285
      ScaleWidth      =   4830
      TabIndex        =   7
      Top             =   1560
      Width           =   4890
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TIP: You can use * as wildcard in searches"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   105
      TabIndex        =   5
      Top             =   4980
      Width           =   3315
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      Height          =   3420
      Left            =   105
      Top             =   1515
      Width           =   4950
   End
End
Attribute VB_Name = "frmBrowseAPPRs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cAPPR As c_APPRs
Dim dDel As d_DEL
Dim lngTPID As Long
Dim strRef As String
Dim enSince As enumSince
Dim dteDate1 As Date
Dim dteDate2 As Date
Dim strDate1 As String
Dim strDate2 As String
Dim blnNoRecordsReturned As Boolean
Dim flgLoading As Boolean
Dim ofrm As frmAPPRPreview
Dim XA As New XArrayDB


Private Sub cbSince_Click()
    enSince = OptionLoop(enSince, 5)
    cbSince.Caption = TranslateSince(CInt(enSince))
    txtArg = ""
    txtArg.SetFocus
End Sub
Private Sub cbSince_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Find
        LoadArray
        Grid.ReBind
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFind1_Click()
        Find
        LoadArray
        Grid.ReBind
End Sub

Private Sub Grid_GotFocus()
    Shape1.Visible = True
End Sub
Private Sub Grid_LostFocus()
    Shape1.Visible = False
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Grid_DblClick
    End If
End Sub


Private Sub txtArg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Find
        LoadArray
        Grid.ReBind
    End If
End Sub

Private Function ArgIsProductCode() As Boolean
Dim oProdCode As New z_ProdCode

    If left(txtArg, 1) = "#" Then
        ArgIsProductCode = True
    Else
        oProdCode.Load txtArg
        If oProdCode.IsCode Or oProdCode.IsISBN Then
            ArgIsProductCode = True
        End If
    End If
End Function
Private Sub SetDateArgs()
    Select Case enSince
    Case enAny
        dteDate1 = CDate("1995-01-01")
        dteDate2 = DateAdd("d", 1, Date)
    Case enWeek
        dteDate1 = DateAdd("d", -7, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case enMonth
        dteDate1 = DateAdd("m", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case enQuarter
        dteDate1 = DateAdd("q", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case enYear
        dteDate1 = DateAdd("yyyy", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    End Select

End Sub

Private Sub Find()
Dim bNotFound As Boolean
Dim frm As frmBrowseCustomers2
Dim lngTPID As Long
    Screen.MousePointer = vbHourglass
    On Error GoTo ERR_Handler
    bNotFound = False
    If txtArg > " " Then
        enSince = 1
        cbSince.Caption = TranslateSince(1)
        If ArgIsProductCode Then
            'Search for product code
            Set cAPPR = Nothing
            Set cAPPR = New c_APPRs
            cAPPR.Load bNotFound, 0, "", "", dteDate1, dteDate2, , txtArg
            Exit Sub
        End If
        'Search for Reference
        Set cAPPR = Nothing
        Set cAPPR = New c_APPRs
        cAPPR.Load bNotFound, 0, "", "", dteDate1, dteDate2, , , txtArg
        If bNotFound Then
            'Search for customer by ACCNO
            Set cAPPR = Nothing
            Set cAPPR = New c_APPRs
            cAPPR.Load bNotFound, 0, txtArg, "", dteDate1, dteDate2
            If bNotFound Then
               Set frm = New frmBrowseCustomers2
               frm.Component txtArg
               frm.Show vbModal
               lngTPID = frm.CustomerID
               Unload frm
               If lngTPID > 0 Then
                    Set cAPPR = Nothing
                    Set cAPPR = New c_APPRs
                    cAPPR.Load bNotFound, lngTPID, "", "", dteDate1, dteDate2
                    txtArg = ""
               End If
            End If
        End If
    Else
        SetDateArgs
        cAPPR.Load bNotFound, 0, "", "", dteDate1, dteDate2
    End If
    Grid.SetFocus

EXIT_Handler:
    Screen.MousePointer = vbDefault
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub


Private Sub cmdFind_LostFocus()
    LoadControls
End Sub


Private Sub Form_Load()
   ' Set tlSupplier = New z_TextList
    Set cAPPR = New c_APPRs
    Me.top = 50
    Me.left = 50
    Me.Width = 5300
    Me.Height = 5970
    LoadControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Set tlSupplier = Nothing
    Set cAPPR = Nothing
    Set dDel = Nothing
    Set ofrm = Nothing
End Sub


Private Sub LoadControls()
    flgLoading = True
    txtArg = ""
    lngTPID = 0
    enSince = enWeek
    cbSince.Caption = TranslateSince(CInt(enSince))
    flgLoading = False
End Sub

Private Sub LoadArray()
Dim objItem As d_APPR
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.Clear
    XA.ReDim 1, cAPPR.Count, 1, 6
    For i = 1 To cAPPR.Count
        With objItem
            XA.Value(i, 1) = cAPPR(i).TPName
            XA.Value(i, 2) = cAPPR(i).DocCode
            XA.Value(i, 3) = cAPPR(i).TRDateF
            XA.Value(i, 4) = cAPPR(i).DateForSort
            XA.Value(i, 5) = cAPPR(i).TRID & "K"
            XA.Value(i, 6) = cAPPR(i).StatusF
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 4, XORDER_DESCEND, XTYPE_STRING
    Grid.Array = XA
End Sub

Private Sub Grid_DblClick()
Dim lngID As Long
Dim blnEdit As Boolean
    Set ofrm = New frmAPPRPreview
    lngID = val(XA(Grid.Bookmark, 5))
    ofrm.Component lngID    ', False
    ofrm.Show
End Sub
Private Sub Grid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If XA(Bookmark, 6) = "VOID" Or XA(Bookmark, 6) = "CANCELLED" Then
        RowStyle.BackColor = &HC0C0C0
    End If
    If XA(Bookmark, 6) = "IN PROCESS" Then
        RowStyle.BackColor = &H80FF80
    End If
    If XA(Bookmark, 6) = "COMPLETE" Then
        RowStyle.BackColor = &HFFFFC0
    End If

End Sub

