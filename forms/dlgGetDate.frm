VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form dlgGetDate 
   BackColor       =   &H00D3D3CB&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Date of operation"
   ClientHeight    =   3060
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   405
      Left            =   405
      TabIndex        =   3
      Top             =   825
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16187393
      UpDown          =   -1  'True
      CurrentDate     =   36526
      MinDate         =   -73046
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   870
      Picture         =   "dlgGetDate.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1000
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00C4BCA4&
      Caption         =   "OK"
      CausesValidation=   0   'False
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1890
      Picture         =   "dlgGetDate.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   1000
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the effective date of the operation you have chosen to run."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   540
      Left            =   405
      TabIndex        =   2
      Top             =   255
      Width           =   3030
   End
End
Attribute VB_Name = "dlgGetDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim flgGo
Dim tlOperators As z_TextList
Dim lngOperatorID As Long
Dim strMessage As String

Private Sub CancelButton_Click()
    Me.Hide
End Sub
Public Sub SetDialog(val As String, Optional pD As Date)
    strMessage = val
    lbl1.Caption = val
    If CLng(pD) <> 0 Then
        DTPicker1.Value = pD
    End If
    Me.Show vbModal
End Sub
Friend Property Get OperatorID() As Long
    OperatorID = lngOperatorID
End Property

'Private Sub Combo1_Click()
'    lngOperatorID = tlOperators.Key(Combo1.Text)
'End Sub

Private Sub Form_Load()
 '   Set tlOperators = New z_TextList
 '   tlOperators.Load ltStaff
   ' tlOperators.Load  '"vAllStaffMembers ORDER BY SM_Name",
 '   LoadCombo Me.Combo1, tlOperators
'    lngOperatorID = tlOperators.Key(Combo1.Text)
    lbl1.Caption = strMessage
'    Me.DTPicker1.Value = Format(Date, "dd/mm/yyyy")
'   Me.DTPicker1.Refresh
   ' Me.Refresh
End Sub
Public Property Get Response() As Boolean
    Response = flgGo
End Property

Private Sub OKButton_Click()
Dim frmS As frmSecurity
Dim strName As String
    Set frmS = New frmSecurity
    frmS.Show vbModal
    If oPC.Configuration.Staff.GetLevel(oPC.CurrentSecurityCode, strName, lngOperatorID) < 2 Then
        MsgBox "You do not have security to start this operation.", vbExclamation, "Denied"
        Exit Sub
    End If

    If MsgBox(strMessage & Format(Me.DTPicker1.Value, "dd/mm/yyyy") & "?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
        flgGo = True
    Else
        flgGo = False
    End If
    Me.Hide

End Sub
