VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmStartDE 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Day end date settings"
   ClientHeight    =   5520
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtLastDate 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   690
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   750
      Width           =   1905
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   390
      Left            =   900
      TabIndex        =   3
      Top             =   2565
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   -2147483635
      CalendarTitleForeColor=   -2147483635
      Format          =   61472769
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
      Height          =   465
      Left            =   1005
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4260
      Width           =   1230
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Start"
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
      Height          =   465
      Left            =   1020
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3750
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "The dayend was last run on:"
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
      Left            =   315
      TabIndex        =   4
      Top             =   345
      Width           =   2625
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enter the nominal date of the dayend."
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
      Left            =   345
      TabIndex        =   2
      Top             =   1935
      Width           =   2625
   End
End
Attribute VB_Name = "frmStartDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim flgGo
Dim tlOperators As z_TextList
Dim lngOperatorID As Long
Dim strMessage As String
Dim dteLast As Date
Dim dteThis As Date

Private Sub CancelButton_Click()
    Me.Hide
End Sub
Public Sub Component(pLastDate As Date, pDate As Date)
    dteLast = pLastDate
    dteThis = pDate
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
    txtLastDate = Format(dteLast, "dd/mm/yyyy")
    DTPicker1.Value = Format(dteThis, "dd/mm/yyyy")
    DTPicker1.Refresh
End Sub
Public Property Get Response() As Boolean
    Response = flgGo
End Property

Private Sub OKButton_Click()
Dim frmS As frmSecurity
Dim strName As String
    Set frmS = New frmSecurity
    frmS.Show vbModal
    Unload frmS
    If oPC.Configuration.Staff.GetLevel(oPC.CurrentSecurityCode, strName, lngOperatorID) < 2 Then
        MsgBox "You do not have security to start this operation.", vbExclamation, "Denied"
        Exit Sub
    End If

    If MsgBox("You are starting the dayend for: " & Format(Me.DTPicker1.Value, "dddd,dd mmmm yyyy") & "?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
        flgGo = True
    Else
        flgGo = False
    End If
    Me.Hide

End Sub
