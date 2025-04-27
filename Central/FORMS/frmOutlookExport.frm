VERSION 5.00
Begin VB.Form frmOutlookExport 
   Caption         =   "Export to Outlook"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   4620
   Begin VB.TextBox txtOLFolder 
      Height          =   330
      Left            =   2160
      TabIndex        =   6
      Top             =   3150
      Width           =   2115
   End
   Begin VB.CommandButton cmdOutlook_Export 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export"
      Enabled         =   0   'False
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
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3690
      Width           =   1000
   End
   Begin VB.TextBox txtPrefix 
      Alignment       =   2  'Center
      Enabled         =   0   'False
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
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Text            =   "AB"
      Top             =   2580
      Width           =   795
   End
   Begin VB.TextBox txtPartsize 
      Alignment       =   2  'Center
      Enabled         =   0   'False
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
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Text            =   "100"
      Top             =   2040
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mode"
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
      Height          =   1545
      Left            =   1110
      TabIndex        =   0
      Top             =   300
      Width           =   2415
      Begin VB.PictureBox Picture 
         Height          =   945
         Left            =   135
         ScaleHeight     =   885
         ScaleWidth      =   2160
         TabIndex        =   8
         Top             =   360
         Width           =   2220
         Begin VB.OptionButton optAll 
            Caption         =   "Export all"
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
            Height          =   345
            Left            =   135
            TabIndex        =   10
            Top             =   75
            Value           =   -1  'True
            Width           =   1965
         End
         Begin VB.OptionButton optPart 
            Caption         =   "Export in parts"
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
            Height          =   345
            Left            =   135
            TabIndex        =   9
            Top             =   495
            Width           =   1965
         End
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Outlook folder name"
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
      Height          =   285
      Left            =   75
      TabIndex        =   7
      Top             =   3180
      Width           =   1950
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Folder name prefix"
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
      Height          =   405
      Left            =   180
      TabIndex        =   4
      Top             =   2610
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Part size"
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
      Height          =   405
      Left            =   1260
      TabIndex        =   2
      Top             =   2070
      Width           =   765
   End
End
Attribute VB_Name = "frmOutlookExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oS As New z_Search
Dim cCust As c_C_Customer

Public Sub component(pcCust As c_C_Customer)
    On Error GoTo errHandler
    Set cCust = pcCust
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOutlookExport.component(pcCust)", pcCust
End Sub


Private Sub cmdOutlook_Export_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    If optAll Then
        oS.ExportToOutlook txtOLFolder, cCust
    Else
        If CLng(txtPartsize) <= 1 Then
            MsgBox "The part size you have chosen is too small. ", vbInformation, "Can't export"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        oS.ExportToOutlookByPart txtOLFolder, cCust, CLng(txtPartsize), Trim(txtPrefix)
    End If
    Screen.MousePointer = vbDefault
    
    MsgBox "Export to Outlook finished.", vbInformation, "Status"
    
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOutlookExport.cmdOutlook_Export_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optAll_Click()
    On Error GoTo errHandler
    txtPartsize.Enabled = Not optAll
    txtPrefix.Enabled = Not optAll
    txtOLFolder.Enabled = optAll
    Me.cmdOutlook_Export.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOutlookExport.optAll_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optPart_Click()
    On Error GoTo errHandler
    txtPartsize.Enabled = optPart
    txtPrefix.Enabled = optPart
    txtOLFolder.Enabled = Not optPart
    Me.cmdOutlook_Export.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOutlookExport.optPart_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPartsize_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtPartsize)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOutlookExport.txtPartsize_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtOLFolder_Change()
    On Error GoTo errHandler
    cmdOutlook_Export.Enabled = (Len(txtOLFolder) > 5)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOutlookExport.txtOLFolder_Change", , EA_NORERAISE
    HandleError
End Sub

