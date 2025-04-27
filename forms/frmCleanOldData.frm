VERSION 5.00
Begin VB.Form frmCleanOldData 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Clean old data from database"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkCustomer 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Also remove all customerrecords that are not associated with any transaction and do not have an account number."
      ForeColor       =   &H8000000D&
      Height          =   720
      Left            =   615
      TabIndex        =   5
      Top             =   2400
      Width           =   3240
   End
   Begin VB.CheckBox chkRemoveStock 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Also remove all product records that are not associated with any transaction and have not been counted."
      ForeColor       =   &H8000000D&
      Height          =   720
      Left            =   600
      TabIndex        =   4
      Top             =   1530
      Width           =   3240
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D7D1BF&
      Caption         =   "&Cancel"
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
      Left            =   1365
      Picture         =   "frmCleanOldData.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cboGo 
      BackColor       =   &H00D7D1BF&
      Caption         =   "&Remove old data"
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
      Left            =   1365
      Picture         =   "frmCleanOldData.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3570
      Width           =   1935
   End
   Begin VB.ComboBox cboStktke 
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
      Height          =   360
      Left            =   795
      TabIndex        =   0
      Top             =   765
      Width           =   2715
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Select stock take, prior to which transactions will be removed"
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
      Height          =   630
      Left            =   780
      TabIndex        =   1
      Top             =   180
      Width           =   2910
   End
End
Attribute VB_Name = "frmCleanOldData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tlStocktakes As New z_TextList
Dim lngStockTakeID As Long
Dim lngOPID As Long

Public Sub Component(pOPID As Long)
    lngOPID = pOPID
End Sub
Private Sub cboGo_Click()
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim lngResult As Long
Dim lngPosition As Long
    Screen.MousePointer = vbHourglass
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    Set cmd.ActiveConnection = oPC.CO
    cmd.CommandText = "RemoveOldTRs"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@pPID", adVarChar, adParamInput, 10, ReverseDate(cboStktke))
    cmd.Parameters.Append prm

    Set prm = cmd.CreateParameter("@pTPID", adInteger, adParamInput, , gSTAFFID)
    cmd.Parameters.Append prm
    
    Set prm = cmd.CreateParameter("@ErrCode", adInteger, adParamOutput)
    cmd.Parameters.Append prm

    Set prm = cmd.CreateParameter("@pPosition", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
   ' On Error Resume Next
    cmd.Execute
    On Error GoTo errHandler
    If cmd.Parameters(2).Value <> 0 Then
        Err.Raise 9999, "SQL", "Error in RemoveOldTRs: Error code = " & cmd.Parameters(2).Value & ", Position = " & cmd.Parameters(3).Value
    End If
    
    If Me.chkRemoveStock = 1 Then
        Set cmd = Nothing
        Set cmd = New ADODB.Command
        cmd.CommandTimeout = 0
        Set cmd.ActiveConnection = oPC.CO
        cmd.CommandText = "DeleteUnusedProducts"
        cmd.CommandType = adCmdStoredProc
        cmd.Execute
    End If
    
    If Me.chkCustomer = 1 Then
        Set cmd = Nothing
        Set cmd = New ADODB.Command
        cmd.CommandTimeout = 0
        Set cmd.ActiveConnection = oPC.CO
        cmd.CommandText = "DeleteUnusedProducts"
        cmd.CommandType = adCmdStoredProc
        cmd.Execute
    End If
    
    Screen.MousePointer = vbDefault
    MsgBox "Operation completed", , "Status"
    Unload Me

    Exit Sub
errHandler:
    ErrorIn "frmCleanOldData.cboGo_Click", , EA_NORERAISE
    HandleError
End Sub
'[RemoveInactiveCustomers]

Private Sub cboStktke_Validate(Cancel As Boolean)
    lngStockTakeID = tlStocktakes.Key(cboStktke)
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
    tlStocktakes.Load ltStockTake
    LoadCombo Me.cboStktke, tlStocktakes
End Sub
