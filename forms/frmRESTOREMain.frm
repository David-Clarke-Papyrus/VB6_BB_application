VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Restore PBKS"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
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
      Height          =   330
      Left            =   1785
      TabIndex        =   3
      Top             =   1875
      Width           =   1125
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
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
      Height          =   330
      Left            =   1785
      TabIndex        =   2
      Top             =   1455
      Width           =   1125
   End
   Begin VB.CommandButton cmdRestore 
      BackColor       =   &H00D7D1BF&
      Caption         =   "RESTORE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   1545
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   1605
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
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
      Left            =   735
      TabIndex        =   5
      Top             =   1875
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Username"
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
      Left            =   735
      TabIndex        =   4
      Top             =   1500
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "This utility will restore the PBKS database and transaction log files to their original folder from the default backup device."
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
      Height          =   1095
      Left            =   285
      TabIndex        =   0
      Top             =   270
      Width           =   4095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnPapy As ADODB.Connection
Private strServerRootPath As String
Dim iFilenum2 As Integer
Dim strServer As String


Private Sub cmdRestore_Click()
On Error GoTo ERRH
    Screen.MousePointer = vbHourglass
    Set cnPapy = New ADODB.Connection
    cnPapy.Provider = "sqloledb"
    cnPapy.Open "Data Source=" & Me.Servername & ";Initial Catalog=MASTER;User Id=" & txtUsername & ";Password=" & txtPassword & "; Network Library=dbmssocn;"
    cnPapy.Execute "sp_detach_db 'NEWDB', 'true'"
    cnPapy.Execute "RESTORE DATABASE PBKS FROM DISK = '" & strServerRootPath & "\BU\PBKS.BAK' WITH NORECOVERY,REPLACE "
    cnPapy.Execute "RESTORE LOG PBKS FROM DISK = '" & strServerRootPath & "\BU\PBKS.BAK' WITH RECOVERY "
    'cnPapy.Execute "RESTORE LOG PBKS FROM DISK = '" & strServerRootPath & "\BU\PBKSLOG.BAK' WITH NORECOVERY"
    cnPapy.Close
    Screen.MousePointer = vbDefault
    MsgBox "Restore completed", vbInformation + vbOKOnly, "Status"

    Exit Sub
    
ERRH:
    MsgBox Error
    Exit Sub

End Sub

Public Property Get Servername() As String
Dim fs As New FileSystemObject
On Error GoTo ERRH
    strServerRootPath = fs.GetParentFolderName(App.Path)
    iFilenum2 = FreeFile
    Open strServerRootPath & "\ServerInstance.txt" For Input As #iFilenum2
    Line Input #iFilenum2, strServer

    Servername = strServer
    Close #iFilenum2
ERRH:
    Exit Property
    Resume
End Property

