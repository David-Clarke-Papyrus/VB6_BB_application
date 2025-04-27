VERSION 5.00
Begin VB.Form frmManageconnection 
   BackColor       =   &H00C8B9B3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dial-Up Networking Demo"
   ClientHeight    =   3525
   ClientLeft      =   5160
   ClientTop       =   1890
   ClientWidth     =   4110
   Icon            =   "frmManageConnection.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCheck 
      Caption         =   "C&heck Connection"
      Height          =   540
      Left            =   1575
      TabIndex        =   2
      Top             =   1380
      Width           =   1065
   End
   Begin VB.TextBox txtStatus 
      Height          =   1065
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2280
      Width           =   3900
   End
   Begin VB.ListBox lstConnections 
      Height          =   840
      Left            =   105
      TabIndex        =   0
      Top             =   315
      Width           =   3900
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "&Disconnect"
      Height          =   540
      Left            =   2730
      TabIndex        =   3
      Top             =   1380
      Width           =   1065
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   540
      Left            =   420
      TabIndex        =   1
      Top             =   1380
      Width           =   1065
   End
   Begin VB.Label lblHeadings 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      ForeColor       =   &H8000000D&
      Height          =   225
      Index           =   1
      Left            =   180
      TabIndex        =   6
      Top             =   2025
      Width           =   2115
   End
   Begin VB.Label lblHeadings 
      BackStyle       =   0  'Transparent
      Caption         =   "Connections:"
      ForeColor       =   &H8000000D&
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   105
      Width           =   1905
   End
End
Attribute VB_Name = "frmManageconnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Instance of the WinInet class.
'
Private WithEvents fInet As wininet
Attribute fInet.VB_VarHelpID = -1
Dim bConnected As Boolean

Public Property Get Connected() As Boolean
    Connected = bConnected
End Property
Public Sub Hangup()
Dim lngResult As Boolean
    lngResult = fInet.Hangup
End Sub

Public Sub Component(pINET As wininet)
    Set fInet = pINET
End Sub
Private Sub cmdCheck_Click()
    '
    ' Check the connection status.
    '
    If fInet.IsNetConnectOnline Then
        Call pAddToStatus("Internet connection active.")
    Else
        Call pAddToStatus("No active internet connection.")
    End If
End Sub
Private Sub cmdConnect_Click()
Dim lngResult As Long
    
    bConnected = False
    With lstConnections
    
    If .ListIndex = -1 Then
        Call MsgBox("Please select a DUN connection to use.", vbExclamation, "WinInet Demo")
    Else
        lngResult = fInet.StartDUN(Me.hwnd, .List(.ListIndex))
        
        If lngResult = 0 Then
            Call pAddToStatus("Connection to " & .List(.ListIndex) & " established.")
            bConnected = True
            Me.Hide
        Else
            If lngResult = -1 Then
                Call pAddToStatus("Already connected.")
            Else
                Call pAddToStatus("Error " & lngResult & " attempting to connect to " & .List(.ListIndex))
            End If
        End If
    End If
    
    End With
End Sub
Private Sub cmdDisconnect_Click()
Dim lngResult As Long
    
    lngResult = fInet.Hangup
    
    If lngResult = 0 Then
        Call pAddToStatus("Connection terminated.")
    Else
        If lngResult = -1 Then
            Call pAddToStatus("No existing connection.")
        Else
            Call pAddToStatus("Unable to terminate connection, error: " & lngResult)
        End If
    End If
End Sub
Private Sub fInet_ConnectionMade()
    '
    ' This event does not "monitor" the connection status.
    ' It is fired when a successfull connection is made via
    ' the .StartDUN method.
    '
    Call pAddToStatus("Connection made event fired.")
    
End Sub
Private Sub Form_Load()
Dim strDuns() As String
Dim lngIndex  As Long

    Set fInet = New wininet
    '
    Call fInet.ListDUNs(strDuns)
    lstConnections.Clear
    
    For lngIndex = 0 To UBound(strDuns)
        lstConnections.AddItem strDuns(lngIndex)
    Next
    
    txtStatus.Text = "Class initialized." & vbCrLf
End Sub
Private Sub pAddToStatus(ByVal strText As String)
    '
    ' Add the text to the textbox and
    ' scroll it into view if necessary.
    '
    With txtStatus
        .Text = .Text & strText & vbCrLf
        .SelStart = Len(.Text)
        .SelLength = 0
    End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set fInet = Nothing
End Sub
