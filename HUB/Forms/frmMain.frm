VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Papyrus HUB manager"
   ClientHeight    =   7260
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10380
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuRaw 
         Caption         =   "&Raw data"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuDataTransmission 
         Caption         =   "&Data transmission"
      End
      Begin VB.Menu mnuTESTAR 
         Caption         =   "TEST AR Designer"
      End
   End
   Begin VB.Menu mnuWind 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuDataTransmission_Click()
Dim frm As New frmTransmissionControl
    frm.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuRaw_Click()
Dim frm As New frmRawData
    
    frm.Show
End Sub

Private Sub mnuTESTAR_Click()
Dim f As New frmAR
    f.Show
End Sub
