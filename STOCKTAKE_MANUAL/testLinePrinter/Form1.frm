VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1185
      Left            =   1005
      TabIndex        =   0
      Top             =   675
      Width           =   2280
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Since the Printer might be connected to LPT1 or LPT2 it
'had been changed to Printer.Port - KSM (21/11/2002)
'On Error GoTo To_NPrint
Dim strToPrint As String
Open Printer.Port For Output As #2
barcodeON = Chr(27) & Chr(16) & Chr(65) & Chr(4) & Chr(0) & Chr(2) & Chr(0) & Chr(1)
strToPrint = barcodeON & Chr(27) & Chr(16) & Chr(66) & Chr(13) & "9780285634114" & "                  " & Chr(27) & Chr(16) & Chr(66) & Chr(13) & "9780285634114"
'Line Input #1, strToPrint
If Right(strToPrint, 1) = " " Then strToPrint = Mid(strToPrint, 1, Len(strToPrint) - 1)
Print #2, strToPrint
Close #2
End Sub

