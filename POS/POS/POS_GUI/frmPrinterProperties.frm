VERSION 5.00
Object = "{C9E1AFB0-1172-11D7-83AD-0050DA238ADA}#1.0#0"; "Coptr19.ocx"
Begin VB.Form frmPrinterProperties 
   Caption         =   "Printer properties"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFetch 
      Caption         =   "Fetch"
      Height          =   450
      Left            =   1920
      TabIndex        =   8
      Top             =   1290
      Width           =   1035
   End
   Begin VB.TextBox txtPrintername 
      Height          =   330
      Left            =   1335
      TabIndex        =   6
      Text            =   "txtProperty"
      Top             =   840
      Width           =   2220
   End
   Begin VB.CommandButton cmdPrinterPropertySet 
      Caption         =   "Set"
      Height          =   450
      Left            =   1935
      TabIndex        =   4
      Top             =   3015
      Width           =   1035
   End
   Begin VB.TextBox txtValue 
      Height          =   330
      Left            =   1365
      TabIndex        =   2
      Text            =   "txtProperty"
      Top             =   2370
      Width           =   2220
   End
   Begin VB.TextBox txtProperty 
      Height          =   330
      Left            =   1335
      TabIndex        =   0
      Text            =   "txtProperty"
      Top             =   1935
      Width           =   2220
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Printer name"
      Height          =   300
      Left            =   240
      TabIndex        =   7
      Top             =   885
      Width           =   1035
   End
   Begin COPTRLib.OPOSPOSPrinter OPOSPOSPrinter 
      Left            =   210
      Top             =   2985
      _Version        =   65536
      _ExtentX        =   1429
      _ExtentY        =   979
      _StockProps     =   0
   End
   Begin VB.Label Label 
      Caption         =   "Only works for Epson printers"
      Height          =   390
      Left            =   465
      TabIndex        =   5
      Top             =   165
      Width           =   3195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Value"
      Height          =   300
      Left            =   240
      TabIndex        =   3
      Top             =   2415
      Width           =   1035
   End
   Begin VB.Label lblProperty 
      Alignment       =   1  'Right Justify
      Caption         =   "Property"
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   1980
      Width           =   1035
   End
End
Attribute VB_Name = "frmPrinterProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OPOSPrinter As Object
Dim lngResult As Long

Private Sub cmdFetch_Click()
        Set OPOSPrinter = Me.OPOSPOSPrinter

70            With OPOSPrinter
                .Close
80                lngResult = .Open(txtPrintername)
90                If lngResult = 0 Then
100                   lngResult = .ClaimDevice(50)
110                   If lngResult = OPOS_SUCCESS Then
120                       .ClaimDevice 1000
130                       .DeviceEnabled = True
140                       .MapMode = PTR_MM_METRIC
                         txtProperty = .CharacterSet
150                       .RecLetterQuality = True
160                       .RecLineChars = 40
                          .ReleaseDevice
170                   Else
                        
180                       MsgBox "The till printer (" & txtPrintername & ") cannot be claimed by the application." & vbCrLf & "Result is " & CStr(lngResult) & ". This application will close."
200                       Exit Sub
210                   End If
                        
220               Else
230                   MsgBox "The till printer is not online. This application will close."
250                   Exit Sub
260               End If
270           End With

End Sub

Private Sub cmdPrinterPropertySet_Click()
        Set OPOSPrinter = Me.OPOSPOSPrinter
        
70            With OPOSPrinter
                .Close
80                lngResult = .Open(txtPrintername)
90                If lngResult = 0 Then
100                   lngResult = .ClaimDevice(50)
110                   If lngResult = OPOS_SUCCESS Then
120                       .ClaimDevice 1000
130                       .DeviceEnabled = True
140                       .MapMode = PTR_MM_METRIC
                         txtProperty = .CharacterSet
                         .CharacterSet = CLng(txtValue)
MsgBox .CharacterSet
150                       .RecLetterQuality = True
160                       .RecLineChars = 40
                          .ReleaseDevice
170                   Else
                        
180                       MsgBox "The till printer (" & txtPrintername & ") cannot be claimed by the application." & vbCrLf & "Result is " & CStr(lngResult) & ". This application will close."
200                       Exit Sub
210                   End If
                        
220               Else
230                   MsgBox "The till printer is not online. This application will close."
250                   Exit Sub
260               End If
270           End With


End Sub
