VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
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
   Begin VB.CommandButton cmdOutput 
      Caption         =   "Command3"
      Height          =   195
      Left            =   4395
      TabIndex        =   6
      Top             =   1530
      Width           =   210
   End
   Begin VB.CommandButton cmdStylesheet 
      Caption         =   "Command2"
      Height          =   195
      Left            =   4305
      TabIndex        =   5
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "Command1"
      Height          =   225
      Left            =   4380
      TabIndex        =   4
      Top             =   345
      Width           =   240
   End
   Begin VB.TextBox txtOutput 
      Height          =   375
      Left            =   270
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1350
      Width           =   4020
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   630
      Left            =   1530
      TabIndex        =   2
      Top             =   1890
      Width           =   1155
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   90
      Top             =   2730
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtStylesheet 
      Height          =   375
      Left            =   270
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   780
      Width           =   4020
   End
   Begin VB.TextBox txtSource 
      Height          =   330
      Left            =   270
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   255
      Width           =   4050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSource As String
Dim strStylesheet As String
Dim strOutput As String


Dim XSLDOC As New MSXML2.DOMDocument30
Dim opXMLDOC As New MSXML2.DOMDocument30
Dim XMLDOC  As New MSXML2.DOMDocument30


Private Sub cmdGo_Click()
    Set XMLDOC = Nothing
    Set XMLDOC = New MSXML2.DOMDocument30
    XMLDOC.async = False
    XMLDOC.validateOnParse = False
    XMLDOC.resolveExternals = False
    XMLDOC.Load strSource
    
    Set XSLDOC = Nothing
    Set XSLDOC = New MSXML2.DOMDocument30
    XSLDOC.async = False
    XSLDOC.validateOnParse = False
    XSLDOC.resolveExternals = False
    XSLDOC.Load strStylesheet
    
    Set opXMLDOC = New MSXML2.DOMDocument30
    opXMLDOC.async = False
    opXMLDOC.validateOnParse = False
    opXMLDOC.resolveExternals = False
  '  XMLDOC.docObject.transformNodeToObject XSLDOC, opXMLDOC
    XMLDOC.transformNodeToObject XSLDOC, opXMLDOC
    
    docWriteTostream "C:\PBKS\TEMP\edi.XML", opXMLDOC, "UNICODE"
    

End Sub

Private Sub cmdOutput_Click()
    CD1.ShowOpen
    strOutput = CD1.FileName
    txtOutput = strOutput
End Sub

Private Sub cmdSource_Click()
    CD1.ShowOpen
    strSource = CD1.FileName
    txtSource = strSource
End Sub

Private Sub cmdStylesheet_Click()
    CD1.ShowOpen
    strStylesheet = CD1.FileName
    txtStylesheet = strStylesheet
End Sub

Private Sub docWriteTostream(ByVal FilePath As String, obj As MSXML2.DOMDocument30, _
                Optional ByVal CharSet As String = "UNICODE")
    Dim s As Object
    Set s = CreateObject("ADODB.Stream")
    With s
        If CharSet <> "" Then .CharSet = CharSet
        .Open
        .WriteText obj.xml
        .SaveToFile FilePath, 2 'adSaveCreateOverWrite
        .Close
    End With
    Exit Sub
End Sub



