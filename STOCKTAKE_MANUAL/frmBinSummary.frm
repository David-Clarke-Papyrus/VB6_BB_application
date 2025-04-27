VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmBinSummary 
   Caption         =   "Form1"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCLose 
      BackColor       =   &H00F2E0D9&
      Caption         =   "Close"
      Height          =   435
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3660
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F2E0D9&
      Caption         =   "Print"
      Height          =   435
      Left            =   165
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3645
      Width           =   1275
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Height          =   3315
      Left            =   165
      OleObjectBlob   =   "frmBinSummary.frx":0000
      TabIndex        =   0
      Top             =   180
      Width           =   9255
   End
End
Attribute VB_Name = "frmBinSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub LoadArray()
10        On Error GoTo errHandler
      Dim oTF As z_TextFile
      Dim fs As New FileSystemObject
      Dim txtStream As Scripting.TextStream
      Dim fol
      Dim fils
      Dim f
      Dim i As Integer
      Dim k As Integer
      Dim s As String
      Dim x As New XArrayDB
      Dim ar() As String
      Dim strLine As String
      Dim lngQty As Long
      Dim bErrorCount As Boolean
      

20        Set fol = fs.GetFolder(oPC.SharedFolderRoot & IIf(Right(oPC.SharedFolderRoot, 1) = "\", "", "\") & "Stocktke")
30        Set fils = fol.Files
40        x.ReDim 1, fils.Count, 1, 3
50        k = 0
60        For Each f In fils
70            If f.Type <> "Text Document" Then
80                MsgBox "Skipping file: " & f.Name & " - not a valid count file type (.txt)"
90                GoTo EndOfLoop
100           End If
110           k = k + 1
120           i = 0
130          bErrorCount = False
             
140           Set txtStream = fs.OpenTextFile(f.Path)
150           Do While Not txtStream.AtEndOfStream
160               strLine = txtStream.ReadLine
170               If strLine = "" Then GoTo skip
180               ar = Split(strLine, ",")
190               If UBound(ar) = 0 Then
200                   lngQty = 1
210               Else
220                     If IsNumeric(Trim(ar(1))) Then
230                         lngQty = CLng(Trim(ar(1)))
240                     Else
250                         lngQty = 0
260                         bErrorCount = True
270                     End If
280               End If
290               i = i + lngQty
skip:
300           Loop
310           txtStream.Close
320           Set txtStream = Nothing
              
330           x(k, 1) = f.Name
340           x(k, 2) = IIf(bErrorCount = True, "Error in qyt captured(not numeric)", CStr(i))
350           x(k, 3) = f.Path
EndOfLoop:
360       Next
370       G.Array = x

380       Exit Sub
errHandler:
390       If ErrMustStop Then Debug.Assert False: Resume
400       ErrorIn "frmBinSummary.LoadArray"
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
10        G.PrintInfo.PageHeader = "Summary list of bins at: " & Format(Now(), "dd-mm-yyyy   HH:NN AMPM")
20        G.PrintInfo.PrintPreview
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBinSummary.cmdPrint_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
10        LoadArray
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBinSummary.Form_Load"
End Sub
