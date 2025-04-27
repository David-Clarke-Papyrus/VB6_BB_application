VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WordManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oWD As Word.Application
Attribute oWD.VB_VarHelpID = -1

Dim curDoc As Object 'Word.Document

Public Property Get WORDApplication() As Object
    On Error GoTo errHandler
    Set WORDApplication = oWD
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WordManager.WORDApplication"
End Property

Public Sub SetVisible(pVal As Boolean)
    On Error GoTo errHandler
    oWD.Visible = pVal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WordManager.SetVisible(pVal)", pVal
End Sub
Public Property Get WordIsVisible() As Boolean
    On Error GoTo errHandler
    WordIsVisible = oWD.Visible
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WordManager.WordIsVisible"
End Property
Public Function WordIsRunning() As Boolean
    On Error GoTo errHandler
    WordIsRunning = Not (oWD.Visible)
        WordIsRunning = True
    Exit Function
errHandler:
    If Err = 462 Then
        WordIsRunning = False
        Err.Clear
        Exit Function
    End If
    ErrorIn "WordManager.WordIsRunning"
End Function
Public Sub StartWORD(Optional pOK As Boolean)

'See this article at Microsoft
'Article ID  :   188546
'Last Review:    December 15, 2003
'Revision    :   1.0

Dim temp As Word.Application
    On Error GoTo errHandler
    pOK = True
    Set temp = CreateObject("Word.Application")
    Set oWD = Nothing
    Set oWD = CreateObject("Word.application")
    temp.Quit
    Set temp = Nothing
    Set curDoc = oWD.Documents.Add()
    oWD.Visible = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WordManager.StartWORD(pOK)", pOK
End Sub
Private Sub wd_Quit()
    On Error GoTo errHandler
'MsgBox "Here"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WordManager.wd_Quit"
End Sub

Public Function StartNewDocument(pTemplate As String) As Object 'Word.Document
    On Error GoTo errHandler
'MsgBox "In StartNewDocument pos 1"
    Set curDoc = oWD.Documents.Add(pTemplate)
'MsgBox "In StartNewDocument pos 2"
    Set StartNewDocument = curDoc
'MsgBox "In StartNewDocument pos 3"
    Exit Function
    
    
'ERRH:
'    MsgBox Error
'    Exit Function
'    Resume
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WordManager.StartNewDocument(pTemplate)", pTemplate
End Function
Public Sub StopWORD()
Dim oDoc As Word.Document

    On Error GoTo errHandler
    If Not oWD Is Nothing Then
        For Each oDoc In oWD.Documents
            oDoc.Close 0
        Next
        oWD.Quit
    End If
    Set oWD = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WordManager.StopWORD"
End Sub

Public Sub FillBM(pBM As String, pVal As String)
    On Error GoTo errHandler
Dim ran As Object 'Word.range
    If oWD.ActiveDocument.Bookmarks.Exists(pBM) = True Then
        Set ran = oWD.ActiveDocument.Bookmarks(pBM).range
        ran.InsertBefore pVal
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WordManager.FillBM(pBM,pVal)", Array(pBM, pVal)
End Sub

Public Sub ReplaceMarker(pMarker As String, pVal As String)
    On Error GoTo errHandler
Dim myrange As Object 'Word.range
Dim Result
    For Each myrange In oWD.ActiveDocument.StoryRanges
    
        Result = myrange.Find.Execute(FindText:=pMarker, replaceWith:=pVal, Wrap:=1)
        If Result = True Then
            Exit For
        End If

        Do While Not (myrange.NextStoryRange Is Nothing)
            Set myrange = myrange.NextStoryRange
            Result = myrange.Find.Execute(FindText:=pMarker, replaceWith:=pVal, Wrap:=1)
            If Result = True Then
                Exit Do
            End If
        Loop
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WordManager.ReplaceMarker(pMarker,pVal)", Array(pMarker, pVal)
End Sub

