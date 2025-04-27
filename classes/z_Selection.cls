VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Selection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colNos As Collection
Event ValidSelection(OK As Boolean)
Public Sub Parse(pSelection As String)
Dim flgNoMoreParts As Boolean
Dim strPart As String
Dim START, ctr, lengthL, pos As Integer
    START = 1
    ctr = 0
    lengthL = Len(pSelection)
    
    flgNoMoreParts = Not Len(pSelection) > 0
    If flgNoMoreParts Then
        RaiseEvent ValidSelection(False)
        GoTo EXIT_Handler
    Else
        RaiseEvent ValidSelection(True)
    End If
    Do While Not flgNoMoreParts
        ctr = ctr + 1
        pos = InStr(START, pSelection, ",")
        If pos = 0 Then
          pos = lengthL + 1
          flgNoMoreParts = True
        End If
        If pos > START Then
            strPart = Mid(pSelection, START, (pos - START))
            START = pos + 1
        End If
        ParsePart strPart
    Loop
EXIT_Handler:
    Exit Sub

End Sub
Public Sub Load(objInserts As Collection)
    Set objInserts = colNos
End Sub
Private Sub ParsePart(pPart As String)
    If IsNumeric(pPart) Then
        If CInt(pPart) < 0 Then
            RaiseEvent ValidSelection(False)
        Else
            colNos.Add CInt(pPart)
        End If
    Else
        ParseRange (pPart)
    End If
End Sub
Private Sub ParseRange(pRange As String)
Dim iHyphenPos, iStartRange, iEndRange As Integer
Dim strStartRange, strEndRange As String

    iHyphenPos = InStr(1, pRange, "-")
    If iHyphenPos = 0 Then
        RaiseEvent ValidSelection(False)
        GoTo EXIT_Handler
    End If
    strStartRange = Left(pRange, iHyphenPos - 1)
    If IsNumeric(strStartRange) Then
        iStartRange = CInt(strStartRange)
    Else
        RaiseEvent ValidSelection(False)
        GoTo EXIT_Handler
    End If
    strEndRange = Right(pRange, Len(pRange) - iHyphenPos)
    If IsNumeric(strEndRange) Then
        iEndRange = CInt(strEndRange)
    Else
        RaiseEvent ValidSelection(False)
    End If
    For iStartRange = iStartRange To iEndRange
        colNos.Add iStartRange
    Next iStartRange
EXIT_Handler:
    Exit Sub
End Sub
Private Sub Class_Initialize()
    Set colNos = New Collection
End Sub
