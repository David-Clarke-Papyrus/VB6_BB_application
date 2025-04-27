VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Google"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function LoadFromGoogle(strGOOXML As String, XGOO As XArrayDB, GooIndex As Integer) As Boolean
Dim resnext
Dim vwr As ujXML
Dim f As ujXML
Dim bElementEnd As Boolean
Dim sElementName As String
Dim MoreEntries As Boolean
Dim mXML As ujXML

    LoadFromGoogle = True
  '  GooIndex = 0
    Set mXML = New ujXML
    mXML.docLoadXML strGOOXML
    resnext = mXML.navLocate("entry")
    If Not resnext Then
        Set mXML = Nothing
        LoadFromGoogle = False
        Exit Function
    End If
    sElementName = ""
    MoreEntries = True

    Do
        Set vwr = mXML.docCreateViewer(True)
        GooIndex = GooIndex + 1
        XGOO.ReDim 1, GooIndex, 1, 15


        Do
            Select Case vwr.Element.nodeName
            Case "entry"
                resnext = vwr.navFirstChild
                
            Case "dc:date"
                Do
                    XGOO(GooIndex, 6) = XGOO(GooIndex, 6) & vwr.Element.text
                    resnext = vwr.navNext
                Loop While vwr.Element.nodeName = "dc:date" And resnext = True
            Case "dc:identifier"
                bElementEnd = False
                If GetISBN(vwr.Element.text) > "" Then
                    XGOO(GooIndex, 1) = GetISBN(vwr.Element.text)
                    bElementEnd = True
                End If
                resnext = vwr.navNext
                Do
                    If resnext Then
                        If GetISBN(vwr.Element.text) > "" Then
                            XGOO(GooIndex, 1) = GetISBN(vwr.Element.text)
                            bElementEnd = True
                        End If
                    Else
                        XGOO(GooIndex, 1) = ""
                    End If
                    resnext = vwr.navNext
                Loop While vwr.Element.nodeName = "dc:identifier" And bElementEnd = False
                XGOO(GooIndex, 1) = ConvertISBN10toEAN(XGOO(GooIndex, 1))
            Case "dc:title"
                Do
                    XGOO(GooIndex, 2) = XGOO(GooIndex, 2) & " " & vwr.Element.text
                    resnext = vwr.navNext
                Loop While vwr.Element.nodeName = "dc:title" And resnext = True
            Case "dc:creator"
                Do
                    XGOO(GooIndex, 3) = XGOO(GooIndex, 3) & " " & vwr.Element.text
                    resnext = vwr.navNext
                Loop While vwr.Element.nodeName = "dc:creator" And resnext = True
            Case "dc:publisher"
                Do
                    XGOO(GooIndex, 4) = XGOO(GooIndex, 4) & " " & vwr.Element.text
                    resnext = vwr.navNext
                Loop While vwr.Element.nodeName = "dc:publisher" And resnext = True
            Case "dc:format"
                Do
                    XGOO(GooIndex, 5) = XGOO(GooIndex, 5) & " " & vwr.Element.text
                    resnext = vwr.navNext
                Loop While vwr.Element.nodeName = "dc:format" And resnext = True
            Case Else
                resnext = vwr.navNext
            End Select
        Loop While resnext = True
        MoreEntries = mXML.navNext
    Loop While MoreEntries = True
    Set mXML = Nothing

End Function
Private Function GetISBN(p As String) As String
Dim a() As String
    GetISBN = ""
    a = Split(p, ":")
    If UBound(a) > 0 Then
        If a(0) = "ISBN" Then
            GetISBN = a(1)
        End If
    End If
End Function

