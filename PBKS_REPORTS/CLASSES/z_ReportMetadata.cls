VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_ReportMetadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mMetadata As String
Dim xMLDoc As New ujXML
Dim SortedDoc As ujXML
Dim mSortString As String
Dim mfilterString As String
Dim Res As Boolean
Dim sColumnToRemove As String
Dim strSQL As String
Dim ItemToadd As String
Event LoadSortedItem(vak As String)
Public Property Let Report_view(val As String)
    Res = xMLDoc.navLocate("report_view")
    If Res Then
        xMLDoc.Element.text = val
    End If
End Property
Public Property Let RecordSource(val As String)
    Res = xMLDoc.navLocate("Control", "Name", "ADOData")
    'RecordSource
    If Res Then
        xMLDoc.attrS("RecordSource") = val
    End If
End Property
Public Property Let ConnectionString(val As String)
Dim a
    Res = xMLDoc.navLocate("Control", "Name", "ADOData")
    'ConnectionString
    If Res Then
        xMLDoc.attrS("ConnectionString") = val
    End If
End Property

Public Property Get Report_view() As String
    Res = xMLDoc.navLocate("report_view")
    If Res Then
        Report_view = xMLDoc.Element.text
    Else
        Report_view = ""
    End If
End Property
Public Property Let Report_name(val As String)
    Res = xMLDoc.navLocate("report_name")
    If Res Then
        xMLDoc.Element.text = val
    End If
End Property
Public Property Get Report_name() As String
    Res = xMLDoc.navLocate("report_name")
    If Res Then
        Report_name = xMLDoc.Element.text
    Else
        Report_name = ""
    End If
End Property

Public Property Let sortString(val As String)
    mSortString = Trim(val)
End Property
Public Property Get sortString() As String
    sortString = Trim(mSortString)
End Property
'Public Property Let filterString(val As String)
'    mfilterString = Trim(val)
'End Property
'Public Property Get filterString() As String
'    filterString = Trim(mfilterString)
'End Property

Public Sub ProcessElement(ByVal xObj As ujXML, ByVal NavAction As XENUM_ITER_NAV, ByRef Param As Variant, ByRef SkipChildren As Boolean)
Dim s As String
    If IsMissing(Param) Then
        Param = ""
    End If
    If Param = "SQL" Then
        If xObj.Element.nodeName = "COLUMN" Then
            If NavAction <> XNAV_TO_PARENT Then
                Res = xObj.navFirstChild
                sortString = sortString & IIf(sortString > "", ", ", "") & xObj.Element.text
                Res = xObj.navNext
                sortString = sortString & " " & xObj.Element.text & ""
                sortString = sortString & IIf(sortString > "", ", ", "") & s
                xObj.navUP
            End If
        End If
    ElseIf Param = "Delete" Then
        If xObj.Element.nodeName = "COLUMN" Then
           '  If NavAction <> XNAV_TO_PARENT Then
                Res = xObj.navFirstChild
                If Res Then
                    If xObj.Element.text = sColumnToRemove Then
                        xObj.navUP
                       xObj.bkmkSet "ColToDelete"
                    Else
                        xObj.navUP
                    End If
                End If
           '  End If
        End If
    Else
        If xObj.Element.nodeName = "COLUMN" Then
             If NavAction <> XNAV_TO_PARENT Then
                Res = xObj.navFirstChild
                ItemToadd = xObj.Element.text
                Res = xObj.navNext
                ItemToadd = ItemToadd & "(" & xObj.Element.text & ")"
                RaiseEvent LoadSortedItem(ItemToadd)
                xObj.navUP
             End If
        End If
    End If
    If Len(sortString) > 2 Then
        sortString = IIf(Right(sortString, 2) = ", ", Left(sortString, Len(sortString) - 2), sortString)
        sortString = IIf(Right(sortString, 1) = ",", Left(sortString, Len(sortString) - 1), sortString)
    End If
End Sub
Public Function GetSortString() As String
    sortString = ""
    SortedDoc.navTop
    SortedDoc.elForEachElem Me, "SQL"
    GetSortString = sortString
End Function
Public Function GetFilterString() As String
    GetFilterString = ""
    Res = xMLDoc.navLocate("filterBy")
    If Res Then
        GetFilterString = xMLDoc.Element.text
    Else
        GetFilterString = ""
    End If
    GetFilterString = " " & GetFilterString
End Function

Public Sub InitializeMetaDataXML()
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "REPORT_A"
        .chCreate "MessageType"
            .elText = "REPORT_METADATA"
        .elCreateSibling "MessageCreationDate"
            .elText = Format(Now(), "yyyymmddHHNN")
        .elCreateSibling "report_name"
        .elCreateSibling "report_view"
        .elCreateSibling "sortedBy"
        Set SortedDoc = .docCreateViewer(True)
        .elCreateSibling "filterBy"
        .elCreateSibling "selectColumns"
        .elCreateSibling "layout"
        .navUP
    End With
End Sub
Public Property Get Metadata_XML() As String
    Metadata_XML = xMLDoc.docXML
End Property

Public Sub ParseMetadata()
    SortedDoc.navTop
    SortedDoc.elForEachElem Me
End Sub

Public Sub LoadMetadataToXML(Metadata As String)
    On Error GoTo errHandler
    xMLDoc.docLoadXML Metadata
    xMLDoc.navLocate "sortedBy"
    Set SortedDoc = xMLDoc.docCreateViewer(True)
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ReportMetadata.LoadMetadataToXML(Metadata)", Metadata, , , "line number", Array(Erl())
End Sub
Public Property Get Layout_fromXML() As String
Dim s As String

    Res = xMLDoc.navLocate("ActiveReportsLayout")
'    s = xMLDoc.Element.Text
'    s = left(s, Len(s) - 3)
'    s = right(s, Len(s) - 9)
    If Res Then
        Layout_fromXML = xMLDoc.Element.xml
    End If
End Property

Public Sub AddSortedVolumn(fld As String, Seq As String)
    SortedDoc.chCreate "COLUMN", True, True
    SortedDoc.chCreate "FIELD", True, True
        SortedDoc.elText = fld
    SortedDoc.elCreateSibling "SEQUENCE", True
        SortedDoc.elText = Seq
    SortedDoc.navUP
    SortedDoc.navUP
End Sub
Public Sub RemoveSortedVolumn(fld As String)
    sColumnToRemove = fld
    SortedDoc.navTop
    SortedDoc.elForEachElem Me, "Delete"
    Res = SortedDoc.bkmkGoto("ColToDelete")
    SortedDoc.elRemove
 '   SortedDoc.elForEachElem Me
End Sub

Public Function GetSQL() As String
    sortString = ""
    strSQL = "SELECT TOP 100 PERCENT * FROM " & Report_view & Me.GetFilterString & " ORDER BY " & GetSortString
    GetSQL = strSQL
End Function
Public Function GetSQLShort() As String
Dim s As String
    s = "SELECT * FROM " & Report_view
    GetSQLShort = s
End Function
Public Function GetSQLOrderBy() As String
Dim s As String
    s = " ORDER BY " & GetSortString
    GetSQLOrderBy = s
End Function

Public Function AppendLayout(XMLString As String)
    Res = xMLDoc.navLocate("ActiveReportsLayout")
    
    If Res Then
        xMLDoc.elRemove
    End If
    xMLDoc.docAppendXML XMLString
End Function
Public Function AppendFilter(XMLString As String)
TOP:
    Res = xMLDoc.navLocate("filterBy")
    If Res = False Then
        Res = xMLDoc.navLocate("sortedBy")
        If Res = True Then
            xMLDoc.elCreateSibling "filterBy", True
            GoTo TOP
        End If
    End If
    xMLDoc.Element.text = XMLString
End Function
Public Function filterstring() As String
    Res = xMLDoc.navLocate("filterBy")
    
    filterstring = xMLDoc.Element.text
End Function

Public Sub clearSortingFilteringSelection()
    xMLDoc.navLocate "sortedby"
    xMLDoc.elRemoveSubElements
End Sub
