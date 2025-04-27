VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ujXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
'                               ujXML
'                    from URFIN JUS (www.urfinjus.net)
'                  Copyright 2001, 2002. All rights reserved.
'version 2.1.2 10/05/2003
'=========================================================================================
'Public interface
'
'Enumerations
Option Explicit

Public Enum XENUM_ERR
    XERR_UJXML = ERRMAP_RESERVED_FIRST + 20
End Enum
    
Public Enum XENUM_ELEM_STATUS
    XES_EXISTING = 0
    XES_NEW
    XES_MODIFIED
    XES_DELETED
    XES_JOINED
End Enum
    
Public Enum XENUM_COPY_DIR
    XDIR_IN
    XDIR_OUT
    XDIR_XMLCLEAR
    XDIR_CTRLCLEAR
End Enum
    
Public Enum XENUM_OPER
    XOP_MIN
    XOP_MAX
    XOP_SUM
    XOP_CONCAT
    XOP_LIST
    XOP_COUNT
End Enum

Public Enum XENUM_ITER_NAV
    XNAV_START
    XNAV_TO_CHILD
    XNAV_TO_NEXT
    XNAV_TO_PARENT
End Enum

Public Enum XENUM_OPTIONS
    XO_FREETHREADED = &H1
    XO_USE_DICT = &H2
    XO_ISVIEWER = &H4
    XO_TRACK_CHANGES = &H8
    XO_NO_SYS_HANDLER = &H10
    XO_IIS_RETRIEVED = &H1000000
End Enum

Public Enum XENUM_SQLTYPE
    XSQL_NONE = 0
    XSQL_INSERT = &H1
    XSQL_UPDATE = &H2
    XSQL_DELETE = &H4
    XSQL_AUTO = &H10
End Enum

'Note: %1 is placeholder for CRLF
Const _
    SQL_SELECT = "SELECT * FROM %1", _
    SQL_INSERT = "INSERT INTO  %2 %1  (%3) %1  VALUES (%4)   ", _
    SQL_UPDATE = "UPDATE  %2 %1  SET %3 %1  WHERE %4  ", _
    SQL_DELETE = "DELETE FROM %2 %1  WHERE %3"

Const S_DICT = "__dict"
Const S_OPT = "__options"

Const S_STATUS = "__status"
Const S_BKMK = "__bkmks"
Const S_DATETIME = "__datetime"
Const S_FALSE = "False"
Const S_TRUE = "True"
Const S_DEFAULT_PROGID = "MSXML2.DOMDocument"
Const PROGID_FREETHREADEDDOC = "MSXML2.FreeThreadedDOMDocument"

Private mDoc As IXMLDOMDocument
Private mTopElem As IXMLDOMElement
Private mElem As IXMLDOMElement

Private mDocProgID As String
Private mOptions As Long
Private mCopyDir As XENUM_COPY_DIR
Private mDocModified As Boolean

'Private m as MSXML2.FreeThreadedDOMDocument

Private mDict As String, mBkmkList As String
Private mhistColl As Collection, mbkmkColl As Collection
Private mSavedEl As msxml2.IXMLDOMElement
Private miisReport As String

'Transactional stuff
Private mtranSavedXML As String



'Element =============================================================================================================
Public Property Get attrV(ByVal AttrName As String) As Variant
    On Error GoTo errHandler
    Dim V, vEmpty
    If docOptions(XO_USE_DICT) Then dictCheckName AttrName
    V = mElem.getAttribute(AttrName)
    attrV = IIf(IsNull(V), vEmpty, V)
    Exit Property
errHandler:
    ErrorIn "ujXML.attrV(AttrName)", AttrName
End Property

Public Property Let attrV(ByVal AttrName As String, ByVal AValue As Variant)
    On Error GoTo errHandler
    Dim OldValue
    If docOptions(XO_USE_DICT) Then dictCheckName AttrName
    With mElem
        OldValue = .getAttribute(AttrName)
        If IsNull(AValue) Or IsEmpty(AValue) Then
            .removeAttribute AttrName
            Else
            .setAttribute AttrName, AValue
        End If
        If docOptions(XO_TRACK_CHANGES) And Left$(AttrName, 1) <> "_" Then
            Dim NewV
            NewV = .getAttribute(AttrName)
            If uCompare(OldValue, NewV) Then elSetModified True
        End If
    End With
    Exit Property
errHandler:
    ErrorIn "ujXML.attrV(AttrName,AValue)", Array(AttrName, AValue)
End Property

Public Property Get attrV4DB(ByVal AttrName As String) As Variant
    On Error GoTo errHandler
    Dim V
    V = attrV(AttrName)
    If IsEmpty(V) Then
        attrV4DB = Null
        Else
        attrV4DB = V
    End If
    Exit Property
errHandler:
    ErrorIn "ujXML.attrV4DB(AttrName)", AttrName
End Property


Public Property Get attrS(ByVal AttrName As String) As String
    On Error GoTo errHandler
    attrS = uSafeStr(attrV(AttrName))
    Exit Property
errHandler:
    ErrorIn "ujXML.attrS(AttrName)", AttrName
End Property

Public Property Let attrS(ByVal AttrName As String, ByVal Value As String)
    On Error GoTo errHandler
    attrV(AttrName) = Value
    Exit Property
errHandler:
    ErrorIn "ujXML.attrS(AttrName,Value)", Array(AttrName, Value)
End Property

Public Property Get attrI(ByVal AttrName As String) As Long
    On Error GoTo errHandler
    attrI = uSafeLng(attrV(AttrName))
    Exit Property
errHandler:
    ErrorIn "ujXML.attrI(AttrName)", AttrName
End Property

Public Property Let attrI(ByVal AttrName As String, ByVal Value As Long)
    On Error GoTo errHandler
    attrV(AttrName) = Value
    Exit Property
errHandler:
    ErrorIn "ujXML.attrI(AttrName,Value)", Array(AttrName, Value)
End Property

Public Property Get attrB(ByVal AttrName As String) As Boolean
    On Error GoTo errHandler
    attrB = (attrV(AttrName) = S_TRUE)
    Exit Property
errHandler:
    ErrorIn "ujXML.attrB(AttrName)", AttrName
End Property

Public Property Let attrB(ByVal AttrName As String, ByVal Value As Boolean)
    On Error GoTo errHandler
    attrV(AttrName) = IIf(Value, S_TRUE, S_FALSE)
    Exit Property
errHandler:
    ErrorIn "ujXML.attrB(AttrName,Value)", Array(AttrName, Value)
End Property

Public Property Get attrF(ByVal AttrName As String) As Double
    On Error GoTo errHandler
    attrF = uSafeDbl(attrV(AttrName))
    Exit Property
errHandler:
    ErrorIn "ujXML.attrF(AttrName)", AttrName
End Property

Public Property Let attrF(ByVal AttrName As String, ByVal Value As Double)
    On Error GoTo errHandler
    attrV(AttrName) = Value
    Exit Property
errHandler:
    ErrorIn "ujXML.attrF(AttrName,Value)", Array(AttrName, Value)
End Property

Public Property Get attrDt(ByVal AttrName As String) As Date
    On Error GoTo errHandler
    attrDt = uSafeDate(attrV(AttrName))
    Exit Property
errHandler:
    ErrorIn "ujXML.attrDt(AttrName)", AttrName
End Property

Public Property Let attrDt(ByVal AttrName As String, ByVal Value As Date)
    On Error GoTo errHandler
    attrV(AttrName) = Value
    Exit Property
errHandler:
    ErrorIn "ujXML.attrDt(AttrName,Value)", Array(AttrName, Value)
End Property

Public Property Get attrEscHtm(ByVal AttrName As String) As String
    On Error GoTo errHandler
    attrEscHtm = uEscapeHTML(attrS(AttrName))
    Exit Property
errHandler:
    ErrorIn "ujXML.attrEscHtm(AttrName)", AttrName
End Property

Public Property Get attrEscXml(ByVal AttrName As String) As String
    On Error GoTo errHandler
    attrEscXml = uEscapeXML(attrS(AttrName))
    Exit Property
errHandler:
    ErrorIn "ujXML.attrEscXml(AttrName)", AttrName
End Property


Public Property Get attrQt(ByVal AttrName As String) As String
    On Error GoTo errHandler
    attrQt = uQuote(attrS(AttrName))
    Exit Property
errHandler:
    ErrorIn "ujXML.attrQt(AttrName)", AttrName
End Property

Public Property Get attrDqt(ByVal AttrName As String) As String
    On Error GoTo errHandler
    attrDqt = uDQuote(attrS(AttrName))
    Exit Property
errHandler:
    ErrorIn "ujXML.attrDqt(AttrName)", AttrName
End Property

Public Property Get attrSmartQt(ByVal AttrName As String) As String
    On Error GoTo errHandler
    attrSmartQt = uSmartQuote(attrS(AttrName))
    Exit Property
errHandler:
    ErrorIn "ujXML.attrSmartQt(AttrName)", AttrName
End Property

'-----------
Public Property Get attrIsNull(ByVal AttrName As String) As Boolean
    On Error GoTo errHandler
    attrIsNull = (mElem.getAttributeNode(AttrName) Is Nothing)
    Exit Property
errHandler:
    ErrorIn "ujXML.attrIsNull(AttrName)", AttrName
End Property

Public Property Get attrNotEmpty(ByVal AttrName As String) As Boolean
    On Error GoTo errHandler
    attrNotEmpty = (Trim$(attrS(AttrName)) <> "")
    Exit Property
errHandler:
    ErrorIn "ujXML.attrNotEmpty(AttrName)", AttrName
End Property

Public Property Get attrNZ(ByVal AttrName As String) As Boolean
    On Error GoTo errHandler
    attrNZ = (attrF(AttrName) <> 0)
    Exit Property
errHandler:
    ErrorIn "ujXML.attrNZ(AttrName)", AttrName
End Property

'Doesn't change mElemStatus!!
Public Sub attrSet(ByVal AttrName As String, ByVal Value)
    On Error GoTo errHandler
    mElem.setAttribute AttrName, Value
    Exit Sub
errHandler:
    ErrorIn "ujXML.attrSet(AttrName,Value)", Array(AttrName, Value)
End Sub

Public Sub elSetModified(ByVal AModified As Boolean)
    On Error GoTo errHandler
    Select Case elStatus
        Case XES_NEW: 'Nothing to do
        Case XES_EXISTING: If AModified Then mElem.setAttribute S_STATUS, "M"
        Case XES_MODIFIED: If Not AModified Then mElem.removeAttribute S_STATUS
        Case XES_DELETED: 'nothing to do
        Case XES_JOINED:  'nothing to do
    End Select
    If AModified Then mDocModified = True
    Exit Sub
errHandler:
    ErrorIn "ujXML.elSetModified(Modified)", AModified
End Sub

Public Sub attrSetIfEmpty(ByVal AttrName As String, ByVal Value)
    On Error GoTo errHandler
    If Not attrNotEmpty(AttrName) Then attrV(AttrName) = Value
    Exit Sub
errHandler:
    ErrorIn "ujXML.attrSetIfEmpty(AttrName,Value)", Array(AttrName, Value)
End Sub

Public Sub attrSetAsLkpFrom(ByVal AttrName As String, IndexAttr As String, _
        ByVal FromList As String, Optional ByVal Delim As String = ",")
    On Error GoTo errHandler
    attrS(AttrName) = uSubstrByIndex(FromList, attrI(IndexAttr), Delim)
    Exit Sub
errHandler:
    ErrorIn "ujXML.attrSetAsLkpFrom(AttrName,IndexAttr,FromList,Delim)", Array(AttrName, IndexAttr, _
         FromList, Delim)
End Sub

Public Sub attrSetAsIndexFrom(ByVal AttrName As String, ByVal ValueAttr As String, _
        ByVal FromList As String, Optional ByVal Delim As String = ",")
    On Error GoTo errHandler
    attrI(AttrName) = uIndexOf(attrS(ValueAttr), FromList, Delim)
    Exit Sub
errHandler:
    ErrorIn "ujXML.attrSetAsIndexFrom(AttrName,ValueAttr,FromList,Delim)", Array(AttrName, ValueAttr, _
         FromList, Delim)
End Sub

Public Function elGetAttrList() As String
    On Error GoTo errHandler
    Dim Ndl As IXMLDOMNodeList, i As Long, sList As String
    Set Ndl = mElem.selectNodes("@*")
    For i = 0 To Ndl.Length - 1
        sList = uCombine(sList, ",", Ndl.Item(i).nodeName)
    Next i
    elGetAttrList = sList
    Exit Function
errHandler:
    ErrorIn "ujXML.elGetAttrList"
End Function


Public Sub attrCopy(ByVal xFrom As ujXML, _
                Optional ByVal Specs As String = "", _
                Optional ByVal AddToDict As Boolean = False)
    On Error GoTo errHandler
    Dim arrD, arrS, arrQ
    If Specs = "" Then Specs = xFrom.elGetAttrList
    uUnpackEQList Specs, arrD, arrS, arrQ
    If AddToDict Then dictAdd Join(arrD, ",")
    Dim i As Long
    For i = 0 To UBound(arrD)
        If arrQ(i) Then
            attrV(arrD(i)) = arrS(i)
            Else
            attrV(arrD(i)) = xFrom.attrV(arrS(i))
        End If
    Next i
    Exit Sub
errHandler:
    ErrorIn "ujXML.attrCopy(xFrom,Specs,AddToDict)", Array(xFrom, Specs, AddToDict)
End Sub

Public Function attrConstruct(ByVal AttrName As String, ByVal Template As String, _
        ByVal PrmAttrList As String) As String
    On Error GoTo errHandler
    attrV(AttrName) = uInsInto(Template, attrGetArr(PrmAttrList))
    Exit Function
errHandler:
    ErrorIn "ujXML.attrConstruct(AttrName,Template,PrmAttrList)", Array(AttrName, Template, _
         PrmAttrList)
End Function

Public Function attrGetArr(ByVal AttrList As String, Optional ByVal InsValues) As Variant
    On Error GoTo errHandler
    Dim arrA() As String, arrV(), Ub As Long, i As Long, InsIndex As Long, V, AName As String
    arrA = Split(AttrList, ",")
    If Not IsMissing(InsValues) Then
        If Not IsArray(InsValues) Then InsValues = Array(InsValues)
    End If
    ReDim arrV(LBound(arrA) To UBound(arrA))
    For i = LBound(arrV) To UBound(arrV)
        AName = arrA(i)
        If AName = "" Then
            Check Not IsMissing(InsValues), XERR_UJXML, "Skipped attribute found but InsValues array is not provided."
            V = InsValues(InsIndex)
            InsIndex = InsIndex + 1
            Else
            V = attrV(AName)
        End If
        arrV(i) = V
    Next i
    attrGetArr = arrV
    Exit Function
errHandler:
    ErrorIn "ujXML.attrGetArr(AttrList,InsValues)", Array(AttrList, InsValues)
End Function

Public Sub attrSetArr(ByVal AttrList As String, ByVal Values)
    On Error GoTo errHandler
    Dim i As Long, arrA() As String
    If AttrList = "" Then Exit Sub
    arrA = Split(AttrList, ",")
    Check UBound(arrA) = UBound(Values), XERR_UJXML, "List counts mismatch: Count(AttrList)=%1, Count(Values)=%2", _
        Array(UBound(arrA) + 1, UBound(Values) + 1)
    For i = 0 To UBound(arrA)
        attrV(arrA(i)) = Values(i)
    Next i
    Exit Sub
errHandler:
    ErrorIn "ujXML.attrSetArr(AttrList,Values)", Array(AttrList, Values)
End Sub

Public Sub attrRemove(ByVal AttrName As String)
    On Error GoTo errHandler
    mElem.removeAttribute AttrName
    Exit Sub
errHandler:
    ErrorIn "ujXML.attrRemove(AttrName)", AttrName
End Sub


Public Property Get Element() As IXMLDOMElement
    Set Element = mElem
End Property

Public Property Set Element(ByVal AnElem As IXMLDOMElement)
  Set mElem = AnElem
End Property

Public Sub elClear()
    On Error GoTo errHandler
    If (mElem Is docRootElem) Then
        docInit elName
        Else
        elCreateSibling elName, True
        elRemove
    End If
    Exit Sub
errHandler:
    ErrorIn "ujXML.elClear"
End Sub

Public Function elRemove() As Boolean
    On Error GoTo errHandler
    Dim Parent As IXMLDOMElement, NextEl As IXMLDOMElement
    Check Not mElem Is docTopElem, ERR_GENERAL, "Cannot remove document root mElement"
    Set Parent = mElem.parentNode
    Set NextEl = mElem.previousSibling
    If NextEl Is Nothing Then Set NextEl = mElem.nextSibling
    If NextEl Is Nothing Then
        Set NextEl = Parent
        elRemove = False
        Else
        elRemove = True
    End If
    Parent.removeChild mElem
    mDocModified = True
    elAfterDelete mElem
    Set mElem = NextEl
    Exit Function
errHandler:
    ErrorIn "ujXML.elRemove"
End Function

Public Property Get elText() As String
    On Error GoTo errHandler
    elText = mElem.Text
    Exit Property
errHandler:
    ErrorIn "ujXML.elText"
End Property

Public Property Let elText(ByVal Value As String)
    On Error GoTo errHandler
    If (mElem.Text <> Value) And docOptions(XO_TRACK_CHANGES) Then
        elSetModified True
    End If
    mElem.Text = Value
    Exit Property
errHandler:
    ErrorIn "ujXML.elText(Value)", Value
End Property

Public Sub elRemoveSubElements()
    On Error GoTo errHandler
    Do While navFirstChild
        If elRemove Then navUP
    Loop
    Exit Sub
errHandler:
    ErrorIn "ujXML.elRemoveSubElements"
End Sub

'Index/Position --------------------------------------------
Public Function elIsFirst() As Boolean
    On Error GoTo errHandler
    elIsFirst = (mElem.previousSibling Is Nothing)
    Exit Function
errHandler:
    ErrorIn "ujXML.elIsFirst"
End Function

Public Function elIsLast() As Boolean
    On Error GoTo errHandler
    elIsLast = (mElem.nextSibling Is Nothing)
    Exit Function
errHandler:
    ErrorIn "ujXML.elIsLast"
End Function

'Element neighbours =====================================================
Public Property Get elNextSibling() As IXMLDOMElement
    On Error GoTo errHandler
    Set elNextSibling = mElem.nextSibling
    Exit Property
errHandler:
    ErrorIn "ujXML.elNextSibling"
End Property

Public Property Get elPrevSibling() As IXMLDOMElement
    On Error GoTo errHandler
    Set elPrevSibling = mElem.previousSibling
    Exit Property
errHandler:
    ErrorIn "ujXML.elPrevSibling"
End Property

Public Property Get elFirstSibling() As IXMLDOMElement
    On Error GoTo errHandler
    If elParent Is Nothing Then
        Set elFirstSibling = mElem
        Else
        Set elFirstSibling = elParent.firstChild
    End If
    Exit Property
errHandler:
    ErrorIn "ujXML.elFirstSibling"
End Property

Public Property Get elLastSibling() As IXMLDOMElement
    On Error GoTo errHandler
    If elParent Is Nothing Then
        Set elLastSibling = mElem
        Else
        Set elLastSibling = elParent.lastChild
    End If
    Exit Property
errHandler:
    ErrorIn "ujXML.elLastSibling"
End Property

Public Property Get elParent() As IXMLDOMElement
    Dim Nd As IXMLDOMNode
    On Error GoTo errHandler
    Set Nd = mElem.parentNode
    If TypeOf Nd Is IXMLDOMElement Then Set elParent = Nd
    Exit Property
errHandler:
    ErrorIn "ujXML.elParent"
End Property

Public Property Get elFirstChild() As IXMLDOMElement
    On Error GoTo errHandler
    Set elFirstChild = mElem.firstChild
    Exit Property
errHandler:
    ErrorIn "ujXML.elFirstChild"
End Property

Public Property Get elLastChild() As IXMLDOMElement
    On Error GoTo errHandler
    Set elLastChild = mElem.lastChild
    Exit Property
errHandler:
    ErrorIn "ujXML.elLastChild"
End Property

Public Function elLocateXsl(ByVal Path As String, Optional ByVal Values, _
        Optional ByVal RaiseErr As Boolean = True) As Boolean
    On Error GoTo errHandler
    Dim tEl As IXMLDOMElement
    If Not IsMissing(Values) Then Path = uInsInto(Path, Values)
    Set tEl = mElem.selectSingleNode(Path)
    If tEl Is Nothing Then
        If RaiseErr Then
            Check False, XERR_UJXML, "Navigation from mElement [%1] using path [%2] failed.", _
                    Array(elPath, Path)
        End If
        Else
        Set mElem = tEl
        elLocateXsl = True
    End If
    Exit Function
errHandler:
    ErrorIn "ujXML.elLocateXsl(Path,Values,RaiseErr)", Array(Path, Values, RaiseErr)
End Function

Public Function elGetChildByIndex(ByVal Index As Long) As IXMLDOMElement
    On Error GoTo errHandler
    elGetChildByIndex = mElem.childNodes(Index)
    Exit Function
errHandler:
    ErrorIn "ujXML.elGetChildByIndex(Index)", Index
End Function

Public Function elGetSiblingByIndex(ByVal Index As Long) As IXMLDOMElement
    On Error GoTo errHandler
    If elParent Is Nothing Then Exit Function
    elGetSiblingByIndex = elParent.childNodes(Index)
    Exit Function
errHandler:
    ErrorIn "ujXML.elGetSiblingByIndex(Index)", Index
End Function

'Moving mElement =============================================================
Public Sub elMoveTo(Optional ByVal BeforeEl As IXMLDOMElement, _
        Optional ByVal AfterEl As IXMLDOMElement, _
        Optional ByVal NewPrntEl As IXMLDOMElement)
    On Error GoTo errHandler
    Check Not ((BeforeEl Is docRootElem) Or (AfterEl Is docRootElem)), XERR_UJXML, _
        "Before/After reference mElement may not be document root."
    If AfterEl Is Nothing Then
        If BeforeEl Is Nothing Then
            Check Not (NewPrntEl Is Nothing), XERR_UJXML, "At least one reference mElement must be provided."
            NewPrntEl.appendChild mElem
            Else
            BeforeEl.parentNode.insertBefore mElem, BeforeEl
        End If
      Else
        If AfterEl.nextSibling Is Nothing Then
            AfterEl.parentNode.appendChild mElem
            Else
            AfterEl.parentNode.insertBefore mElem, AfterEl.nextSibling
        End If
    End If
    Exit Sub
errHandler:
    ErrorIn "ujXML.elMoveTo(BeforeEl,AfterEl,NewPrntEl)", Array(BeforeEl, AfterEl, NewPrntEl)
End Sub

Public Function elMoveNext() As Boolean
    On Error GoTo errHandler
    If Not elNextSibling Is Nothing Then
        elMoveTo , elNextSibling
        elMoveNext = True
    End If
    Exit Function
errHandler:
    ErrorIn "ujXML.elMoveNext"
End Function

Public Function elMovePrev() As Boolean
    On Error GoTo errHandler
    If Not elPrevSibling Is Nothing Then
        elMoveTo elPrevSibling
        elMovePrev = True
    End If
    Exit Function
errHandler:
    ErrorIn "ujXML.elMovePrev"
End Function

Public Sub elCreateSibling(Optional ByVal AnElemName As String, Optional InsAfter As Boolean = True)
    On Error GoTo errHandler
    Check Not mElem Is mTopElem, XERR_UJXML, "Cannot create sibling for root Element."
    If AnElemName = "" Then AnElemName = elName
    If InsAfter Then
        docCreateElem AnElemName, , mElem
        Else
        docCreateElem AnElemName, mElem
    End If
    Exit Sub
errHandler:
    ErrorIn "ujXML.elCreateSibling(AnElemName,InsAfter)", Array(AnElemName, InsAfter)
End Sub

'Returns zero-based mElement's index in siblings collection
Public Function elIndex() As Long
    On Error GoTo errHandler
    Dim tDoc As msxml2.DOMDocument, tEl As msxml2.IXMLDOMElement, Ind As Long
    Set tEl = mElem
    While Not tEl.previousSibling Is Nothing
        Set tEl = tEl.previousSibling
        Ind = Ind + 1
    Wend
    elIndex = Ind
    Exit Function
errHandler:
    ErrorIn "ujXML.elIndex"
End Function

Public Property Get elName() As String
    On Error GoTo errHandler
    elName = mElem.nodeName
    Exit Property
errHandler:
    ErrorIn "ujXML.elName"
End Property

Public Property Get elIsTop() As Boolean
    On Error GoTo errHandler
    elIsTop = (mElem Is mTopElem)
    Exit Property
errHandler:
    ErrorIn "ujXML.elIsTop"
End Property

Public Property Get elIsRoot() As Boolean
    On Error GoTo errHandler
    elIsRoot = (mElem Is docRootElem)
    Exit Property
errHandler:
    ErrorIn "ujXML.elIsRoot"
End Property

Public Property Get elPath() As String
    On Error GoTo errHandler
    Dim El As IXMLDOMElement
    Set El = mElem
    elPath = PathPriv
    Set mElem = El
    Exit Property
errHandler:
    ErrorIn "ujXML.elPath"
End Property

Private Function PathPriv() As String
    On Error GoTo errHandler
    Dim Ind As Long, ParPath As String, ElRef As String
    If elIsRoot Then
        PathPriv = "/" & elName
        Else
        Ind = elIndex
        ElRef = elName
        If Ind > 0 Then ElRef = ElRef & "[" & Ind & "]"
        navUP
        ParPath = PathPriv()
        PathPriv = ParPath & "/" & ElRef
    End If
    Exit Function
errHandler:
    ErrorIn "ujXML.PathPriv"
End Function

Public Property Get elXML() As String
    On Error GoTo errHandler
    elXML = mElem.xml
    Exit Property
errHandler:
    ErrorIn "ujXML.elXML"
End Property

Public Property Let elXML(ByVal xmlElem As String)
    On Error GoTo errHandler
    Dim xTmp As ujXML, ParNd As IXMLDOMNode, NewEl As IXMLDOMElement
    If mElem Is docRootElem Then
        docLoadXML xmlElem
        Else
        Set ParNd = mElem.parentNode
        Set xTmp = New ujXML
        xTmp.docProgID = docProgID
        xTmp.docLoadXML xmlElem
        Set NewEl = xTmp.docRootElem
        ParNd.replaceChild NewEl, mElem
    End If
    If mTopElem Is mElem Then Set mTopElem = NewEl
    elAfterDelete mElem
    Set mElem = NewEl
    Exit Property
errHandler:
    ErrorIn "ujXML.elXML(xmlElem)", xmlElem
End Property

Private Sub elAfterDelete(ByVal El As IXMLDOMElement)
    On Error GoTo errHandler
    bkmkRemoveEl El
    histRemoveEl El
    Exit Sub
errHandler:
    ErrorIn "ujXML.elAfterDelete(El)", El
End Sub

Public Property Get elStatus() As XENUM_ELEM_STATUS
    On Error GoTo errHandler
    Dim Ch As String
    Ch = attrS(S_STATUS)
    Select Case UCase(Ch)
        Case "", "E": elStatus = XES_EXISTING
        Case "N": elStatus = XES_NEW
        Case "M": elStatus = XES_MODIFIED
        Case "D": elStatus = XES_DELETED
        Case "J": elStatus = XES_JOINED
    End Select
    Exit Property
errHandler:
    ErrorIn "ujXML.elStatus"
End Property

Public Property Let elStatus(ByVal AValue As XENUM_ELEM_STATUS)
    On Error GoTo errHandler
    Select Case AValue
        Case XES_EXISTING: mElem.removeAttribute S_STATUS
        Case XES_NEW: attrS(S_STATUS) = "N"
        Case XES_MODIFIED: attrS(S_STATUS) = "M"
        Case XES_DELETED: attrS(S_STATUS) = "D"
        Case XES_JOINED: attrS(S_STATUS) = "J"
    End Select
    Exit Property
errHandler:
    ErrorIn "ujXML.elStatus(AValue)", AValue
End Property

Public Property Get elVisible() As Boolean
    On Error GoTo errHandler
    Select Case elStatus
        Case XES_EXISTING, XES_MODIFIED, XES_NEW:  elVisible = True
    End Select
    Exit Property
errHandler:
    ErrorIn "ujXML.elVisible"
End Property

Public Function elCheckValues(ByVal AttrNames, ByVal AttrValues) As Boolean
    On Error GoTo errHandler
    Dim arrNames() As String, i As Long
    If Not IsArray(AttrValues) Then AttrValues = Array(AttrValues)
    If IsArray(AttrNames) Then
        arrNames = AttrNames
        Else
        If AttrNames = "" Then elCheckValues = True: Exit Function
        arrNames = Split(AttrNames, ",")
        Check UBound(arrNames) = UBound(AttrValues), XERR_UJXML, _
            "Attribute names/values count mismatch: Attributes=%1, Values = %2", _
            Array(UBound(arrNames) + 1, UBound(AttrValues) + 1)
    End If
    For i = LBound(arrNames) To UBound(arrNames)
        If attrS(arrNames(i)) <> AttrValues(i) Then Exit Function  'returning false
    Next i
    elCheckValues = True
    Exit Function
errHandler:
    ErrorIn "ujXML.elCheckValues(AttrNames,AttrValues)", Array(AttrNames, AttrValues)
End Function

Public Function elJoin(ByVal xWithList As ujXML, ByVal OnAttr As String, ByVal CopyList As String, _
        Optional ByVal TblBkmk As String)
    On Error GoTo errHandler
    Dim SrchAttr As String, SrcAttr As String, arrT() As String, KeyV
    arrT = Split(OnAttr, "=")
    SrcAttr = arrT(0)
    If UBound(arrT) = 0 Then SrchAttr = SrcAttr Else SrchAttr = arrT(1)
    KeyV = attrS(SrcAttr)
    If Not xWithList.tblLocateElem(TblBkmk, SrchAttr, KeyV) Then Exit Function
    attrCopy xWithList, CopyList
    Exit Function
errHandler:
    ErrorIn "ujXML.elJoin(xWithList,OnAttr,CopyList,TblBkmk)", Array(xWithList, OnAttr, CopyList, _
         TblBkmk)
End Function

Public Sub elForEachElem(ByVal Iterator As Object, Optional Param)
    On Error GoTo errHandler
    Dim SkipChildren As Boolean
    Iterator.ProcessElement Me, XNAV_START, Param, SkipChildren
    If chCount > 0 And Not SkipChildren Then elForEachChild Iterator, Param
    Exit Sub
errHandler:
    ErrorIn "ujXML.elForEachElem(Iterator,Param)", Array(Iterator, Param)
End Sub

Private Sub elForEachChild(ByVal Iterator As Object, Optional Param)
    On Error GoTo errHandler
    Dim SkipChildren As Boolean, nav As XENUM_ITER_NAV
    If Not navFirstChild Then Exit Sub
    nav = XNAV_TO_CHILD
    Do
        SkipChildren = False
        Iterator.ProcessElement Me, nav, Param, SkipChildren
        If chCount > 0 And Not SkipChildren Then elForEachChild Iterator, Param
        nav = XNAV_TO_NEXT
    Loop While navNext
    navUP
    Iterator.ProcessElement Me, XNAV_TO_PARENT, Param, SkipChildren
    Exit Sub
errHandler:
    ErrorIn "ujXML.elForEachChild(Iterator,Param)", Array(Iterator, Param)
End Sub

'Moving/Navigating =============================================================================================
Public Function navNext() As Boolean
    On Error GoTo errHandler
    Dim El As msxml2.IXMLDOMElement, sName As String
    Set El = mElem.nextSibling
    If Not El Is Nothing Then
        Set mElem = El
        navNext = True
    End If
    Exit Function
errHandler:
    ErrorIn "ujXML.navNext"
End Function

Public Function navPrevious() As Boolean
    On Error GoTo errHandler
    Dim El As msxml2.IXMLDOMElement, sName As String
    Set El = mElem.previousSibling
    If Not El Is Nothing Then
        Set mElem = El
        navPrevious = True
    End If
    Exit Function
errHandler:
    ErrorIn "ujXML.navPrevious"
End Function

Public Sub navFirst()
    On Error GoTo errHandler
    Set mElem = mElem.parentNode.firstChild
    Exit Sub
errHandler:
    ErrorIn "ujXML.navFirst"
End Sub

Public Sub navLast()
    On Error GoTo errHandler
    Set mElem = mElem.parentNode.lastChild
    Exit Sub
errHandler:
    ErrorIn "ujXML.navLast"
End Sub

Public Function navFirstChild() As Boolean
    On Error GoTo errHandler
    Dim El As msxml2.IXMLDOMElement
    Set El = mElem.firstChild
    If Not El Is Nothing Then
        Set mElem = El
        navFirstChild = True
    End If
    Exit Function
errHandler:
    ErrorIn "ujXML.navFirstChild"
End Function

Public Function navLastChild() As Boolean
    On Error GoTo errHandler
    Dim El As msxml2.IXMLDOMElement
    Set El = mElem.lastChild
    If Not El Is Nothing Then
        Set mElem = El
        navLastChild = True
    End If
    Exit Function
errHandler:
    ErrorIn "ujXML.navLastChild"
End Function

Public Sub navChildByIndex(ByVal Index As Long)
    On Error GoTo errHandler
    Dim El As msxml2.IXMLDOMElement
    Set El = mElem.childNodes(Index)
    Check Not El Is Nothing, XERR_UJXML, "Child not found for index = " & Index
    Set mElem = El
    Exit Sub
errHandler:
    ErrorIn "ujXML.navChildByIndex"
End Sub

Public Sub navUP()
    On Error GoTo errHandler
    Dim Nd As msxml2.IXMLDOMNode
    Set Nd = mElem.parentNode
    Check (TypeOf Nd Is IXMLDOMElement), XERR_UJXML, "Cannot move up from document root"
    Set mElem = Nd
    Exit Sub
errHandler:
    ErrorIn "ujXML.navUP"
End Sub

Public Sub navTop()
    On Error GoTo errHandler
    Set mElem = docTopElem
    Exit Sub
errHandler:
    ErrorIn "ujXML.navTop"
End Sub

Public Function navLocate(ByVal AnElemName As String, Optional ByVal AttrName As String, Optional ByVal Value As String) As Boolean
    On Error GoTo errHandler
    Dim El As msxml2.IXMLDOMElement, Path As String
    Path = "//" & AnElemName
    If AttrName <> "" Then Path = Path & "[@" & AttrName & "=" & uSmartQuote(Value) & "]"
    Set El = mElem.selectSingleNode(Path)
    If Not El Is Nothing Then
        Set mElem = El
        navLocate = True
    End If
    Exit Function
errHandler:
    ErrorIn "ujXML.navLocate(ElemName,AttrName,Value)", Array(AnElemName, AttrName, Value)
End Function

Public Function navTo(ByVal Path As String, Optional ByVal Index As Long = -1) As Boolean
    On Error GoTo errHandler
    Dim El As msxml2.IXMLDOMElement
    If Index <> -1 Then Path = Path & "[" & Index & "]"
    Set El = mElem.selectSingleNode(Path)
    If Not El Is Nothing Then
        Set mElem = El
        navTo = True
    End If
    Exit Function
errHandler:
    ErrorIn "ujXML.navTo(Path,Index)", Array(Path, Index)
End Function

'Child ===========================================================================================
Public Sub chCreate(ByVal AnElemName As String, Optional ByVal NavigateToIt As Boolean = True, _
        Optional ByVal AsLast As Boolean = True, Optional ByVal MarkAsNew As Boolean = True)
    On Error GoTo errHandler
    Dim El As msxml2.IXMLDOMElement
    Set El = mDoc.createElement(AnElemName)
    With mElem
        If Not .hasChildNodes Then AsLast = True
        If AsLast Then
            .appendChild El
        Else
            .insertBefore El, .firstChild
        End If
        If docOptions(XO_TRACK_CHANGES) And MarkAsNew Then El.setAttribute S_STATUS, "N"
    End With
    If NavigateToIt Then Set mElem = El
    Exit Sub
errHandler:
    ErrorIn "ujXML.chCreate(AnElemName,NavigateToIt,AsLast,MarkAsNew)", Array(AnElemName, NavigateToIt, _
         AsLast, MarkAsNew)
End Sub

Public Property Get chCount() As Long
    On Error GoTo errHandler
    chCount = mElem.childNodes.Length
    Exit Property
errHandler:
    ErrorIn "ujXML.chCount"
End Property

Public Function chLocate(Optional ChElemName As String, Optional ByVal AttrNames As String, _
        Optional ByVal AttrValues) As Boolean
    On Error GoTo errHandler
    Dim Found As Boolean, NameMismatch As Boolean, arrNames() As String
    If AttrNames <> "" Then
        arrNames = Split(AttrNames, ",")
        If Not IsArray(AttrValues) Then AttrValues = Array(AttrValues)
        Check UBound(arrNames) = UBound(AttrValues), XERR_UJXML, "Invalid parameters: " & _
            "Number of attributes (%1) doesn't match number of Values (%2)", _
        Array(UBound(arrNames) + 1, UBound(AttrValues) + 1)
    End If
    If Not navFirstChild Then Exit Function
    Do
        NameMismatch = (ChElemName <> "") And (elName <> ChElemName)
        If (Not NameMismatch) Then
            If (AttrNames = "") Then
                chLocate = True
                Exit Function
            Else
                If elCheckValues(arrNames, AttrValues) Then chLocate = True: Exit Function
            End If '(AttrNames = ....
        End If '(Not NameMismatch ....
    Loop While navNext
    navUP
    Exit Function
errHandler:
    ErrorIn "ujXML.chLocate(ChElemName,AttrNames,AttrValues)", Array(ChElemName, AttrNames, _
         AttrValues)
End Function

Public Function chLocateByIndex(ByVal Index As Long, Optional ByVal VisOnly As Boolean)
    On Error GoTo errHandler
    Dim Nd As IXMLDOMNode, Path As String
    Set Nd = mElem.selectSingleNode("*[" & Index & "]")
    Exit Function
errHandler:
    ErrorIn "ujXML.chLocateByIndex(Index,VisOnly)", Array(Index, VisOnly)
End Function

Public Property Get chText(ByVal ChildName As String) As String
    On Error GoTo errHandler
    If chLocate(ChildName) Then
        chText = elText
        navUP
    End If
    Exit Property
errHandler:
    ErrorIn "ujXML.chText(ChildName)", ChildName
End Property

Public Property Let chText(ByVal ChildName As String, ByVal Value As String)
    On Error GoTo errHandler
    If chLocate(ChildName) Then
        elText = Value
        Else
        chCreate ChildName
        elText = Value
    End If
    navUP
    Exit Property
errHandler:
    ErrorIn "ujXML.chText(ChildName,Value)", Array(ChildName, Value)
End Property

Public Sub chAConstruct(ByVal AttrName As String, ByVal Template As String, ByVal PrmAttrList As String)
    On Error GoTo errHandler
    If navFirstChild Then
        Do
            attrConstruct AttrName, Template, PrmAttrList
        Loop While navNext
        navUP
    End If
    Exit Sub
errHandler:
    ErrorIn "ujXML.chAConstruct(AttrName,Template,PrmAttrList)", Array(AttrName, Template, _
         PrmAttrList)
End Sub

Public Function chCalc(ByVal AttrName As String, ByVal Oper As XENUM_OPER, _
        Optional ByVal FilterAttr As String, Optional ByVal FilterValues, _
        Optional ByVal Delim As String = ",") As Variant
    On Error GoTo errHandler
    Dim Res, V, dblV As Double
    If Not IsArray(FilterValues) Then FilterValues = Array(FilterValues)
    If navFirstChild Then
        Do
            If elCheckValues(FilterAttr, FilterValues) Then
                V = Me.attrV(AttrName)
                dblV = uSafeDbl(V)
                Select Case Oper
                    Case XOP_MIN: If IsEmpty(Res) Then Res = dblV Else Res = IIf(dblV < Res, dblV, Res)
                    Case XOP_MAX: If IsEmpty(Res) Then Res = dblV Else Res = IIf(dblV > Res, dblV, Res)
                    Case XOP_SUM: If IsEmpty(Res) Then Res = dblV Else Res = Res + dblV
                    Case XOP_CONCAT: If IsEmpty(Res) Then Res = V Else Res = Res & Delim & V
                    Case XOP_LIST
                        If IsEmpty(Res) Then
                            Res = V
                            Else
                            Res = uListAdd(Res, V, Delim)
                        End If
                    Case XOP_COUNT: If IsEmpty(Res) Then Res = 1 Else Res = Res + 1
                End Select
            End If 'elCheckFilter...
        Loop While navNext
        navUP
    End If
    chCalc = Res
    Exit Function
errHandler:
    ErrorIn "ujXML.chCalc(AttrName,Oper,FilterAttr,FilterValues,Delim)", Array(AttrName, Oper, _
         FilterAttr, FilterValues, Delim)
End Function

Public Sub chACopy(ByVal xFrom As ujXML, _
                Optional ByVal AttrList As String = "")
    On Error GoTo errHandler
    histPush
    xFrom.histPush
    If navFirstChild And xFrom.navFirstChild Then
        Do
            attrCopy xFrom, AttrList
        Loop While navNext And xFrom.navNext
    End If
    histGoBack
    xFrom.histGoBack
    Exit Sub
errHandler:
    ErrorIn "ujXML.chACopy(xFrom,AttrList)", Array(xFrom, AttrList)
End Sub

Public Sub chASet(ByVal AttrName As String, ByVal Value, _
        Optional ByVal FilterAtrs As String, Optional FilterValues)
    On Error GoTo errHandler
    Dim ToAssign As Boolean, arrAtrs() As String
    If FilterAtrs <> "" Then arrAtrs = Split(FilterAtrs, ",")
    If navFirstChild Then
        Do
            If FilterAtrs = "" Then
                ToAssign = True
                Else
                ToAssign = elCheckValues(arrAtrs, FilterValues)
            End If
            If ToAssign Then attrV(AttrName) = Value
        Loop While navNext
        navUP
    End If
    Exit Sub
errHandler:
    ErrorIn "ujXML.chASet(AttrName,Value,FilterAtrs,FilterValues)", Array(AttrName, Value, FilterAtrs, _
         FilterValues)
End Sub

Public Sub chASetAsLkpFrom(ByVal AttrName As String, IndexAttr As String, _
        ByVal FromList As String, Optional ByVal Delim As String = ",")
    On Error GoTo errHandler
    If navFirstChild Then
        Do
            attrSetAsLkpFrom AttrName, IndexAttr, FromList, Delim
        Loop While navNext
        navUP
    End If
    Exit Sub
errHandler:
    ErrorIn "ujXML.chASetAsLkpFrom(AttrName,IndexAttr,FromList,Delim)", Array(AttrName, IndexAttr, _
         FromList, Delim)
End Sub

Public Sub chASetAsIndexFrom(ByVal AttrName As String, ByVal ValueAttr As String, _
        ByVal FromList As String, Optional ByVal Delim As String = ",")
    On Error GoTo errHandler
    If navFirstChild Then
        Do
            attrSetAsIndexFrom AttrName, ValueAttr, FromList, Delim
        Loop While navNext
        navUP
    End If
    Exit Sub
errHandler:
    ErrorIn "ujXML.chASetAsIndexFrom(AttrName,ValueAttr,FromList,Delim)", Array(AttrName, ValueAttr, _
         FromList, Delim)
End Sub


Public Sub chCopyElements(ByVal xFrom As ujXML, Optional ClearExisting As Boolean = True)
    On Error GoTo errHandler
    Dim El As msxml2.IXMLDOMElement
    If ClearExisting Then chRemoveAll
    With xFrom
        If .navFirstChild Then
            Do
                chCreate "Temp"
                elXML = .elXML
                navUP
            Loop While .navNext
        End If
    End With
    Exit Sub
errHandler:
    ErrorIn "ujXML.chCopyElements(xFrom,ClearExisting)", Array(xFrom, ClearExisting)
End Sub

Public Sub chJoinLeft(ByVal xWithList As ujXML, ByVal OnAttr As String, _
        Optional ByVal CopyList As String, Optional ByVal TblBkmk As String)
    On Error GoTo errHandler
    Dim SrchAttr As String, SrcAttr As String, arrT() As String
    arrT = Split(OnAttr, "=")
    SrcAttr = arrT(0)
    If UBound(arrT) = 0 Then SrchAttr = SrcAttr Else SrchAttr = arrT(1)
    If navFirstChild Then
        Do
            If xWithList.tblLocateElem(TblBkmk, SrchAttr, attrS(SrcAttr)) Then attrCopy xWithList, CopyList
        Loop While navNext
        navUP
    End If
    Exit Sub
errHandler:
    ErrorIn "ujXML.chJoinLeft(xWithList,OnAttr,CopyList,TblBkmk)", Array(xWithList, OnAttr, _
         CopyList, TblBkmk)
End Sub

Public Sub chJoinRight(ByVal xWithList As ujXML, ByVal OnAttr As String, _
        ByVal CopyList As String, ByVal NewElName As String, _
        Optional ByVal TblBkmk As String, Optional ByVal NewElAttrList As String, _
        Optional ByVal arrNewElValues)
    On Error GoTo errHandler
    Dim SrchAttr As String, SrcAttr As String, arrT() As String, KeyV, i As Long
    Dim tListEl As msxml2.IXMLDOMElement, tLastIns As msxml2.IXMLDOMElement
    arrT = Split(OnAttr, "=")
    SrcAttr = arrT(0)
    If UBound(arrT) = 0 Then SrchAttr = SrcAttr Else SrchAttr = arrT(1)
    Set tListEl = mElem
    With xWithList
        If TblBkmk <> "" Then .tblLocate TblBkmk
        If .navLastChild Then
            Do
                Set mElem = tListEl
                If chLocate(, SrcAttr, .attrS(SrchAttr)) Then
                    attrCopy xWithList, CopyList
                    Else
                    Set mElem = tListEl
                    chCreate NewElName, , False
                    attrCopy xWithList, CopyList
                    attrSetArr NewElAttrList, arrNewElValues
                    If docOptions(XO_TRACK_CHANGES) Then elStatus = XES_JOINED
                End If
                'Insert it at the end
                tListEl.removeChild mElem
                'Set mElem.parentNode = Nothing
                If tLastIns Is Nothing Then
                    tListEl.appendChild mElem
                    Else
                    tListEl.insertBefore mElem, tLastIns
                End If
                Set tLastIns = mElem
            Loop While .navPrevious
        End If
    End With
    Set mElem = tListEl
    Exit Sub
errHandler:
    ErrorIn "ujXML.chJoinRight(xWithList,OnAttr,CopyList,NewElName,TblBkmk,NewElAttrList," & _
        "arrNewElValues)", Array(xWithList, OnAttr, CopyList, NewElName, TblBkmk, NewElAttrList, _
         arrNewElValues)
End Sub

Public Sub chSort(ByVal AttrTypeDirSpecs As String, _
            Optional ByVal DefaultDir As String = "asc")
    On Error GoTo errHandler
    Dim arrSpecs() As String, arrNames() As String, arrTypes() As String, arrDirs() As String
    Dim UbndSpecs As Long, arrSDet() As String, iElem As Long, jSpec As Long
    Dim mElemCnt As Long, arrElems(), arrV(), arrElValues()
    If AttrTypeDirSpecs = "" Then Exit Sub
    AttrTypeDirSpecs = Replace(AttrTypeDirSpecs, vbNewLine, ",")
    AttrTypeDirSpecs = Replace(AttrTypeDirSpecs, ",,", ",")
    mElemCnt = chCount
    If chCount <= 1 Then Exit Sub
    Check uInList("asc,desc", DefaultDir), XERR_UJXML, "Invalid DefaultDir specification: " & DefaultDir
    arrSpecs = Split(AttrTypeDirSpecs, ",")     'Unpack AttrTypeDirSpecs
    UbndSpecs = UBound(arrSpecs)
    ReDim arrNames(0 To UbndSpecs), arrTypes(0 To UbndSpecs), arrDirs(0 To UbndSpecs)
    For iElem = 0 To UbndSpecs
        arrSDet = Split(arrSpecs(iElem), ":")
        arrNames(iElem) = arrSDet(0)
        If UBound(arrSDet) < 1 Then arrTypes(iElem) = "string" Else arrTypes(iElem) = arrSDet(1)
        If UBound(arrSDet) < 2 Then arrDirs(iElem) = DefaultDir Else arrDirs(iElem) = LCase$(arrSDet(2))
        Check uInList("asc,desc", arrDirs(iElem)), XERR_UJXML, "Invalid Dir specification: " & arrDirs(iElem)
    Next iElem
    'Now create mElements array, and array of [Values array]
    ReDim arrElems(0 To mElemCnt - 1), arrV(0 To mElemCnt - 1)
    navFirstChild
    For iElem = 0 To mElemCnt - 1
        Set arrElems(iElem) = mElem
        arrElValues = uCreateArr(UbndSpecs)
        For jSpec = 0 To UbndSpecs
            Select Case LCase(arrTypes(jSpec))
                Case "string": arrElValues(jSpec) = attrS(arrNames(jSpec))
                Case "number": arrElValues(jSpec) = attrF(arrNames(jSpec))
                Case "date": arrElValues(jSpec) = attrDt(arrNames(jSpec))
                Case Else:
                    navUP
                    Check False, XERR_UJXML, "Invalid type specification [%1]. Must be one of the following: Date,String,Number", arrTypes(jSpec)
            End Select
        Next jSpec
        arrV(iElem) = arrElValues
        navNext
    Next iElem
    sortShell arrElems, arrV, arrDirs
    navUP
    'Now order mElements --------------------------------------------------------
    Dim tLastInsEl As IXMLDOMElement, tCurrEl As IXMLDOMElement
    For iElem = mElemCnt - 1 To 0 Step -1
        Set tCurrEl = arrElems(iElem)
        mElem.removeChild tCurrEl
        If tLastInsEl Is Nothing Then
            mElem.appendChild tCurrEl
            Else
            mElem.insertBefore tCurrEl, tLastInsEl
        End If
        Set tLastInsEl = tCurrEl
    Next iElem
    Exit Sub
errHandler:
    ErrorIn "ujXML.chSort(AttrTypeDirSpecs,DefaultDir)", Array(AttrTypeDirSpecs, DefaultDir)
End Sub

Private Sub sortShell(ByRef arrElems(), ByRef arrCompV(), ByRef arrDirs() As String)
    On Error GoTo errHandler
    Dim NewFP As Long, NewLP As Long
    Dim Dist As Long, FwdLimit As Long, Cnt As Long
    Dim FPtr As Long, LPtr As Long, ExchFlag As Boolean
    Cnt = UBound(arrElems)
    Dist = Cnt / 2
    Do Until Dist < 1
        FwdLimit = Cnt - Dist
        FPtr = 0
        Do Until FPtr > FwdLimit
            NewFP = FPtr
            LPtr = FPtr + Dist
            If (sortCompare(arrCompV(NewFP), arrCompV(LPtr), arrDirs) = 1) Then
                ExchFlag = True
                Do While ExchFlag
                    sortSwap arrElems, NewFP, LPtr
                    sortSwap arrCompV, NewFP, LPtr
                    NewLP = NewFP
                    NewFP = NewFP - Dist
                    If NewFP < LBound(arrElems) Then
                        ExchFlag = False
                    Else
                        If (sortCompare(arrCompV(NewFP), arrCompV(NewLP), arrDirs) = 1) Then
                            LPtr = NewLP
                        Else
                            ExchFlag = False
                        End If
                    End If
                Loop
            End If
            FPtr = FPtr + 1
        Loop
    Dist = Dist / 2
    Loop
    Exit Sub
errHandler:
    ErrorIn "ujXML.sortShell(arrElems,arrCompV,arrDirs)", Array(arrElems, arrCompV, arrDirs)
End Sub

Private Function sortCompare(ByRef V1, ByRef V2, arrDirs() As String) As Long
    On Error GoTo errHandler
    Dim i As Long
    sortCompare = 0
    For i = LBound(V1) To UBound(V1)
        If V1(i) < V2(i) Then
            sortCompare = -1
        ElseIf V1(i) > V2(i) Then
            sortCompare = 1
        End If
        If arrDirs(i) = "desc" Then sortCompare = -sortCompare
        If sortCompare <> 0 Then
            Exit Function
        End If
    Next i
    Exit Function
errHandler:
    ErrorIn "ujXML.sortCompare(V1,V2,arrDirs)", Array(V1, V2, arrDirs)
End Function

Private Sub sortSwap(ByRef arr(), ByVal Ind1 As Long, ByVal Ind2 As Long)
    On Error GoTo errHandler
    Dim Tmp
    If IsObject(arr(Ind1)) Then
        Set Tmp = arr(Ind1)
        Set arr(Ind1) = arr(Ind2)
        Set arr(Ind2) = Tmp
        Else
        Tmp = arr(Ind1)
        arr(Ind1) = arr(Ind2)
        arr(Ind2) = Tmp
    End If
    Exit Sub
errHandler:
    ErrorIn "ujXML.sortSwap(Arr,Ind1,Ind2)", Array(arr, Ind1, Ind2)
End Sub

Public Function chFindDuplicates(ByVal AttrList As String, _
        Optional ByVal VisOnly As Boolean = False) As String
    On Error GoTo errHandler
    Dim strDuplicates As String
    Dim arrANms() As String, strElemValues As String, strValues As String, i As Long
    strDuplicates = ""
    arrANms = Split(AttrList, ",")
    histPush
    If navFirstChild Then
        Do
            strElemValues = ""
            For i = 0 To UBound(arrANms)
                strElemValues = uCombine(strElemValues, vbTab, attrS(arrANms(i)))
                Next i
            If InStr(1, strValues, strElemValues) > 0 Then
                'Add strElemValues to duplicates if it's not there
                If InStr(1, strDuplicates, strElemValues) = 0 Then strDuplicates = uCombine(strDuplicates, vbNewLine, strElemValues)
                Else
                'add strElemValues to values list
                strValues = uCombine(strValues, ";", strElemValues)
            End If 'InStr(1, strValues....
        Loop While navNext
    End If
    chFindDuplicates = strDuplicates
    Exit Function
errHandler:
    ErrorIn "ujXML.chFindDuplicates(AttrList,VisOnly)", Array(AttrList, VisOnly)
End Function

Public Sub chRemoveAll()
    On Error GoTo errHandler
    If navFirstChild Then
        While elRemove
        Wend
    End If
    Exit Sub
errHandler:
    ErrorIn "ujXML.chRemoveAll"
End Sub


'Document ========================================================================================
'Use docInit when you're going to build xml document from scratch. Use docLoad to when you're
'going to manipulate with xml message that you already have.
Public Sub docInit(ByVal RootElemName As String, Optional ByVal AnOptions)
    On Error GoTo errHandler
    CheckSetSysHandler
    Set mDoc = CreateDocObject
    Set mTopElem = mDoc.createElement(RootElemName)
    mDoc.appendChild mTopElem
    Set mElem = mTopElem
    If Not IsMissing(AnOptions) Then mOptions = AnOptions
    mDocModified = False
    Exit Sub
errHandler:
    ErrorIn "ujXML.docInit(RootElemName)", RootElemName
End Sub

Public Sub docLoadXML(ByVal xmlData As String)
    On Error GoTo errHandler
    CheckSetSysHandler
    Set mDoc = CreateDocObject
    Check mDoc.loadXML(xmlData), ERR_GENERAL, "LoadXML method failed"
    Set mTopElem = docRootElem ' mDoc.selectSingleNode("*[nodeType()=1]") ' .firstChild
    Set mElem = mTopElem
    mDocModified = False
    Exit Sub
errHandler:
    ErrGetFromServer New ujEEDomDoc, mDoc
    ErrorIn "ujXML.docLoadXML(xmlData)", xmlData
End Sub

Private Function CreateDocObject() As Object
    On Error GoTo errHandler
    Set CreateDocObject = CreateObject(docProgID)
    Exit Function
errHandler:
    ErrorIn "ujXML.CreateDocObject"
End Function

Public Property Get docProgID() As String
    docProgID = IIf(mDocProgID = "", S_DEFAULT_PROGID, mDocProgID)
End Property

Public Property Let docProgID(ByVal Value As String)
    mDocProgID = Value
End Property

Public Property Get docObject() As IXMLDOMDocument
  Set docObject = mDoc
End Property

Public Property Set docObject(ByVal ADoc As IXMLDOMDocument)
    On Error GoTo errHandler
    Set mDoc = ADoc
    Set mTopElem = docRootElem
    Set mElem = mTopElem
    bkmkClearAll
    histClear
    Exit Property
errHandler:
    ErrorIn "ujXML.docObject(ADoc)", ADoc
End Property

Public Property Get docRootElem() As IXMLDOMElement
    On Error GoTo errHandler
    Dim Nd As IXMLDOMNode
    Set Nd = mDoc.selectSingleNode("*[nodeType()=1]")
    Check Not (Nd Is Nothing), XERR_UJXML, "Failed to find root node." ' .documentElement"
    Set docRootElem = Nd
    Exit Property
errHandler:
    ErrorIn "ujXML.docRootElem"
End Property

Public Property Get docOptions(ByVal mask As XENUM_OPTIONS) As Boolean
    On Error GoTo errHandler
    docOptions = ((mOptions And mask) <> 0)
    Exit Property
errHandler:
    ErrorIn "ujXML.docOptions(mask)", mask
End Property

Public Property Let docOptions(ByVal mask As XENUM_OPTIONS, ByVal Value As Boolean)
    On Error GoTo errHandler
    If Value Then mOptions = mOptions Or mask Else mOptions = mOptions And Not mask
    Exit Property
errHandler:
    ErrorIn "ujXML.docOptions(mask,Value)", Array(mask, Value)
End Property

Public Sub docSetOptions(ByVal AnOptions As Long)
    mOptions = AnOptions
End Sub

Public Property Get docModified() As Boolean
    On Error GoTo errHandler
    docModified = mDocModified
    Exit Property
errHandler:
    ErrorIn "ujXML.docModified"
End Property

Public Property Let docModified(ByVal AValue As Boolean)
    On Error GoTo errHandler
    mDocModified = AValue
    Exit Property
errHandler:
    ErrorIn "ujXML.docModified(AValue)", AValue
End Property

Public Property Get docXML() As String
    On Error GoTo errHandler
    docXML = mDoc.xml ' mTopElem.xml
    Exit Property
errHandler:
    ErrorIn "ujXML.docXML"
End Property

Public Sub docReadFromFile(ByVal FilePath As String, Optional ByVal CharSet As String = "UNICODE")
    On Error GoTo errHandler
    Dim sXml As String
    Dim S As Object
    Set S = CreateObject("ADODB.Stream")
    With S
        If CharSet <> "" Then .CharSet = CharSet
        .Open
        .LoadFromFile FilePath
        sXml = .ReadText
        .Close
    End With
    docLoadXML sXml
    mDocModified = False
    Exit Sub
errHandler:
    ErrorIn "ujXML.docReadFromFile(FilePath,CharSet)", Array(FilePath, CharSet)
End Sub

Public Sub docWriteToFile(ByVal FilePath As String, Optional ByVal CurrElementOnly As Boolean = False, _
                Optional ByVal CharSet As String = "UNICODE")
    On Error GoTo errHandler
    Dim S As Object
    Set S = CreateObject("ADODB.Stream")
    With S
        If CharSet <> "" Then .CharSet = CharSet
        .Open
        If CurrElementOnly Then .WriteText elXML Else .WriteText docXML
        .SaveToFile FilePath, 2 'adSaveCreateOverWrite
        .Close
    End With
    Exit Sub
errHandler:
    ErrorIn "ujXML.docWriteToFile(FilePath,CurrElementOnly,Charset)", Array(FilePath, CurrElementOnly, CharSet)
End Sub

Public Sub docAFormat(ByVal AttrList As String, ByVal Fmt As String)
    On Error GoTo errHandler
    Dim Nodes As msxml2.IXMLDOMNodeList, Nd As msxml2.IXMLDOMNode, D As Date
    Set Nodes = mTopElem.selectNodes(".//(@" & Replace(AttrList, ",", "|@") & ")")
    For Each Nd In Nodes
        If IsNull(Nd.nodeValue) Then
            'Nd.nodeValue = Null
            Else
            Nd.nodeValue = uSafeFormat(Nd.nodeValue, Fmt)
        End If 'Nd.nodeType ...
    Next Nd
    Exit Sub
errHandler:
    ErrorIn "ujXML.docAFormat(AttrList,Fmt)", Array(AttrList, Fmt)
End Sub

Public Sub docAFormatDates(ByVal AttrList As String, ByVal Fmt As String, Optional ByVal ZeroAsEmpty As Boolean = True)
    On Error GoTo errHandler
    Dim Nodes As msxml2.IXMLDOMNodeList, Nd As msxml2.IXMLDOMNode, D As Date
    Set Nodes = mTopElem.selectNodes(".//(@" & Replace(AttrList, ",", "|@") & ")")
    For Each Nd In Nodes
        D = uSafeDate(Nd.nodeValue)
        If D = 0 And ZeroAsEmpty Then
            Nd.nodeValue = ""
            Else
            Nd.nodeValue = uSafeFormat(D, Fmt)
        End If
    Next Nd
    Exit Sub
errHandler:
    ErrorIn "ujXML.docAFormatDates(AttrList,Fmt,ZeroAsEmpty)", Array(AttrList, Fmt, ZeroAsEmpty)
End Sub

Public Sub docAppendXML(ByVal xmlData As String, Optional ByVal bkmkAs As String)
    On Error GoTo errHandler
    Dim xTmp As ujXML
    Set xTmp = New ujXML
    xTmp.docProgID = docProgID
    xTmp.docLoadXML xmlData
    docRootElem.appendChild xTmp.Element
    Set mElem = xTmp.Element
    Exit Sub
errHandler:
    ErrorIn "ujXML.docAppendXML(xmlData,bkmkAs)", Array(xmlData, bkmkAs)
End Sub

Public Function docXslLocate(ByVal XslPath As String, Optional ByVal Params, _
        Optional ByVal FromTop As Boolean = True) As Boolean
    On Error GoTo errHandler
    Dim tEl As IXMLDOMElement, Nd As IXMLDOMNode
    Set tEl = IIf(FromTop, mElem, docRootElem)
    If Not IsMissing(Params) Then XslPath = uInsInto(XslPath, Params, False)
    Set Nd = tEl.selectSingleNode(XslPath)
    If Nd Is Nothing Then Exit Function
    Set mElem = Nd
    docXslLocate = True
    Exit Function
errHandler:
    ErrorIn "ujXML.docXslLocate(XslPath,Params,FromTop)", Array(XslPath, Params, _
         FromTop)
End Function

Public Sub docPersistPrepare()
    On Error GoTo errHandler
    bkmkSaveToA
    docTopSetAttr S_DICT, mDict
    Exit Sub
errHandler:
    ErrorIn "ujXML.docPersistPrepare"
End Sub

Public Sub docPersistRestore()
    On Error GoTo errHandler
    Set mhistColl = Nothing
    Set mbkmkColl = Nothing
    Set mSavedEl = Nothing
    mBkmkList = ""
    bkmkRestoreFromA
    mDict = docTopGetAttr(S_DICT)
    Exit Sub
errHandler:
    ErrorIn "ujXML.docPersistRestore"
End Sub

Public Sub docPersistCleanup()
    On Error GoTo errHandler
    bkmkClearA
    mTopElem.removeAttribute S_DICT
    Exit Sub
errHandler:
    ErrorIn "ujXML.docPersistClear"
End Sub

Public Function docTopGetAttr(ByVal AttrName As String) As Variant
    On Error GoTo errHandler
    docTopGetAttr = mTopElem.getAttribute(AttrName)
    Exit Function
errHandler:
    ErrorIn "ujXML.docTopGetAttr(AttrName)", AttrName
End Function

Public Sub docTopSetAttr(ByVal AttrName As String, ByVal Value)
    On Error GoTo errHandler
    mTopElem.setAttribute AttrName, Value
    Exit Sub
errHandler:
    ErrorIn "ujXML.docTopSetAttr(AttrName,Value)", Array(AttrName, Value)
End Sub

Public Function docCreateViewer(Optional ByVal UseCurrentAsTop As Boolean = True, _
            Optional ByVal ADict As String, Optional AnOptions) As ujXML
    On Error GoTo errHandler
    Dim x As ujXML
    Set x = New ujXML
    With x
        Set .docObject = docObject
        Set .docTopElem = IIf(UseCurrentAsTop, mElem, docTopElem)
        Set .Element = .docTopElem
        .docSetOptions IIf(IsMissing(AnOptions), mOptions, AnOptions)
        If ADict <> "" Then
            .dictList = ADict
        ElseIf .attrS(S_DICT) <> "" Then
            .dictList = .attrS(S_DICT)
        Else
            .dictList = dictList
        End If
        .docOptions(XO_ISVIEWER) = True
        .docModified = False
    End With 'x
    Set docCreateViewer = x
    Exit Function
errHandler:
    ErrorIn "ujXML.docCreateViewer(UseCurrentAsTop,ADict,AnOptions)", Array(UseCurrentAsTop, ADict, _
         AnOptions)
End Function

Public Property Get docIsViewer() As Boolean
    docIsViewer = docOptions(XO_ISVIEWER)
End Property

Public Property Get docTopElem() As IXMLDOMElement
    Set docTopElem = mTopElem
End Property

Public Property Set docTopElem(ByVal AnElem As IXMLDOMElement)
    Set mTopElem = AnElem
End Property

Public Sub docAfterUpdate(Optional ByVal DontDeleteElems As String = "")
    On Error GoTo errHandler
    Dim Ndl As IXMLDOMNodeList, i As Long
    navTop
    Set Ndl = docTopElem.selectNodes(".//*[@" & S_STATUS & "]")
    For i = 0 To Ndl.Length - 1
        Set mElem = Ndl.Item(i)
        Select Case elStatus
            Case XES_DELETED:  If uInList(DontDeleteElems, elName) Then elStatus = XES_JOINED Else elRemove
            Case XES_MODIFIED, XES_NEW: elStatus = XES_EXISTING
            Case XES_EXISTING, XES_JOINED: 'Nothing to do
        End Select
    Next i
    navTop
    Exit Sub
errHandler:
    ErrorIn "ujXML.docAfterUpdate(DontDeleteElems)", DontDeleteElems
End Sub

Public Sub docCreateElem(ByVal ElemName As String, _
        Optional ByVal BeforeEl As IXMLDOMElement, _
        Optional ByVal AfterEl As IXMLDOMElement, _
        Optional ByVal AsChildOf As IXMLDOMElement)
    On Error GoTo errHandler
    Dim NewEl As IXMLDOMElement
    Set NewEl = mDoc.createElement(ElemName)
    Check Not ((BeforeEl Is docRootElem) Or (AfterEl Is docRootElem)), XERR_UJXML, _
        "Before/After reference may not be document root."
    If AfterEl Is Nothing Then
        If BeforeEl Is Nothing Then
            Check Not (AsChildOf Is Nothing), XERR_UJXML, "At least one reference mElement must be provided."
            AsChildOf.appendChild NewEl
            Else
            BeforeEl.parentNode.insertBefore NewEl, BeforeEl
        End If
      Else
        If AfterEl.nextSibling Is Nothing Then
            AfterEl.parentNode.appendChild NewEl
            Else
            AfterEl.parentNode.insertBefore NewEl, AfterEl.nextSibling
        End If
    End If
    Set mElem = NewEl
    Exit Sub
errHandler:
    ErrorIn "ujXML.docCreateElem(ElemName,BeforeEl,AfterEl,AsChildOf)", Array(ElemName, BeforeEl, _
         AfterEl, AsChildOf)
End Sub

'########################################### Bookmarks #########################################
Public Property Get bkmkColl() As Collection
    If mbkmkColl Is Nothing Then Set mbkmkColl = New Collection
    Set bkmkColl = mbkmkColl
End Property

Public Sub bkmkSet(ByVal ABkmk As String)
    On Error GoTo errHandler
    If bkmkExists(ABkmk) Then bkmkRemove ABkmk
    bkmkColl.Add mElem, ABkmk
    mBkmkList = uListAdd(mBkmkList, ABkmk)
    Exit Sub
errHandler:
    ErrorIn "ujXML.bkmkSet(ABkmk)", ABkmk
End Sub

Public Sub bkmkClearAll()
    On Error Resume Next
    Set mbkmkColl = Nothing
    mBkmkList = ""
End Sub

Public Function bkmkRemove(ByVal ABkmk As String) As Boolean
    On Error Resume Next
    bkmkColl.Remove ABkmk
    mBkmkList = uListRemove(mBkmkList, ABkmk)
End Function

Private Function bkmkRemoveEl(ByVal El As IXMLDOMElement) As Boolean
    On Error GoTo errHandler
    Dim i As Long
    With bkmkColl
        For i = .Count To 1 Step -1
            If uIsChildOf(.Item(i), El) Then .Remove i
        Next i
    End With
    Exit Function
errHandler:
    ErrorIn "ujXML.bkmkRemoveEl(El)", El
End Function

Private Function bkmkGetElem(ByVal ABkmk As String) As msxml2.IXMLDOMElement
    On Error Resume Next
    Set bkmkGetElem = bkmkColl(ABkmk)
End Function

Public Function bkmkExists(ByVal ABkmk As String) As Boolean
    On Error GoTo errHandler
    bkmkExists = uInList(mBkmkList, ABkmk)
    Exit Function
errHandler:
    ErrorIn "ujXML.bkmkExists(ABkmk)", ABkmk
End Function

Public Function bkmkGoto(ByVal ABkmk As String, Optional ByVal RaiseErr As Boolean = True) As Boolean
    On Error GoTo errHandler
    Dim tEl As msxml2.IXMLDOMElement
    Set tEl = bkmkGetElem(ABkmk)
    If tEl Is Nothing Then
        Check Not RaiseErr, XERR_UJXML, "Bookmark [%1] does not exist.", ABkmk
        Else
        Set mElem = tEl
        bkmkGoto = True
    End If
    Exit Function
errHandler:
    ErrorIn "ujXML.bkmkGoto(ABkmk,RaiseErr)", Array(ABkmk, RaiseErr)
End Function

Private Sub bkmkSaveToA()
    On Error GoTo errHandler
    Dim arrB() As String, i As Long
    arrB = Split(mBkmkList, ",")
    For i = LBound(arrB) To UBound(arrB)
        If bkmkExists(arrB(i)) Then
            bkmkGoto arrB(i)
            attrS(S_BKMK) = uListAdd(attrS(S_BKMK), arrB(i))
        End If
    Next i
    Exit Sub
errHandler:
    ErrorIn "ujXML.bkmkSaveToA"
End Sub

Private Sub bkmkClearA()
    On Error GoTo errHandler
    Dim Ndl As msxml2.IXMLDOMNodeList, tEl As msxml2.IXMLDOMElement
    Set Ndl = docTopElem.selectNodes("//*[@" & S_BKMK & "]")
    For Each tEl In Ndl
        tEl.removeAttribute S_BKMK
    Next tEl
    Exit Sub
errHandler:
    ErrorIn "ujXML.bkmkClearA"
End Sub

Private Sub bkmkRestoreFromA()
    On Error GoTo errHandler
    Set mbkmkColl = Nothing
    mBkmkList = ""
    histPush
    Dim Ndl As msxml2.IXMLDOMNodeList, tEl As msxml2.IXMLDOMElement, arrB() As String, i As Long
    Set Ndl = docTopElem.selectNodes("//*[@" & S_BKMK & "]")
    For Each tEl In Ndl
        Set mElem = tEl
        arrB = Split(attrS(S_BKMK), ",")
        For i = LBound(arrB) To UBound(arrB)
            If arrB(i) <> "" Then
                bkmkSet arrB(i)
            End If
        Next i
    Next tEl
    histGoBack
    Exit Sub
errHandler:
    ErrorIn "ujXML.bkmkRestoreFromA"
End Sub



'History =================================================================================================
Private Function histColl() As Collection
    On Error GoTo errHandler
    If mhistColl Is Nothing Then Set mhistColl = New Collection
    Set histColl = mhistColl
    Exit Function
errHandler:
    ErrorIn "ujXML.histColl"
End Function

Public Sub histPush()
    On Error GoTo errHandler
    histColl.Add mElem
    Exit Sub
errHandler:
    ErrorIn "ujXML.elPush"
End Sub

Public Property Get histCanGoBack() As Boolean
    On Error GoTo errHandler
    histCanGoBack = (histColl.Count > 0)
    Exit Property
errHandler:
    ErrorIn "ujXML.histCanGoBack"
End Property

Public Sub histGoBack(Optional ByVal Purge As Boolean = True)
    On Error GoTo errHandler
    Check histCanGoBack, XERR_UJXML, "History is empty."
    With histColl
        Set mElem = .Item(.Count)
    End With
    If Purge Then histPurge
    Exit Sub
errHandler:
    ErrorIn "ujXML.histGoBack(Purge)", Purge
End Sub

Public Sub histPurge()
    On Error GoTo errHandler
    Check histCanGoBack, XERR_UJXML, "History is empty."
    With histColl
        .Remove .Count
    End With
    Exit Sub
errHandler:
    ErrorIn "ujXML.histPurge"
End Sub

Public Sub histClear()
    Set mhistColl = New Collection
End Sub

Private Function histRemoveEl(ByVal El As IXMLDOMElement) As Boolean
    On Error GoTo errHandler
    Dim i As Long
    With histColl
        For i = .Count To 1 Step -1
            If uIsChildOf(.Item(i), El) Then .Remove i
        Next i
    End With
    Exit Function
errHandler:
    ErrorIn "ujXML.histRemoveEl(El)", El
End Function

'################################################# Dictionary ##############################################
Public Property Get dictList() As String
    dictList = mDict
End Property

Public Property Let dictList(ByVal AList As String)
    mDict = AList
    docOptions(XO_USE_DICT) = (mDict <> "")
End Property

Public Function dictCheckName(ByVal AName As String, _
        Optional RaiseErr As Boolean = True) As Boolean
    On Error GoTo errHandler
    If uInList(mDict, AName) Then
        dictCheckName = True
        Else
        Check Not RaiseErr, XERR_UJXML, "Name [%1] is not in XML dictionary; Current mElement: %2.", _
            Array(AName, elName)
    End If
    Exit Function
errHandler:
    ErrorIn "ujXML.dictCheckName(AName,RaiseErr)", Array(AName, RaiseErr), , , "Dictionary", mDict
End Function

Public Function dictCheckDoc() As String
    On Error GoTo errHandler
    Dim Ndl As IXMLDOMNodeList, i As Long, NdName As String, sRes As String
    Set Ndl = mTopElem.selectNodes("//@*")
    For i = 0 To Ndl.Length - 1
        NdName = Ndl.Item(i).nodeName
        If Not uInList(mDict, NdName) Then sRes = uListAdd(sRes, NdName)
    Next i
    dictCheckDoc = sRes
    Exit Function
errHandler:
    ErrorIn "ujXML.dictCheckDoc"
End Function

Public Sub dictAdd(ByVal NameList As String)
    On Error GoTo errHandler
    Dim arrL() As String, i As Long
    arrL = Split(NameList, ",")
    For i = 0 To UBound(arrL)
        If Not uInList(mDict, arrL(i)) Then mDict = uListAdd(mDict, arrL(i))
    Next i
    docOptions(XO_USE_DICT) = (mDict <> "")
    Exit Sub
errHandler:
    ErrorIn "ujXML.dictAdd(NameList)", NameList
End Sub

Public Sub dictRemove(ByVal NameList As String)
    On Error GoTo errHandler
    Dim arrL() As String, i As Long
    arrL = Split(NameList, ",")
    For i = 0 To UBound(arrL)
        If uInList(mDict, arrL(i)) Then mDict = uListRemove(mDict, arrL(i))
    Next i
    docOptions(XO_USE_DICT) = (mDict <> "")
    Exit Sub
errHandler:
    ErrorIn "ujXML.dictRemove(NameList)", NameList
End Sub

'Table ====================================================================================================
Public Sub tblAdd(ByVal TblBkmk As String, ByVal xmlList As String)
    On Error GoTo errHandler
    docAppendXML xmlList, TblBkmk
    bkmkSet TblBkmk
    Exit Sub
errHandler:
    ErrorIn "ujXML.tblAdd(TblBkmk,xmlList)", Array(TblBkmk, xmlList)
End Sub

Public Sub tblRemove(ByVal TblBkmk As String)
    On Error GoTo errHandler
    bkmkGoto TblBkmk
    elRemove
    Exit Sub
errHandler:
    ErrorIn "ujXML.tblRemove(TblBkmk)", TblBkmk
End Sub

Public Sub tblLocate(ByVal TblBkmk As String)
    On Error GoTo errHandler
    bkmkGoto TblBkmk
    Exit Sub
errHandler:
    ErrorIn "ujXML.tblLocate(TblBkmk)", TblBkmk
End Sub

Public Function tblLookup(ByVal TblBkmk As String, _
        ByVal KeyAtrs As String, ByVal KeyValues, _
        ByVal ReturnAttr As String) As String
    On Error GoTo errHandler
    Check tblLocateElem(TblBkmk, KeyAtrs, KeyValues), XERR_UJXML, "Lookup failed."
    tblLookup = attrS(ReturnAttr)
    Exit Function
errHandler:
    ErrorIn "ujXML.tblLookup(TblBkmk,KeyAtrs,KeyValues,ReturnAttr)", Array(TblBkmk, KeyAtrs, _
         KeyValues, ReturnAttr)
End Function

Public Function tblLocateElem(ByVal TblBkmk As String, ByVal SrchAttrList As String, ByVal SrchValues) As Boolean
    On Error GoTo errHandler
    tblLocate TblBkmk
    tblLocateElem = chLocate(, SrchAttrList, SrchValues)
    Exit Function
errHandler:
    ErrorIn "ujXML.tblLocateElem(TblBkmk,SrchAttrList,SrchValues)", Array(TblBkmk, SrchAttrList, _
         SrchValues)
End Function

Public Sub tblLoadTables(ByVal Connection, ByVal TblBkmks As String, _
            Optional ByVal SQLs, Optional ByVal ExecInBatch As Boolean = True)
    On Error GoTo errHandler
    Dim Conn As Object, SQL As String, sBuf As String
    Dim arrTN, arrDTN, Ubnd As Long, i As Long, RS As Object
    arrTN = Split(TblBkmks, ",")
    Ubnd = UBound(arrTN)
    If IsMissing(SQLs) Then
        SQLs = uCreateArr(Ubnd)
        For i = 0 To Ubnd
           SQLs(i) = uInsInto(SQL_SELECT, arrTN(i))
        Next i
      Else
        If Not IsArray(SQLs) Then SQLs = Array(SQLs)
        Check UBound(SQLs) = Ubnd, XERR_UJXML, "List counts mismatch (tables and SQLs)"
    End If 'IsMissing...
    If IsObject(Connection) Then
        Set Conn = Connection
        Else
        Set Conn = CreateObject("ADODB.Connection")
        Conn.Open Connection
    End If
    If ExecInBatch Then
        Set RS = Conn.Execute(Join(SQLs, "; " & vbNewLine))
        For i = 0 To Ubnd
            sBuf = uExtractRecords(RS, "Item", arrTN(i))
            tblAdd arrTN(i), sBuf
            Set RS = RS.NextRecordset
        Next i
      Else
        For i = 0 To Ubnd
            Set RS = Conn.Execute(SQLs(i))
            sBuf = uExtractRecords(RS, "Item", arrTN(i))
            tblAdd arrTN(i), sBuf
        Next i
    End If
    Exit Sub
errHandler:
    ErrorIn "ujXML.tblLoadTables(Connection,TblBkmks,SQLs,ExecInBatch)", Array(Connection, TblBkmks, _
         SQLs, ExecInBatch)
End Sub


'Transactions ================================================================================================
Public Sub tranBegin()
    On Error GoTo errHandler
    If docIsViewer Then Exit Sub
    docPersistPrepare
    navTop
    mtranSavedXML = elXML
    docPersistCleanup
    Exit Sub
errHandler:
    ErrorIn "ujXML.tranBegin"
End Sub

Public Sub tranCommit()
    On Error GoTo errHandler
    If docIsViewer Then Exit Sub
    mtranSavedXML = ""
    Exit Sub
errHandler:
    ErrorIn "ujXML.tranCommit"
End Sub

Public Sub tranRollback()
    On Error GoTo errHandler
    If docIsViewer Then Exit Sub
    Check mtranSavedXML <> "", XERR_UJXML, "No transaction in progress."
    docLoadXML mtranSavedXML
    docPersistRestore
    docPersistCleanup
    mtranSavedXML = ""
    navTop
    Exit Sub
errHandler:
    ErrorIn "ujXML.tranRollback"
End Sub

Public Property Get tranSavedXML() As String
    tranSavedXML = mtranSavedXML
End Property

'HTML ======================================================================================
Public Function htmElConstruct(ByVal Template As String, ByVal AttrList As String, _
        Optional ByVal SendToIISResponse As Boolean = False, _
        Optional ByVal CtrlChar As String = "%") As String
    On Error GoTo errHandler
    Dim arrA() As String, Res As String
    Template = Replace(Template, CtrlChar, uChrBk)
    arrA = Split(AttrList, ",")
    Res = htmElConstrPriv(Template, arrA)
    Res = Replace(Res, uChrBk, CtrlChar)
    htmElConstruct = Res
    If SendToIISResponse And iisInIIS Then iisWrite Res
    Exit Function
errHandler:
    ErrorIn "ujXML.htmElConstruct(Template,AttrList,CtrlChar)", Array(Template, AttrList, CtrlChar)
End Function

Public Function htmChConstruct(ByVal Template As String, ByVal AttrList As String, _
        Optional ByVal SendToIISResponse As Boolean = False, _
        Optional ByVal SelAttr As String, Optional SelValues As String, _
        Optional ShowSelAs As String = "SELECTED", Optional ByVal CtrlChar As String = "%", _
        Optional ByVal Crlf As String = "||", Optional ByVal Indent As Long = 8) As String
    On Error GoTo errHandler
    Dim arrA() As String, Res As String, i As Long
    Dim sTempl As String, S As String, sSel As String, sIndent As String
    If Indent > 0 Then
        sIndent = Space(Indent)
        sTempl = sIndent & sTempl
    End If
    sTempl = Template
    If Crlf <> "" Then sTempl = Replace(sTempl, Crlf, sIndent & vbNewLine)
    sTempl = Replace(sTempl, CtrlChar, uChrBk)
    arrA = Split(AttrList, ",")
    If navFirstChild Then
        Do
            S = htmElConstrPriv(sTempl, arrA)
            If SelAttr <> "" Then
                sSel = IIf(uInList(SelValues, attrS(SelAttr)), ShowSelAs, "")
                S = Replace(S, uChrBk & "0", sSel)
            End If
            If Res <> "" Then Res = Res & vbNewLine
            Res = Res & Space(Indent) & S
        Loop While navNext
    End If
    Res = Replace(Res, uChrBk, CtrlChar)
    htmChConstruct = Res
    If SendToIISResponse And iisInIIS Then iisWrite Res
    Exit Function
errHandler:
    ErrorIn "ujXML.htmChConstruct(Template,AttrList,SendToIISResponse,SelAttr,SelValues,ShowSelAs," & _
        "CtrlChar,Crlf,Indent)", Array(Template, AttrList, SendToIISResponse, SelAttr, SelValues, ShowSelAs, _
         CtrlChar, Crlf, Indent)
End Function

Private Function htmElConstrPriv(ByRef Template As String, ByRef arrANms) As String
    On Error GoTo errHandler
    Dim i As Long, sV As String, Res As String
    Res = Template
    For i = 0 To UBound(arrANms)
        sV = attrEscHtm(arrANms(i))
        Res = Replace(Res, uChrBk & (i + 1), sV)
    Next i
    htmElConstrPriv = Res
    Exit Function
errHandler:
    ErrorIn "ujXML.htmElConstrPriv(Template,arrANms)", Array(Template, arrANms)
End Function

'SQL ====================================================================================
Public Function sqlElConstruct(ByVal TableName As String, _
            ByVal FldEqAttrList As String, ByVal KeyFldEqAttrList As String, _
            Optional ByVal SQLType As XENUM_SQLTYPE = XSQL_AUTO) As String
    On Error GoTo errHandler
    Dim Ubnd As Long, sRes As String
    Dim arrF, arrA, arrV, arrQ
    Dim arrKF, arrKA, arrKV, arrKQ
    If SQLType = XSQL_AUTO Then SQLType = sqlElGetDftSQLType
    Ubnd = uUnpackEQList(FldEqAttrList, arrF, arrA, arrQ)
    arrV = uCreateArr(Ubnd)
    Ubnd = uUnpackEQList(KeyFldEqAttrList, arrKF, arrKA, arrKQ)
    arrKV = uCreateArr(Ubnd)
    sqlElFillArray arrA, arrQ, arrV
    sqlElFillArray arrKA, arrKQ, arrKV
    sRes = sqlElConstructPriv(TableName, SQLType, arrF, arrA, arrQ, arrV, arrKF, arrKA, arrKQ, arrKV)
    sqlElConstruct = sRes
    Exit Function
errHandler:
    ErrorIn "ujXML.sqlElConstruct(TableName,FldEqAttrList,KeyFldEqAttrList,SQLType)", Array(TableName, _
         FldEqAttrList, KeyFldEqAttrList, SQLType)
End Function

Public Function sqlChConstruct(ByVal TableName As String, _
            ByVal FldEqAttrList As String, ByVal KeyFldEqAttrList As String, _
            Optional ByVal CascDelTableNames As String, _
            Optional ByVal CascKeyFldEqAttrList As String) As Variant
    On Error GoTo errHandler
    Dim Ubnd As Long
    Dim arrF, arrA, arrV, arrQ
    Dim arrKF, arrKA, arrKV, arrKQ
    Dim arrDelF, arrDelA, arrDelV, arrDelQ, arrDelTables
    Dim colRes As Collection, arrRes, i As Long, sSQL As String, SType As XENUM_SQLTYPE, vE
    'Prepare arrays
    Set colRes = New Collection
    Ubnd = uUnpackEQList(FldEqAttrList, arrF, arrA, arrQ)
    arrV = uCreateArr(Ubnd)
    Ubnd = uUnpackEQList(KeyFldEqAttrList, arrKF, arrKA, arrKQ)
    arrKV = uCreateArr(Ubnd)
    If CascDelTableNames <> "" Then
        arrDelTables = Split(CascDelTableNames, ",")
        If CascKeyFldEqAttrList = "" Then CascKeyFldEqAttrList = KeyFldEqAttrList
        Ubnd = uUnpackEQList(CascKeyFldEqAttrList, arrDelF, arrDelA, arrDelQ)
        arrDelV = uCreateArr(Ubnd)
    End If
    'Now create SQLs
    If navFirstChild Then
        Do
            sqlElFillArray arrA, arrQ, arrV
            sqlElFillArray arrKA, arrKQ, arrKV
            SType = sqlElGetDftSQLType
            If SType = XSQL_DELETE And CascDelTableNames <> "" Then
                sqlElFillArray arrDelA, arrDelQ, arrDelV
                For i = 0 To UBound(arrDelTables)
                    sSQL = sqlElConstructPriv(arrDelTables(i), XSQL_DELETE, arrF, arrA, arrQ, arrV, arrDelF, arrDelA, arrDelQ, arrDelV)
                    colRes.Add sSQL
                Next i
            End If
            sSQL = sqlElConstructPriv(TableName, SType, arrF, arrA, arrQ, arrV, arrKF, arrKA, arrKQ, arrKV)
            If sSQL <> "" Then colRes.Add sSQL
        Loop While navNext
    End If
    If colRes.Count = 0 Then Exit Function
    ReDim arrRes(0 To colRes.Count - 1)
    For i = 0 To colRes.Count - 1
        arrRes(i) = colRes(i + 1)
    Next i
    sqlChConstruct = arrRes
    Exit Function
errHandler:
    ErrorIn "ujXML.sqlChConstruct(TableName,FldEqAttrList,KeyFldEqAttrList)", Array(TableName, FldEqAttrList, _
         KeyFldEqAttrList)
End Function

Private Function sqlElConstructPriv(ByVal TableName As String, ByVal SQLType As Long, _
                ByRef arrF, ByRef arrA, ByRef arrQ, ByRef arrVbuf, _
                ByRef arrKF, ByRef arrKA, ByRef arrKQ, ByRef arrKVbuf) As String
    On Error GoTo errHandler
'    Const _
'        SQL_INSERT = "INSERT INTO  %1 (%2) VALUES (%3);", _
'        SQL_UPDATE = "UPDATE  %1 SET %2 WHERE %3;", _
'        SQL_DELETE = "DELETE FROM %1 WHERE %2;", _
'        SQL_FIND = "SELECT %2 FROM %1 WHERE %2;"
    Select Case SQLType
        Case XSQL_UPDATE
            sqlElFillArray arrA, arrQ, arrVbuf
            sqlElFillArray arrKA, arrKQ, arrKVbuf
            sqlElConstructPriv = uInsInto(SQL_UPDATE, Array(vbNewLine, TableName, uCreateEQList(arrF, arrVbuf), uCreateEQList(arrKF, arrKVbuf, " AND ")))
        Case XSQL_INSERT
            sqlElFillArray arrA, arrQ, arrVbuf
            sqlElFillArray arrKA, arrKQ, arrKVbuf
            sqlElConstructPriv = uInsInto(SQL_INSERT, Array(vbNewLine, TableName, Join(arrF, ", "), Join(arrVbuf, ", ")))
        Case XSQL_DELETE
            sqlElFillArray arrKA, arrKQ, arrKVbuf
            sqlElConstructPriv = uInsInto(SQL_DELETE, Array(vbNewLine, TableName, uCreateEQList(arrKF, arrKVbuf, " AND ")))
    End Select
    Exit Function
errHandler:
    ErrorIn "ujXML.sqlElConstructPriv(TableName,SQLType,arrF,arrA,arrQ,arrVbuf,arrKF,arrKA,arrKQ," & _
        "arrKVbuf)", Array(TableName, SQLType, arrF, arrA, arrQ, arrVbuf, arrKF, arrKA, arrKQ, arrKVbuf)
End Function

Private Sub sqlElFillArray(ByRef arrA, ByRef arrQ, ByRef arrV)
    On Error GoTo errHandler
    Dim i As Long, sA As String
    For i = LBound(arrA) To UBound(arrA)
        sA = arrA(i)
        If attrIsNull(sA) Then
            arrV(i) = "Null"
        ElseIf arrQ(i) Then
            arrV(i) = attrQt(sA)
        Else
            arrV(i) = attrS(sA)
        End If
    Next i
    Exit Sub
errHandler:
    ErrorIn "ujXML.sqlElFillArray(arrA,arrQ,arrV)", Array(arrA, arrQ, arrV)
End Sub

Public Function sqlElGetDftSQLType() As XENUM_SQLTYPE
    On Error GoTo errHandler
    Dim T As Long
    Select Case elStatus
        Case XES_JOINED, XES_EXISTING: T = XSQL_NONE
        Case XES_MODIFIED: T = XSQL_UPDATE
        Case XES_NEW: T = XSQL_INSERT
        Case XES_DELETED: T = XSQL_DELETE
    End Select
    sqlElGetDftSQLType = T
    Exit Function
errHandler:
    ErrorIn "ujXML.sqlElGetDftSQLType"
End Function

'Private Function sqlTblConstructForCrossTable(ByVal TblBkmk As String, ByVal CrossTableName As String, _
'            ByVal LkpTableKey As String, ByVal MainTableKey As String, ByVal MainKeyValue, _
'            ByVal OldLkpKeyValues As String, ByVal NewLkpKeyValues As String) As Variant
'    On Error GoTo errHandler
'    If TblBkmk <> "" Then tblLocate TblBkmk
'    Check False, XERR_UJXML, "Method Not implemented."
'    Exit Function
'errHandler:
'    ErrorIn "ujXML.sqlTblConstructForCrossTable(TblBkmk,CrossTableName,LkpTableKey,MainTableKey," & _
'        "MainKeyValue,OldLkpKeyValues,NewLkpKeyValues)", Array(TblBkmk, CrossTableName, LkpTableKey, _
'         MainTableKey, MainKeyValue, OldLkpKeyValues, NewLkpKeyValues)
'End Function

'################################# Recordset-related functions ################################################
Public Sub rsElCopy(ByVal AdoRs As Object, _
            Optional ByVal Specs As String, Optional ByVal AddToDict As Boolean = False)
    On Error GoTo errHandler
    Dim arrF, arrA, arrQ
    If Specs = "" Then Specs = rsGetFieldList(AdoRs)
    uUnpackEQList Specs, arrA, arrF, arrQ
    If AddToDict Then dictAdd Join(arrA, ",")
    rsElCopyPriv AdoRs, arrA, arrF, arrQ
    Exit Sub
errHandler:
    ErrorIn "ujXML.rsElCopy(AdoRs,Specs,AddToDict)", Array(AdoRs, Specs, AddToDict)
End Sub

Private Sub rsElCopyPriv(ByVal AdoRs As Object, arrA, arrF, arrQ)
    On Error GoTo errHandler
    Dim i As Long, FName As String, F As Object
    For i = 0 To UBound(arrF)
        If arrQ(i) Then
            attrV(arrA(i)) = arrF(i)
            Else
            Set F = uSafeGetFieldObj(AdoRs, arrF(i))
            Check Not F Is Nothing, XERR_UJXML, "Field [%1] not found.", arrF(i)
            attrV(arrA(i)) = F.Value
        End If
    Next i
    Exit Sub
errHandler:
    ErrorIn "ujXML.rsElCopyPriv(AdoRs,arrA,arrF,arrQ)", Array(AdoRs, arrA, arrF, arrQ)
End Sub

Public Sub rsAddAsChildren(ByVal AdoRs As Object, _
                          Optional ByVal RecElemName As String = "Record", _
                          Optional ByVal AddToDict As Boolean = False)
    On Error GoTo errHandler
    Dim S As String, xBuf As ujXML, Ndl As IXMLDOMNodeList, El As IXMLDOMElement, i As Long
    If AddToDict Then rsDictAdd AdoRs
    S = uExtractRecords(AdoRs, RecElemName)
    If S = "" Then Exit Sub
    S = uEnclose(S, "Recordset")
    Set xBuf = New ujXML
    xBuf.docLoadXML S
    Set Ndl = xBuf.docRootElem.selectNodes("*")
    For i = 0 To Ndl.Length - 1
        Set El = Ndl.Item(i)
        mElem.appendChild El
    Next i
    Exit Sub
errHandler:
    ErrorIn "ujXML.rsAddAsChildren(AdoRS,RecElemName,AddToDict)", Array(AdoRs, RecElemName, _
         AddToDict)
End Sub
    
Public Sub rsAddAsChildrenExt(ByVal AdoRs As Object, ByVal RecElemName As String, _
            Optional ByVal Specs As String, Optional ByVal AddToDict As Boolean = False)
    On Error GoTo errHandler
    Dim arrF, arrA, arrQ
    If Specs = "" Then Specs = rsGetFieldList(AdoRs)
    uUnpackEQList Specs, arrA, arrF, arrQ
    If AddToDict Then dictAdd Join(arrA, ",")
    With AdoRs
        If .BOF And .EOF Then Exit Sub
        .First
        While Not .EOF
            chCreate RecElemName
            rsElCopyPriv AdoRs, arrA, arrF, arrQ
            navUP
            .MoveNext
        Wend
    End With
    Exit Sub
errHandler:
    ErrorIn "ujXML.rsAddAsChildrenExt(AdoRS,RecElemName,Specs)", Array(AdoRs, RecElemName, _
         Specs)
End Sub
    
Public Sub rsDictAdd(ByVal AdoRs As Object)
    On Error GoTo errHandler
    dictAdd rsGetFieldList(AdoRs)
    Exit Sub
errHandler:
    ErrorIn "ujXML.rsDictAdd(AdoRS)", AdoRs
End Sub
    
Private Function rsGetFieldList(ByVal AdoRs As Object) As String
    On Error GoTo errHandler
    Dim F As Object, Res As String
    For Each F In AdoRs.Fields
        Res = uCombine(Res, ",", F.Name)
    Next F
    rsGetFieldList = Res
    Exit Function
errHandler:
    ErrorIn "ujXML.rsGetFieldList(AdoRs)", AdoRs
End Function

'Controls ====================================================================================================
Public Property Get wfCopyDir() As XENUM_COPY_DIR
    wfCopyDir = mCopyDir
End Property

Public Property Let wfCopyDir(ByVal Value As XENUM_COPY_DIR)
    mCopyDir = Value
End Property

Public Sub wfCopy(ByVal AttrList As String, ByVal arrControls, _
        Optional ByVal ListCopyText As Boolean = False)
    On Error GoTo errHandler
    Dim i As Long, arrA() As String
    If Not IsArray(arrControls) Then arrControls = Array(arrControls)
    arrA = Split(AttrList, ",")
    Check UBound(arrA) = UBound(arrControls), XERR_UJXML, _
        "Count mismatch: AttrList: %1; arrControls: %2.", Array(UBound(arrA) + 1, UBound(arrControls) + 1)
    For i = LBound(arrControls) To UBound(arrControls)
        wfCopyOne arrA(i), arrControls(i), ListCopyText
    Next i
    Exit Sub
errHandler:
    ErrorIn "ujXML.wfCopy(AttrList,arrControls,ListCopyText)", Array(AttrList, arrControls, _
         ListCopyText)
End Sub

Public Sub wfFill(ByVal Ctrl As Object, _
                       ByVal TextAttr As String, Optional ByVal DataAttr As String, _
                       Optional ByVal TblBkmk As String, _
                       Optional ByVal FilterAtrs As String, Optional ByVal FilterValues As String, _
                       Optional ByVal SelDataOrIndexes As String)
    On Error GoTo errHandler
    Dim blnAdd As Boolean, lngD As Long, blnSorted As Boolean, Ind As Long
    blnSorted = uSafeListSorted(Ctrl)
    Check Not blnSorted, XERR_UJXML, "wfFillCtrlCtrl doesn't support SORTED list/combo boxes."
    Ctrl.Clear
    If TblBkmk <> "" Then bkmkGoto TblBkmk
    If navFirstChild Then
        Do
            If FilterAtrs = "" Then
                blnAdd = True
                Else
                blnAdd = elCheckValues(FilterAtrs, FilterValues)
            End If
            If blnAdd Then
                Ctrl.AddItem attrS(TextAttr)
                Ind = Ctrl.ListCount - 1
                If DataAttr = "" Then
                    If SelDataOrIndexes <> "" Then
                        If uInList(SelDataOrIndexes, Ind) Then uListSafeSelect Ctrl, Ind, True
                    End If 'Not IsMissing(...
                    Else
                    lngD = attrI(DataAttr)
                    Ctrl.ITEMDATA(Ctrl.ListCount - 1) = lngD
                    If SelDataOrIndexes <> "" Then
                        If uInList(SelDataOrIndexes, lngD) Then uListSafeSelect Ctrl, Ind, True
                    End If 'Not IsMissing(...
                End If 'DataAttr ....
            End If 'blnAdd ....
        Loop While navNext
    End If
'    If Not IsMissing(SelectedIndex) Then
'        If SelectedIndex <= Ctrl.ListCount - 1 Then Ctrl.ListIndex = SelectedIndex
'    End If
    Exit Sub
errHandler:
    ErrorIn "ujXML.wfFill(Ctrl,TextAttr,DataAttr,TblBkmk,FilterAtrs,FilterValues,SelDataOrIndexes)", _
         Array(Ctrl, TextAttr, DataAttr, TblBkmk, FilterAtrs, FilterValues, SelDataOrIndexes)
End Sub

Public Function wfSyncWith(ByVal Ctrl As Object, Optional ByVal TblBkmk As String, _
        Optional ByVal DataAttr As String) As Boolean
    On Error GoTo errHandler
    Dim Ind As Long, D As Long
    If TblBkmk <> "" Then bkmkGoto TblBkmk
    Ind = Ctrl.ListIndex
    If Ind < 0 Then
        Else
        If DataAttr = "" Then
            navChildByIndex Ind
            wfSyncWith = True
            Else
            wfSyncWith = chLocate(DataAttr, Ctrl.ITEMDATA(Ind))
        End If
    End If
    Exit Function
errHandler:
    ErrorIn "ujXML.wfSyncWith(Ctrl,TblBkmk,DataAttr)", Array(Ctrl, TblBkmk, DataAttr)
End Function

Private Sub wfCopyOne(ByVal AttrName As String, ByVal Ctrl As Object, Optional ByVal ListCopyText As Boolean)
    On Error GoTo errHandler
    Select Case mCopyDir
        Case XDIR_IN: wfCopyIn AttrName, Ctrl, ListCopyText
        Case XDIR_OUT: wfCopyOut AttrName, Ctrl, ListCopyText
        Case XDIR_XMLCLEAR: wfClearXML AttrName, Ctrl, ListCopyText
        Case XDIR_CTRLCLEAR: wfClearCtrl AttrName, Ctrl, ListCopyText
    End Select
    Exit Sub
errHandler:
    ErrorIn "ujXML.wfCopyOne(AttrName,Ctrl,ListCopyText)", Array(AttrName, Ctrl, ListCopyText)
End Sub
    
Private Sub wfCopyIn(ByVal AttrName As String, ByVal Ctrl As Object, Optional ByVal ListCopyText As Boolean)
    On Error GoTo errHandler
    Select Case TypeName(Ctrl)
        Case "Label": attrS(AttrName) = Ctrl.Caption
        Case "TextBox": attrS(AttrName) = Ctrl.Text
        Case "ComboBox"
            Select Case uSafeCtrlStyle(Ctrl)
                Case 0, 1: attrS(AttrName) = Ctrl.Text   'combodropdown, simplecombo
                Case 2 'dropdown list
                        attrV(AttrName) = uListGetValue(Ctrl)
                        If ListCopyText Then attrV("_" & AttrName & "_Text") = Ctrl.Text
            End Select
        Case "ListBox"
            If uSafeListMultiSelect(Ctrl) Then
                attrV(AttrName) = uListGetMultiSelect(Ctrl)
                Else
                attrV(AttrName) = uListGetValue(Ctrl)
                If ListCopyText Then attrV("_" & AttrName & "-Text") = Ctrl.Text
            End If
        Case "OptionButton"
            If Ctrl.Index <> "" Then
                If Ctrl.Value Then attrI(AttrName) = Ctrl.Index
            End If
        Case "CheckBox":  attrB(AttrName) = uValueIn(Ctrl.Value, vbChecked, True)
        Case "FlatScrollBar", "Slider", "ProgressBar", "HScrollBar", "VScrollBar", "ScrollBar"
            attrF(AttrName) = Ctrl.Value
        Case "ToggleButton": attrB(AttrName) = Ctrl.Value
    End Select
    Exit Sub
errHandler:
    ErrorIn "ujXML.wfCopyIn(AttrName,Ctrl,ListCopyText)", Array(AttrName, Ctrl, ListCopyText)
End Sub

Private Sub wfCopyOut(ByVal AttrName As String, ByVal Ctrl As Object, Optional ByVal ListCopyText As Boolean)
    On Error GoTo errHandler
    Select Case TypeName(Ctrl)
        Case "Label":  Ctrl.Caption = attrS(AttrName)
        Case "TextBox": Ctrl.Text = attrS(AttrName)
        Case "ComboBox"
            Select Case uSafeCtrlStyle(Ctrl)
                Case 0, 1: Ctrl.Text = attrS(AttrName)           'combodropdown, simplecombo
                Case 2: uListSetValue Ctrl, attrI(AttrName)      'dropdown list
            End Select
        Case "ListBox"
            If uSafeListMultiSelect(Ctrl) Then
                uListSetMultiSelect Ctrl, attrS(AttrName)
                Else
                uListSetValue Ctrl, attrI(AttrName)
            End If
        Case "OptionButton"
            If Ctrl.Index <> "" Then
                 Ctrl.Value = (Ctrl.Index = attrI(AttrName))
            End If
        Case "CheckBox":        Ctrl.Value = IIf(attrB(AttrName), vbChecked, vbUnchecked)
        Case "FlatScrollBar", "Slider", "ProgressBar", "HScrollBar", "VScrollBar", "ScrollBar"
            Ctrl.Value = attrF(AttrName)
        Case "ToggleButton": Ctrl.Value = attrB(AttrName)
    End Select
    Exit Sub
errHandler:
    ErrorIn "ujXML.wfCopyOut(AttrName,Ctrl,ListCopyText)", Array(AttrName, Ctrl, ListCopyText)
End Sub

Private Sub wfClearCtrl(ByVal AttrName As String, ByVal Ctrl As Object, Optional ByVal ListCopyText As Boolean = False)
    On Error GoTo errHandler
    Dim i As Long
    Select Case TypeName(Ctrl)
        Case "Label":  Ctrl.Caption = ""
        Case "TextBox": Ctrl.Text = ""
        Case "ComboBox"
            Select Case uSafeCtrlStyle(Ctrl)
                Case 0, 1: Ctrl.Text = ""           'combodropdown, simplecombo
                Case 2: uListSetValue Ctrl, -1      'dropdown list
            End Select
        Case "ListBox"
            If uSafeListMultiSelect(Ctrl) Then
                For i = 0 To Ctrl.ListCount - 1
                    Ctrl.Selected(i) = False
                Next i
                'uListSetValue Ctrl, -1
                Else
                uListSetValue Ctrl, -1
            End If
        Case "OptionButton":  Ctrl.Value = False
        Case "CheckBox":        Ctrl.Value = False
        Case "FlatScrollBar", "Slider", "ProgressBar", "HScrollBar", "VScrollBar", "ScrollBar"
            Ctrl.Value = Ctrl.Min
        Case "ToggleButton": Ctrl.Value = False
    End Select
    Exit Sub
errHandler:
    ErrorIn "ujXML.wfClearCtrl(AttrName,Ctrl,ListCopyText)", Array(AttrName, Ctrl, ListCopyText)
End Sub

Private Sub wfClearXML(ByVal AttrName As String, ByVal Ctrl As Object, Optional ByVal ListCopyText As Boolean)
    On Error GoTo errHandler
    Select Case TypeName(Ctrl)
        Case "Label": attrS(AttrName) = Ctrl.Caption
        Case "TextBox": attrS(AttrName) = Ctrl.Text
        Case "ComboBox"
            Select Case uSafeCtrlStyle(Ctrl)
                Case 0, 1: attrS(AttrName) = ""   'combodropdown, simplecombo
                Case 2 'dropdown list
                        attrV(AttrName) = "-1"
                        If ListCopyText Then attrV("_" & AttrName & "_Text") = ""
            End Select
        Case "ListBox"
            If uSafeListMultiSelect(Ctrl) Then
                attrV(AttrName) = ""
                Else
                attrV(AttrName) = -1
                If ListCopyText Then attrV("_" & AttrName & "-Text") = ""
            End If
        Case "OptionButton": attrV(AttrName) = Null
        Case "CheckBox":  attrB(AttrName) = False
        Case "FlatScrollBar", "Slider", "ProgressBar", "HScrollBar", "VScrollBar", "ScrollBar"
            attrF(AttrName) = Ctrl.Min
        Case "ToggleButton": attrB(AttrName) = False
    End Select
    Exit Sub
errHandler:
    ErrorIn "ujXML.wfClearXML(AttrName,Ctrl,ListCopyText)", Array(AttrName, Ctrl, ListCopyText)
End Sub

'IIS =========================================================================================================
Public Property Get iisInIIS() As Boolean
    On Error Resume Next
    iisInIIS = Not (iisGetObject("Application") Is Nothing)
End Property

Private Function iisCheckError() As Boolean
    If Err.Number = 0 Then
        iisCheckError = True
        Else
        miisReport = uCombine(miisReport, vbNewLine & vbNewLine, ErrReport)
    End If
End Function

Public Property Get iisReport() As String
    iisReport = miisReport
End Property

Public Property Get iisReportHTML() As String
    iisReportHTML = "<pre>" & uEscapeHTML(miisReport) & "</pre>"
End Property

Public Sub iisReportClear()
    miisReport = ""
End Sub

Public Function iisSaveDocSafe(ByVal DocName As String) As Boolean
    On Error Resume Next
    iisSaveDoc DocName
    iisSaveDocSafe = iisCheckError
End Function

Public Sub iisSaveDoc(ByVal DocName As String)
    On Error GoTo errHandler
    Dim iisApp As Object
    Set iisApp = iisGetObject("Application")
    Check Not (iisApp Is Nothing), XERR_UJXML, "Failed to retrieve IIS Application object."
    docPersistPrepare
    With iisApp
        .Lock
        Set .Contents(DocName) = mDoc
        .UnLock
    End With
    Exit Sub
errHandler:
    ErrorIn "ujXML.iisSaveDocPriv(DocName)", DocName
End Sub

Public Function iisClearDocSafe(ByVal DocName As String) As Boolean
    On Error Resume Next
    iisClearDoc DocName
    iisClearDocSafe = iisCheckError
End Function

Public Sub iisClearDoc(ByVal DocName As String)
    On Error GoTo errHandler
    Dim iisApp As Object, V
    Set iisApp = iisGetObject("Application")
    Check Not (iisApp Is Nothing), XERR_UJXML, "Failed to retrieve IIS Application object."
    With iisApp
        .Lock
        .Contents(DocName) = V
        .UnLock
    End With
    Exit Sub
errHandler:
    ErrorIn "ujXML.iisClearDocPriv(DocName)", DocName
End Sub

Public Function iisRestoreDocSafe(ByVal DocName As String) As Boolean
    On Error Resume Next
    iisRestoreDoc DocName
    iisRestoreDocSafe = iisCheckError
End Function

Public Sub iisRestoreDoc(ByVal DocName As String)
    On Error GoTo errHandler
    Dim iisApp As Object, Obj As Object
    docOptions(XO_IIS_RETRIEVED) = False
    Set iisApp = iisGetObject("Application")
    Check Not (iisApp Is Nothing), XERR_UJXML, "Failed to retrieve IIS Application object "
    Set Obj = iisGetAppContObjSafe(iisApp, DocName)
    Check Not (Obj Is Nothing), XERR_UJXML, _
            "Failed to retrieve document from Application.Contents; DocName=[%1]", DocName
    Check TypeOf Obj Is IXMLDOMDocument, _
            XERR_UJXML, "Invalid type of retrieved object, expected IXMLDOMDocument, retrieved [%1]", TypeName(Obj)
    Set docObject = Obj
    docPersistRestore
    docOptions(XO_IIS_RETRIEVED) = True
    Exit Sub
errHandler:
    ErrorIn "ujXML.iisRestoreDoc(DocName)", DocName
End Sub

Private Function iisGetAppContObjSafe(ByVal iisApp As Object, ByVal ObjName As String) As Object
    On Error Resume Next
    Set iisGetAppContObjSafe = iisApp.Contents(ObjName)
End Function

Public Function iisRestoreLoadTablesSafe(ByVal DocName As String, ByVal Connect, _
        ByVal TableNames As String, Optional ByVal SQLs, Optional ByVal ExecInBatch As Boolean, _
        Optional ByVal MaxAgeSec As Double = 0) As Boolean
    On Error Resume Next
    iisRestoreLoadTables DocName, Connect, TableNames, SQLs, ExecInBatch, MaxAgeSec
    iisRestoreLoadTablesSafe = iisCheckError
End Function

Public Sub iisRestoreLoadTables(ByVal DocName As String, ByVal Connect, _
        ByVal TableNames As String, Optional ByVal SQLs, Optional ByVal ExecInBatch As Boolean, _
        Optional ByVal MaxAgeSec As Double = 0)
    On Error GoTo errHandler
    Const ONE_SEC As Double = 1# / 86400
    Dim ToReload As Boolean, Age As Double
    ToReload = True
    If iisRestoreDocSafe(DocName) Then
        ToReload = False
        If MaxAgeSec > 0 Then
            navTop
            Age = (Now - attrDt(S_DATETIME)) / ONE_SEC
            If Age > MaxAgeSec Then ToReload = True
        End If 'ReloadIfOlderThanSec...
    End If
    docOptions(XO_IIS_RETRIEVED) = Not ToReload
    If ToReload Then
        If mDocProgID = "" Then mDocProgID = PROGID_FREETHREADEDDOC ' "MSXML2.FreeThreadedDOMDocument"
        docInit "Tables"
        tblLoadTables Connect, TableNames, SQLs, ExecInBatch
        navTop
        attrDt(S_DATETIME) = Now
        iisSaveDoc DocName
    End If
    Exit Sub
errHandler:
    ErrorIn "ujXML.iisRestoreLoadTables(DocName,Connect,TableNames,SQLs,ExecInBatch,MaxAgeSec)", _
         Array(DocName, Connect, TableNames, SQLs, ExecInBatch, MaxAgeSec)
End Sub

Public Function iisCopySafe(Optional ByVal AttrList As String) As Boolean
    On Error Resume Next
    iisCopy AttrList
    iisCopySafe = iisCheckError
End Function

Public Sub iisCopy(Optional ByVal AttrList As String)
    On Error GoTo errHandler
    If iisPosted Then
        iisCopyForm AttrList
        Else
        iisCopyQuery AttrList
    End If
    Exit Sub
errHandler:
    ErrorIn "ujXML.iisCopy(AttrList)", AttrList
End Sub
   
Public Function iisCopyFormSafe(Optional ByVal AttrList As String) As Boolean
    On Error Resume Next
    iisCopyForm AttrList
    iisCopyFormSafe = iisCheckError
End Function
   
Public Sub iisCopyForm(Optional ByVal AttrList As String)
    On Error GoTo errHandler
    Dim arrA, arrF, i As Long, Req As Object
    If AttrList = "" Then AttrList = iisFormFields
    uUnpackEQList AttrList, arrA, arrF
    Set Req = iisGetObject("Request")
    For i = 0 To UBound(arrA)
        attrS(arrA(i)) = Req.Form(arrF(i))
    Next i
    Exit Sub
errHandler:
    ErrorIn "ujXML.iisCopyForm(AttrList)", AttrList
End Sub
   
Public Function iisCopyQuerySafe(Optional AttrList As String) As Boolean
    On Error Resume Next
    iisCopyQuery AttrList
    iisCopyQuerySafe = iisCheckError
End Function
   
Public Sub iisCopyQuery(Optional ByVal AttrList As String)
    On Error GoTo errHandler
    Dim arrA, arrF, i As Long, Req As Object
    If AttrList = "" Then AttrList = iisQueryParams
    uUnpackEQList AttrList, arrA, arrF
    Set Req = iisGetObject("Request")
    For i = 0 To UBound(arrA)
        attrS(arrA(i)) = Req.queryString.Item(arrF(i)) '.QueryString.Item(arrF(i))
    Next i
    Exit Sub
errHandler:
    ErrorIn "ujXML.iisCopyQuery(AttrList)", AttrList
End Sub
   
Public Function iisPosted() As Boolean
    On Error Resume Next
    iisPosted = (UCase$(iisRequestMethod) = "POST")
End Function

Public Function iisQueryParamsSafe() As String
    On Error Resume Next
    iisQueryParamsSafe = iisQueryParams
    iisCheckError
End Function
   
Public Function iisQueryParams() As String
    On Error GoTo errHandler
    Dim S As String, i As Long, Req As Object 'As ASPTypeLibrary.Request
    Set Req = iisGetObject("Request")
    For i = 1 To Req.queryString.Count
        S = uCombine(S, ",", Req.queryString.Key(i))
    Next i
    iisQueryParams = S
    Exit Function
errHandler:
    ErrorIn "ujXML.iisQueryParams"
End Function
   
Public Function iisFormFieldsSafe() As String
    On Error Resume Next
    iisFormFieldsSafe = iisFormFields
    iisCheckError
End Function
   
Public Function iisFormFields() As String
    On Error GoTo errHandler
    Dim S As String, i As Long, Req As Object
    Set Req = iisGetObject("Request")
    For i = 1 To Req.Form.Count
        S = uCombine(S, ",", Req.Form.Key(i))
    Next i
    iisFormFields = S
    Exit Function
errHandler:
    ErrorIn "ujXML.iisFormFields"
End Function
   
Public Function iisRequestMethodSafe() As String
    On Error Resume Next
    iisRequestMethodSafe = iisRequestMethod
    iisCheckError
End Function
   
Public Function iisRequestMethod() As String
    On Error GoTo errHandler
    Dim objReq As Object 'ASPTypeLibrary.Request
    Set objReq = iisGetObject("Request")
    Check Not objReq Is Nothing, XERR_UJXML, "Failed to retrieve IIS Request object."
    iisRequestMethod = objReq.ServerVariables("REQUEST_METHOD")
    Exit Function
errHandler:
    ErrorIn "ujXML.iisRequestMethod"
End Function

Public Function iisWriteSafe(ByVal HtmlText As String) As Boolean
    On Error Resume Next
    iisWrite HtmlText
    iisWriteSafe = iisCheckError
End Function

Public Sub iisWrite(ByVal HtmlText As String)
    On Error GoTo errHandler
    Dim objResp As Object 'ASPTypeLibrary.Request
    Set objResp = iisGetObject("Response")
    Check Not objResp Is Nothing, XERR_UJXML, "Failed to retrieve IIS Response object."
    objResp.Write HtmlText
    Exit Sub
errHandler:
    ErrorIn "ujXML.iisWrite(HtmlText)", "..."
End Sub

Public Function iisReportWriteSafe() As Boolean
    On Error Resume Next
    iisReportWrite
    iisReportWriteSafe = iisCheckError
End Function

Public Sub iisReportWrite()
    On Error GoTo errHandler
    iisWrite iisReportHTML
    Exit Sub
errHandler:
    ErrorIn "ujXML.iisReportWrite"
End Sub

Public Function iisGetObjectSafe(ByVal ObjectName As String) As Object
    On Error Resume Next
    Set iisGetObjectSafe = iisGetObject(ObjectName)
    iisCheckError
End Function
   
Public Function iisGetObject(ByVal ObjectName As String) As Object
    On Error GoTo errHandler
    Dim Ctx As ObjectContext
    Set Ctx = GetContext
    Set iisGetObject = Ctx(ObjectName)
    Exit Function
errHandler:
    ErrorIn "ujXML.iisGetObject"
End Function

Private Function GetContext() As ObjectContext
    On Error Resume Next
    Set GetContext = GetObjectContext
End Function

Private Sub CheckSetSysHandler()
    On Error GoTo errHandler
    If (Not docOptions(XO_NO_SYS_HANDLER)) And (Not ErrSysHandlerWasSet) Then
        ErrSysHandlerSet
    End If
    Exit Sub
errHandler:
    ErrorIn "ujXML.CheckSetSysHandler"
End Sub

'Version group ==============================================================
Public Property Get verMajor() As Long
    verMajor = App.Major
End Property

Public Property Get verMinor() As Long
    verMinor = App.Minor
End Property

Public Property Get verRevision() As Long
    verRevision = App.Revision
End Property

Public Function verVersion() As String
    verVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

'Utilities ================================================================
Public Function uEscapeXML(ByVal S As String) As String
    S = Replace(S, "&", "&amp;")
    S = Replace(S, "<", "&lt;")
    S = Replace(S, ">", "&gt;")
    S = Replace(S, """", "&quot;")
    S = Replace(S, "'", "&apos;")
    uEscapeXML = S
End Function

Public Function uEscapeHTML(ByVal S As String) As String
    S = Replace(S, "&", "&amp;")
    S = Replace(S, "<", "&lt;")
    S = Replace(S, ">", "&gt;")
    S = Replace(S, """", "&quot;")
    uEscapeHTML = S
End Function


'Utilities, Private
Private Function uListGetValue(ByVal Ctrl As Object)
    On Error GoTo errHandler
    Dim i As Long
    With Ctrl
        If uListHasData(Ctrl) Then
            If .ListIndex = -1 Then uListGetValue = 0 Else uListGetValue = .ITEMDATA(.ListIndex)
            Else
            uListGetValue = .ListIndex
        End If
    End With
    Exit Function
errHandler:
    ErrorIn "ujXML.uListGetValue(Ctrl)", Ctrl
End Function

Private Sub uListSetValue(ByVal Ctrl As Object, ByVal Value As Long)
    On Error GoTo errHandler
    Dim i As Long
    With Ctrl
        If uListHasData(Ctrl) Then
            For i = 0 To .ListCount - 1
                If .ITEMDATA(i) = Value Then .ListIndex = i: Exit Sub
            Next i
            .ListIndex = -1
            Else
            .ListIndex = Value
        End If
    End With
    Exit Sub
errHandler:
    ErrorIn "ujXML.uListSetValue(Ctrl,Value)", Array(Ctrl, Value)
End Sub

Private Sub uListSafeSelect(ByVal Ctrl As Object, ByVal Ind As Long, ByVal Sel As Boolean)
    On Error Resume Next
    Ctrl.Selected(Ind) = Sel
    If Err.Number <> 0 And Sel Then Ctrl.ListIndex = Ind
    Err.Clear
End Sub

Private Function uListHasData(ByVal Ctrl As Object) As Boolean
    On Error GoTo errHandler
    Dim i As Long
    uListHasData = True
    With Ctrl
        For i = 0 To .ListCount - 1
            If .ITEMDATA(i) <> 0 Then Exit Function
        Next i
    End With
    uListHasData = False
    Exit Function
errHandler:
    uListHasData = False
    'ErrorIn "ujXML.uListHasData(Ctrl)", Ctrl
End Function

Private Function uListGetMultiSelect(ByVal Ctrl As Object) As String
    On Error GoTo errHandler
    Dim i As Long, S As String, blnHasData As Boolean
    blnHasData = uListHasData(Ctrl)
    With Ctrl
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                If S <> "" Then S = S & ","
                If blnHasData Then S = S & .ITEMDATA(i) Else S = S & i
            End If
        Next i
    End With
    uListGetMultiSelect = S
    Exit Function
errHandler:
    ErrorIn "ujXML.uListGetMultiSelect(Ctrl)", Ctrl
End Function

Private Sub uListSetMultiSelect(ByVal Ctrl As Object, ByVal DataList As String)
    On Error GoTo errHandler
    Dim i As Long, V As Long, blnHasData As Boolean
    blnHasData = uListHasData(Ctrl)
    With Ctrl
        For i = 0 To .ListCount - 1
            If blnHasData Then V = .ITEMDATA(i) Else V = i
            .Selected(i) = uInList(DataList, V)
        Next i
    End With
    Exit Sub
errHandler:
    ErrorIn "ujXML.uListSetMultiSelect(Ctrl,DataList)", Array(Ctrl, DataList)
End Sub

Private Function uListAdd(ByVal sList As String, ByVal sValue As String, Optional ByVal Delim As String = ",") As String
    On Error GoTo errHandler
    If Not uInList(sList, sValue, Delim) Then sList = uCombine(sList, Delim, sValue)
    uListAdd = sList
    Exit Function
errHandler:
    ErrorIn "modXUtils.uListAdd(sList,sValue)", Array(sList, sValue)
End Function

Private Function uListRemove(ByVal sList As String, ByVal sValue As String) As String
    On Error GoTo errHandler
    Dim sL As String, sV As String, P As Long
    sL = "," & sList & ","
    sV = "," & sValue & ","
    sL = Replace(sL, sV, ",")
    uListRemove = MID$(sL, 2, Len(sL) - 2) 'cut-off commas
    Exit Function
errHandler:
    ErrorIn "modXUtils.uListRemove(sList,sValue)", Array(sList, sValue)
End Function

Private Function uListRemoveLast(ByVal sList As String) As String
    On Error GoTo errHandler
    Dim P As Long
    P = InStrRev(sList, ",")
    If P = 0 Then sList = "" Else sList = Left$(sList, P - 1)
    uListRemoveLast = sList
    Exit Function
errHandler:
    ErrorIn "modXUtils.uListRemoveLast(sList)", sList
End Function

Private Function uInList(ByVal sList As String, ByVal sValue As String, Optional ByVal Delim As String = ",") As Boolean
    Dim sL As String, sV As String
    sL = Delim & sList & Delim
    sV = Delim & sValue & Delim
    uInList = (InStr(1, sL, sV) > 0)
End Function

Private Function uInsInto(ByVal IntoStr As String, ByVal Values, _
        Optional XMLEscape As Boolean = False) As String
    On Error GoTo errHandler
    Dim i As Long, ChrX As String, strV As String
    ChrX = Chr$(vbKeyBack) 'Spec char to act instead of % during manipulations
    IntoStr = Replace(IntoStr, "%", ChrX)
    If Not IsArray(Values) Then Values = Array(Values)
    For i = LBound(Values) To UBound(Values)
        strV = uSafeStr(Values(i))
        If XMLEscape Then strV = uEscapeXML(strV)  'Replace(strV, "'", "''")  '
        IntoStr = Replace(IntoStr, ChrX & (i + 1), strV)
    Next i
    uInsInto = Replace(IntoStr, ChrX, "%") 'replace back
    Exit Function
errHandler:
    ErrorIn "ujXML.uInsInto(IntoStr,Values)", Array(IntoStr, Values)
End Function


Private Function uDBValue(Value)
  uDBValue = IIf(IsEmpty(Value), Null, Value)
End Function

Private Function uVBValue(Value)
    Dim V
    uVBValue = IIf(IsNull(Value), V, Value)
End Function

Private Function uGetField(ByVal AdoRs As Object, ByVal FldName As String)
    On Error GoTo errHandler
    Dim F As Object
    Set F = AdoRs.Fields(FldName)
    If Not IsNull(F.Value) Then uGetField = F.Value
    Exit Function
errHandler:
    ErrorIn "ujXML.uGetField(AdoRS,FldName)", Array(AdoRs, FldName)
End Function

Private Function uSafeGetFieldObj(ByVal AdoRs As Object, ByVal FldName As String) As Object
    On Error Resume Next
    Set uSafeGetFieldObj = AdoRs.Fields(FldName)
End Function


Private Function uIsElement(ByVal Node As msxml2.IXMLDOMNode) As Boolean
    On Error GoTo errHandler
    uIsElement = (Node.nodeType = NODE_ELEMENT)
    Exit Function
errHandler:
    ErrorIn "ujXML.uIsElement(Node)", Node
End Function

Private Function uIsAttribute(ByVal Node As msxml2.IXMLDOMNode) As Boolean
    On Error GoTo errHandler
    uIsAttribute = (Node.nodeType = NODE_ATTRIBUTE)
    Exit Function
errHandler:
    ErrorIn "ujXML.uIsAttribute(Node)", Node
End Function

Private Function uCompare(ByVal V1, ByVal V2) As Long
    On Error GoTo errHandler
    V1 = uToEmpty(V1)
    V2 = uToEmpty(V2)
    If IsEmpty(V1) Then
        uCompare = IIf(IsEmpty(V2), 0, -1)
    ElseIf IsEmpty(V2) Then
        uCompare = IIf(IsEmpty(V1), 0, 1)
    Else
        If IsNumeric(V1) And IsNumeric(V2) Then
            V1 = uSafeDbl(V1)
            V2 = uSafeDbl(V2)
        ElseIf IsDate(V1) And IsDate(V2) Then
            V1 = uSafeDate(V1)
            V2 = uSafeDate(V2)
        End If
        If V1 = V2 Then
            uCompare = 0
        ElseIf V1 < V2 Then
            uCompare = -1
        Else
            uCompare = 1
        End If
    End If
    Exit Function
errHandler:
    ErrorIn "ujXML.uCompare(V1,V2)", Array(V1, V2)
End Function

Private Property Get uEmpty() As Variant
    Dim V
    uEmpty = V
End Property

Private Function uToEmpty(ByVal Value)
    uToEmpty = Value
    If IsNull(Value) Or IsEmpty(Value) Then
        uToEmpty = uEmpty
    ElseIf VarType(Value) = vbString Then
        If Value = "" Then uToEmpty = uEmpty
    End If
End Function

Private Function uSafeListSorted(ByVal C As Object) As Boolean
    On Error Resume Next
    uSafeListSorted = C.Sorted
End Function

Private Function uSafeListMultiSelect(ByVal C As Object) As Boolean
    On Error Resume Next
    uSafeListMultiSelect = ((C.Style <> 0) Or (C.MultiSelect <> 0))
    If Err.Number <> 0 Then
        uSafeListMultiSelect = ((C.ListStyle <> 0) Or (C.MultiSelect <> 0))
    End If
End Function

Private Function uSafeCtrlStyle(ByVal C As Object) As Long
    On Error Resume Next
    uSafeCtrlStyle = C.Style
End Function

Private Property Get uChrBk() As String
    uChrBk = Chr$(vbKeyBack)
End Property

Private Function uUnpackEQList(ByVal ListAeqB As String, ByRef arrA, ByRef arrB, Optional ByRef arrQ) As Long
    On Error GoTo errHandler
    Dim arrAB() As String, arrT() As String, i As Long, sA As String, SB As String, Ubnd As Long
    Dim arr
    arrAB = Split(ListAeqB, ",")
    Ubnd = UBound(arrAB)
    uUnpackEQList = Ubnd
    arrA = uCreateArr(Ubnd)
    arrB = uCreateArr(Ubnd)
    If Not IsMissing(arrQ) Then arrQ = uCreateArr(Ubnd)
    For i = 0 To Ubnd
        If Not IsMissing(arrQ) Then
            arrQ(i) = (InStr(arrAB(i), "'") > 0)
            If arrQ(i) Then arrAB(i) = Replace(arrAB(i), "'", "")
        End If
        arrT = Split(arrAB(i), "=")
        arrA(i) = arrT(0)
        If UBound(arrT) > 0 Then arrB(i) = arrT(1) Else arrB(i) = arrT(0)
    Next i
    Exit Function
errHandler:
    ErrorIn "modXUtils.uUnpackEQList(ListAeqB,arrA,arrB,arrQ)", Array(ListAeqB, arrA, arrB, arrQ)
End Function


Private Function uCreateArr(ByVal ArrUbnd As Long)
    Dim arr()
    ReDim arr(0 To ArrUbnd)
    uCreateArr = arr
End Function

Private Function uSafeFormat(ByVal V, ByVal Fmt As String) As String
    On Error Resume Next
    uSafeFormat = Format(V, Fmt)
End Function

Private Function uCreateEQList(ByRef arrN, ByRef arrV, Optional ByVal Delim = ", ") As String
    On Error GoTo errHandler
    Dim arrRes() As String, i As Long
    ReDim arrRes(0 To UBound(arrN))
    For i = 0 To UBound(arrN)
        arrRes(i) = arrN(i) & " = " & arrV(i)
    Next i
    uCreateEQList = Join(arrRes, Delim)
    Exit Function
errHandler:
    ErrorIn "ujXML.uCreateEQList(arrN,arrV,Delim)", Array(arrN, arrV, Delim)
End Function

'Recordset =====================================================================================
Private Function uSaveRecordset(ByVal RS As Object) As String
    On Error GoTo errHandler
    Dim S As Object
    Set S = CreateObject("ADODB.Stream")
    S.Open
    RS.Save S, 1 'adPersistXML
    S.Position = 0
    uSaveRecordset = S.ReadText
    Exit Function
errHandler:
    ErrorIn "modXUtils.uSaveRecordset(RS)", RS
End Function

'Extracts z:row elements, changes [z:row] to ZRowElem
Private Function uExtractRecords(ByVal RS As Object, ByVal ZRowElem As String, _
            Optional ByVal RsElem As String = "") As String
    On Error GoTo errHandler
    Dim P1 As Long, P2 As Long, xml As String
    xml = uSaveRecordset(RS)
    P1 = InStr(1, xml, "<rs:data>")
    P2 = InStr(1, xml, "</rs:data>")
    If P2 > P1 And P1 > 0 Then
        P1 = P1 + Len("<rs:data>")
        xml = MID$(xml, P1, P2 - P1)
        xml = uChangeElemName(xml, "z:row", ZRowElem)
        xml = Replace(xml, "<rs:insert>", "")
        xml = Replace(xml, "</rs:insert>", "")
        Else
        xml = ""
    End If
    If RsElem <> "" Then xml = uEnclose(xml, RsElem)
    uExtractRecords = xml
    Exit Function
errHandler:
    Err.Raise Err.Number, Err.Source, "(uExtractRecords) " & Err.Description
End Function

'We can use string matching, because patterns include special symbols that
'are always escaped in XML attribute values
Private Function uChangeElemName(ByVal xml As String, _
                                ByVal OldName As String, _
                                ByVal NewName As String) As String
    xml = Replace(xml, "<" & OldName & " ", "<" & NewName & " ")
    xml = Replace(xml, "<" & OldName & ">", "<" & NewName & ">")
    xml = Replace(xml, "<" & OldName & "/>", "<" & NewName & "/>")
    xml = Replace(xml, "</" & OldName & ">", "</" & NewName & ">")
    uChangeElemName = xml
End Function

Private Function uMakeElement(ByVal Elem As String, ByVal attrS As String, ByVal Content As String) As String
    uMakeElement = "<" & Elem & attrS & IIf(Content = "", "/>", ">" & Content & "</" & Elem & ">") & vbNewLine
End Function

Private Function uMakeAttr(ByVal Attr As String, ByVal Value As String) As String
    uMakeAttr = " " & Attr & "='" & uEscapeXML(Value) & "'"
End Function

Private Function uEnclose(ByVal xmlSrc As String, ByVal ElemOrTag As String, _
        Optional ByVal ClosingTag As String) As String
    Dim xmlRes As String
    If ElemOrTag = "" Then
        xmlRes = xmlSrc
        Else
        If InStr(1, ElemOrTag, "<") = 0 Then
            xmlRes = "<" & ElemOrTag & ">" & xmlSrc
            Else
            xmlRes = ElemOrTag & xmlSrc
        End If
        If ClosingTag = "" Then
            xmlRes = xmlRes & "</" & ElemOrTag & ">"
            Else
            xmlRes = xmlRes & ClosingTag
        End If
    End If
    uEnclose = xmlRes
End Function

'=========================================================================
Private Function uSafeStr(ByVal V, Optional ByVal Default) As String
  On Error Resume Next
  uSafeStr = IIf(IsMissing(Default), "", Default)
  uSafeStr = Trim$(V)
End Function

Private Function uSafeLng(ByVal V, Optional ByVal Default = 0) As Long
  On Error Resume Next
  uSafeLng = IIf(IsMissing(Default), 0, CLng(Default))
  If IsObject(V) Then uSafeLng = V.Value Else uSafeLng = CLng(V)
  End Function

Private Function uSafeDbl(ByVal V, Optional ByVal Default = 0#) As Double
  On Error Resume Next
  uSafeDbl = IIf(IsMissing(Default), 0#, Default)
  uSafeDbl = CDbl(V)
  End Function

Private Function uSafeDate(ByVal V, Optional ByVal Default) As Date
    On Error Resume Next
    uSafeDate = IIf(IsMissing(Default), CDate(0), Default)
    uSafeDate = CDate(V)
    If Err.Number <> 0 And VarType(V) = vbString Then
        If V Like "*T??:??:??" Then uSafeDate = CDate(Replace(V, "T", " "))
    End If
  End Function

Private Function uSafeBool(ByVal V) As Boolean
    On Error Resume Next
    uSafeBool = CBool(V)
End Function

Function uCombine(ByVal St1 As String, Delim As String, ByVal St2 As String) As String
    If St2 = "" Then
        uCombine = St1
    ElseIf St1 = "" Then
        uCombine = St2
    Else
        uCombine = St1 & Delim & St2
    End If
End Function

Private Function uRemoveObject(ByVal col As Collection, ByVal Obj As Object) As Boolean
    On Error GoTo errHandler
  Dim colObj As Object, Index As Long
  uRemoveObject = True
  For Index = 1 To col.Count
    If IsObject(col(Index)) Then
      Set colObj = col(Index)
      If colObj Is Obj Then
        col.Remove Index
        Exit Function
        End If
      End If ' IsObject...
    Next Index
  uRemoveObject = False
    Exit Function
errHandler:
    ErrorIn "modXUtils.uRemoveObject(col,Obj)", Array(col, Obj)
End Function
  
Private Function uValueIn(ByVal Value, ParamArray Values()) As Boolean
  Dim i As Integer
  uValueIn = True
  For i = LBound(Values) To UBound(Values)
    If Value = Values(i) Then Exit Function
    Next i
  uValueIn = False
  End Function

Private Function uQuote(ByVal S As String) As String
    uQuote = "'" & Replace(S, "'", "''") & "'"
End Function

Private Function uDQuote(ByVal S As String) As String
    uDQuote = """" & Replace(S, """", """""") & """"
End Function

Private Function uSmartQuote(ByVal S As String) As String
    If InStr(1, S, "'") = 0 Then
        uSmartQuote = uQuote(S)
        Else
        uSmartQuote = uDQuote(S)
    End If
End Function

Private Function uIndexOf(ByVal SubStr As String, ByVal strList As String, _
        Optional ByVal Delim As String = ",") As Long
    On Error GoTo errHandler
    Dim arrList() As String, i As Long
    uIndexOf = -1
    arrList = Split(strList, Delim)
    For i = 0 To UBound(arrList)
        uIndexOf = i + 1
        If LCase$(arrList(i)) = LCase(SubStr) Then
            uIndexOf = i
            Exit Function
        End If
    Next i
    Exit Function
errHandler:
    ErrorIn "ujXML.uIndexOf(SubStr,strList,Delim)", Array(SubStr, strList, Delim)
End Function

Private Function uSubstrByIndex(ByVal strList As String, ByVal Index As Long, _
            Optional ByVal Delim As String = ",") As String
    On Error GoTo errHandler
    Dim arrList() As String
    arrList = Split(strList, Delim)
    Check uInRange(Index, LBound(arrList), UBound(arrList)), XERR_UJXML, "Substring index (%1) out of range.", Index
    uSubstrByIndex = arrList(Index)
    Exit Function
errHandler:
    ErrorIn "ujXML.uSubstrByIndex(strList,Index,Delim)", Array(strList, Index, Delim)
End Function

Private Function uInRange(ByVal V, ByVal V1, ByVal V2) As Boolean
  On Error Resume Next
  uInRange = (V >= V1) And (V <= V2)
End Function

Private Function uIsChildOf(ByVal ChEl As IXMLDOMElement, ByVal ParEl As IXMLDOMElement) As Boolean
    On Error GoTo errHandler
    Dim ChNd As IXMLDOMNode, ParNd As IXMLDOMNode
    Set ChNd = ChEl
    Set ParNd = ParEl
    Do
        If ChNd Is ParNd Then
            uIsChildOf = True
            Exit Function
        End If
        Set ChNd = ChNd.parentNode
    Loop Until ChNd Is Nothing
    Exit Function
errHandler:
    ErrorIn "ujXML.uIsChildOf(ChEl,ParEl)", Array(ChEl, ParEl)
End Function

'Fix List
'10/05/03  Method elGetAttrList: fixed search criteria


