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
Const strHead = "<?xml version=""1.0"" encoding=""UTF-8""?>" _
            & "<PO_1 xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:noNamespaceSchemaLocation=""C:\PBKS\Templates\PO_1.xsd"">"

Public Enum XENUM_
    X_UJXML = MAP_RESERVED_FIRST + 20
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
Const S_DEFAULT_PROGID = "MSXML2.DOMDocument.6.0"
Const PROGID_FREETHREADEDDOC = "msxml2.FreeThreadedDOMDocument"

Private mDoc As IXMLDOMDocument
Private mTopElem As IXMLDOMElement
Private mElem As IXMLDOMElement

Private mDocProgID As String
Private mOptions As Long
Private mCopyDir As XENUM_COPY_DIR
Private mDocModified As Boolean

'Private m as msxml2.FreeThreadedDOMDocument

Private mDict As String, mBkmkList As String
Private mhistColl As Collection, mbkmkColl As Collection
Private mSavedEl As MSXML2.IXMLDOMElement
Private miisReport As String

'Transactional stuff
Private mtranSavedXML As String



'Element =============================================================================================================
Public Property Get attrV(ByVal AttrName As String) As Variant
8790      On Error GoTo errHandler
          Dim V, vEmpty
8800      If docOptions(XO_USE_DICT) Then dictCheckName AttrName
8810      V = mElem.getAttribute(AttrName)
8820      attrV = IIf(IsNull(V), vEmpty, V)
8830      Exit Property
errHandler:
8840      ErrorIn "ujXML.attrV(AttrName)", AttrName
End Property

Public Property Let attrV(ByVal AttrName As String, ByVal AValue As Variant)
8850      On Error GoTo errHandler
          Dim OldValue
8860      If docOptions(XO_USE_DICT) Then dictCheckName AttrName
8870      With mElem
8880          OldValue = .getAttribute(AttrName)
8890          If IsNull(AValue) Or IsEmpty(AValue) Then
8900              .removeAttribute AttrName
8910              Else
8920              .setAttribute AttrName, AValue
8930          End If
8940          If docOptions(XO_TRACK_CHANGES) And Left$(AttrName, 1) <> "_" Then
                  Dim NewV
8950              NewV = .getAttribute(AttrName)
8960              If uCompare(OldValue, NewV) Then elSetModified True
8970          End If
8980      End With
8990      Exit Property
errHandler:
9000      ErrorIn "ujXML.attrV(AttrName,AValue)", Array(AttrName, AValue)
End Property

Public Property Get attrV4DB(ByVal AttrName As String) As Variant
9010      On Error GoTo errHandler
          Dim V
9020      V = attrV(AttrName)
9030      On Error GoTo errHandler
9040      If IsEmpty(V) Then
9050          attrV4DB = Null
9060          Else
9070          attrV4DB = V
9080      End If
9090      Exit Property
errHandler:
9100      ErrorIn "ujXML.attrV4DB(AttrName)", AttrName
End Property


Public Property Get attrS(ByVal AttrName As String) As String
9110      On Error GoTo errHandler
9120      attrS = uSafeStr(attrV(AttrName))
9130      Exit Property
errHandler:
9140      ErrorIn "ujXML.attrS(AttrName)", AttrName
End Property

Public Property Let attrS(ByVal AttrName As String, ByVal Value As String)
9150      On Error GoTo errHandler
9160      attrV(AttrName) = Value
9170      Exit Property
errHandler:
9180      ErrorIn "ujXML.attrS(AttrName,Value)", Array(AttrName, Value)
End Property

Public Property Get attrI(ByVal AttrName As String) As Long
9190      On Error GoTo errHandler
9200      attrI = uSafeLng(attrV(AttrName))
9210      Exit Property
errHandler:
9220      ErrorIn "ujXML.attrI(AttrName)", AttrName
End Property

Public Property Let attrI(ByVal AttrName As String, ByVal Value As Long)
9230      On Error GoTo errHandler
9240      attrV(AttrName) = Value
9250      Exit Property
errHandler:
9260      ErrorIn "ujXML.attrI(AttrName,Value)", Array(AttrName, Value)
End Property

Public Property Get attrB(ByVal AttrName As String) As Boolean
9270      On Error GoTo errHandler
9280      attrB = (attrV(AttrName) = S_TRUE)
9290      Exit Property
errHandler:
9300      ErrorIn "ujXML.attrB(AttrName)", AttrName
End Property

Public Property Let attrB(ByVal AttrName As String, ByVal Value As Boolean)
9310      On Error GoTo errHandler
9320      attrV(AttrName) = IIf(Value, S_TRUE, S_FALSE)
9330      Exit Property
errHandler:
9340      ErrorIn "ujXML.attrB(AttrName,Value)", Array(AttrName, Value)
End Property

Public Property Get attrF(ByVal AttrName As String) As Double
9350      On Error GoTo errHandler
9360      attrF = uSafeDbl(attrV(AttrName))
9370      Exit Property
errHandler:
9380      ErrorIn "ujXML.attrF(AttrName)", AttrName
End Property

Public Property Let attrF(ByVal AttrName As String, ByVal Value As Double)
9390      On Error GoTo errHandler
9400      attrV(AttrName) = Value
9410      Exit Property
errHandler:
9420      ErrorIn "ujXML.attrF(AttrName,Value)", Array(AttrName, Value)
End Property

Public Property Get attrDt(ByVal AttrName As String) As Date
9430      On Error GoTo errHandler
9440      attrDt = uSafeDate(attrV(AttrName))
9450      Exit Property
errHandler:
9460      ErrorIn "ujXML.attrDt(AttrName)", AttrName
End Property

Public Property Let attrDt(ByVal AttrName As String, ByVal Value As Date)
9470      On Error GoTo errHandler
9480      attrV(AttrName) = Value
9490      Exit Property
errHandler:
9500      ErrorIn "ujXML.attrDt(AttrName,Value)", Array(AttrName, Value)
End Property

Public Property Get attrEscHtm(ByVal AttrName As String) As String
9510      On Error GoTo errHandler
9520      attrEscHtm = uEscapeHTML(attrS(AttrName))
9530      Exit Property
errHandler:
9540      ErrorIn "ujXML.attrEscHtm(AttrName)", AttrName
End Property

Public Property Get attrEscXml(ByVal AttrName As String) As String
9550      On Error GoTo errHandler
9560      attrEscXml = uEscapeXML(attrS(AttrName))
9570      Exit Property
errHandler:
9580      ErrorIn "ujXML.attrEscXml(AttrName)", AttrName
End Property


Public Property Get attrQt(ByVal AttrName As String) As String
9590      On Error GoTo errHandler
9600      attrQt = uQuote(attrS(AttrName))
9610      Exit Property
errHandler:
9620      ErrorIn "ujXML.attrQt(AttrName)", AttrName
End Property

Public Property Get attrDqt(ByVal AttrName As String) As String
9630      On Error GoTo errHandler
9640      attrDqt = uDQuote(attrS(AttrName))
9650      Exit Property
errHandler:
9660      ErrorIn "ujXML.attrDqt(AttrName)", AttrName
End Property

Public Property Get attrSmartQt(ByVal AttrName As String) As String
9670      On Error GoTo errHandler
9680      attrSmartQt = uSmartQuote(attrS(AttrName))
9690      Exit Property
errHandler:
9700      ErrorIn "ujXML.attrSmartQt(AttrName)", AttrName
End Property

'-----------
Public Property Get attrIsNull(ByVal AttrName As String) As Boolean
9710      On Error GoTo errHandler
9720      attrIsNull = (mElem.getAttributeNode(AttrName) Is Nothing)
9730      Exit Property
errHandler:
9740      ErrorIn "ujXML.attrIsNull(AttrName)", AttrName
End Property

Public Property Get attrNotEmpty(ByVal AttrName As String) As Boolean
9750      On Error GoTo errHandler
9760      attrNotEmpty = (Trim$(attrS(AttrName)) <> "")
9770      Exit Property
errHandler:
9780      ErrorIn "ujXML.attrNotEmpty(AttrName)", AttrName
End Property

Public Property Get attrNZ(ByVal AttrName As String) As Boolean
9790      On Error GoTo errHandler
9800      attrNZ = (attrF(AttrName) <> 0)
9810      Exit Property
errHandler:
9820      ErrorIn "ujXML.attrNZ(AttrName)", AttrName
End Property

'Doesn't change mElemStatus!!
Public Sub attrSet(ByVal AttrName As String, ByVal Value)
9830      On Error GoTo errHandler
9840      mElem.setAttribute AttrName, Value
9850      Exit Sub
errHandler:
9860      ErrorIn "ujXML.attrSet(AttrName,Value)", Array(AttrName, Value)
End Sub

Public Sub elSetModified(ByVal AModified As Boolean)
9870      On Error GoTo errHandler
9880      Select Case elStatus
              Case XES_NEW: 'Nothing to do
9890          Case XES_EXISTING: If AModified Then mElem.setAttribute S_STATUS, "M"
9900          Case XES_MODIFIED: If Not AModified Then mElem.removeAttribute S_STATUS
9910          Case XES_DELETED: 'nothing to do
9920          Case XES_JOINED:  'nothing to do
9930      End Select
9940      If AModified Then mDocModified = True
9950      Exit Sub
errHandler:
9960      ErrorIn "ujXML.elSetModified(Modified)", AModified
End Sub

Public Sub attrSetIfEmpty(ByVal AttrName As String, ByVal Value)
9970      On Error GoTo errHandler
9980      If Not attrNotEmpty(AttrName) Then attrV(AttrName) = Value
9990      Exit Sub
errHandler:
10000     ErrorIn "ujXML.attrSetIfEmpty(AttrName,Value)", Array(AttrName, Value)
End Sub

Public Sub attrSetAsLkpFrom(ByVal AttrName As String, IndexAttr As String, _
        ByVal FromList As String, Optional ByVal Delim As String = ",")
10010     On Error GoTo errHandler
10020     attrS(AttrName) = uSubstrByIndex(FromList, attrI(IndexAttr), Delim)
10030     Exit Sub
errHandler:
10040     ErrorIn "ujXML.attrSetAsLkpFrom(AttrName,IndexAttr,FromList,Delim)", Array(AttrName, IndexAttr, _
               FromList, Delim)
End Sub

Public Sub attrSetAsIndexFrom(ByVal AttrName As String, ByVal ValueAttr As String, _
        ByVal FromList As String, Optional ByVal Delim As String = ",")
10050     On Error GoTo errHandler
10060     attrI(AttrName) = uIndexOf(attrS(ValueAttr), FromList, Delim)
10070     Exit Sub
errHandler:
10080     ErrorIn "ujXML.attrSetAsIndexFrom(AttrName,ValueAttr,FromList,Delim)", Array(AttrName, ValueAttr, _
               FromList, Delim)
End Sub

Public Function elGetAttrList() As String
10090     On Error GoTo errHandler
          Dim Ndl As IXMLDOMNodeList, i As Long, sList As String
10100     Set Ndl = mElem.SelectNodes("@*")
10110     For i = 0 To Ndl.Length - 1
10120         sList = uCombine(sList, ",", Ndl.Item(i).nodeName)
10130     Next i
10140     elGetAttrList = sList
10150     Exit Function
errHandler:
10160     ErrorIn "ujXML.elGetAttrList"
End Function


Public Sub attrCopy(ByVal xFrom As ujXML, _
                Optional ByVal Specs As String = "", _
                Optional ByVal AddToDict As Boolean = False)
10170     On Error GoTo errHandler
          Dim arrD, arrS, arrQ
10180     If Specs = "" Then Specs = xFrom.elGetAttrList
10190     uUnpackEQList Specs, arrD, arrS, arrQ
10200     If AddToDict Then dictAdd Join(arrD, ",")
          Dim i As Long
10210     For i = 0 To UBound(arrD)
10220         If arrQ(i) Then
10230             attrV(arrD(i)) = arrS(i)
10240             Else
10250             attrV(arrD(i)) = xFrom.attrV(arrS(i))
10260         End If
10270     Next i
10280     Exit Sub
errHandler:
10290     ErrorIn "ujXML.attrCopy(xFrom,Specs,AddToDict)", Array(xFrom, Specs, AddToDict)
End Sub

Public Function attrConstruct(ByVal AttrName As String, ByVal Template As String, _
        ByVal PrmAttrList As String) As String
10300     On Error GoTo errHandler
10310     attrV(AttrName) = uInsInto(Template, attrGetArr(PrmAttrList))
10320     Exit Function
errHandler:
10330     ErrorIn "ujXML.attrConstruct(AttrName,Template,PrmAttrList)", Array(AttrName, Template, _
               PrmAttrList)
End Function

Public Function attrGetArr(ByVal AttrList As String, Optional ByVal InsValues) As Variant

10340     On Error GoTo errHandler
          Dim arrA() As String, arrV(), Ub As Long, i As Long, InsIndex As Long, V, AName As String
10350     arrA = Split(AttrList, ",")
10360     If Not IsMissing(InsValues) Then
10370         If Not IsArray(InsValues) Then InsValues = Array(InsValues)
10380     End If
10390     ReDim arrV(LBound(arrA) To UBound(arrA))
10400     For i = LBound(arrV) To UBound(arrV)
10410         AName = arrA(i)
10420         If AName = "" Then
10430             Check Not IsMissing(InsValues), X_UJXML, "Skipped attribute found but InsValues array is not provided."
10440             V = InsValues(InsIndex)
10450             InsIndex = InsIndex + 1
10460             Else
10470             V = attrV(AName)
10480         End If
10490         arrV(i) = V
10500     Next i
10510     attrGetArr = arrV
10520     Exit Function
errHandler:
10530     ErrorIn "ujXML.attrGetArr(AttrList,InsValues)", Array(AttrList, InsValues)
End Function

Public Sub attrSetArr(ByVal AttrList As String, ByVal Values)
10540     On Error GoTo errHandler
          Dim i As Long, arrA() As String
10550     If AttrList = "" Then Exit Sub
10560     arrA = Split(AttrList, ",")
10570     Check UBound(arrA) = UBound(Values), X_UJXML, "List counts mismatch: Count(AttrList)=%1, Count(Values)=%2", _
              Array(UBound(arrA) + 1, UBound(Values) + 1)
10580     For i = 0 To UBound(arrA)
10590         attrV(arrA(i)) = Values(i)
10600     Next i
10610     Exit Sub
errHandler:
10620     ErrorIn "ujXML.attrSetArr(AttrList,Values)", Array(AttrList, Values)
End Sub

Public Sub attrRemove(ByVal AttrName As String)
10630     On Error GoTo errHandler
10640     mElem.removeAttribute AttrName
10650     Exit Sub
errHandler:
10660     ErrorIn "ujXML.attrRemove(AttrName)", AttrName
End Sub


Public Property Get Element() As IXMLDOMElement
10670     Set Element = mElem
End Property

Public Property Set Element(ByVal AnElem As IXMLDOMElement)
10680   Set mElem = AnElem
End Property

Public Sub elClear()
10690     On Error GoTo errHandler
10700     If (mElem Is docRootElem) Then
10710         docInit elName
10720         Else
10730         elCreateSibling elName, True
10740         elRemove
10750     End If
10760     Exit Sub
errHandler:
10770     ErrorIn "ujXML.elClear"
End Sub

Public Function elRemove() As Boolean
10780     On Error GoTo errHandler
          Dim Parent As IXMLDOMElement, NextEl As IXMLDOMElement
10790     Check Not mElem Is docTopElem, ERR_GENERAL, "Cannot remove document root mElement"
10800     Set Parent = mElem.ParentNode
10810     Set NextEl = mElem.PreviousSibling
10820     If NextEl Is Nothing Then Set NextEl = mElem.NextSibling
10830     If NextEl Is Nothing Then
10840         Set NextEl = Parent
10850         elRemove = False
10860         Else
10870         elRemove = True
10880     End If
10890     Parent.RemoveChild mElem
10900     mDocModified = True
10910     elAfterDelete mElem
10920     Set mElem = NextEl
10930     Exit Function
errHandler:
10940     ErrorIn "ujXML.elRemove"
End Function

Public Property Get elText() As String
10950     On Error GoTo errHandler
10960     elText = mElem.text
10970     Exit Property
errHandler:
10980     ErrorIn "ujXML.elText"
End Property

Public Property Let elText(ByVal Value As String)
10990     On Error GoTo errHandler
11000     If (mElem.text <> Value) And docOptions(XO_TRACK_CHANGES) Then
11010         elSetModified True
11020     End If
11030     mElem.text = Value
11040     Exit Property
errHandler:
11050     ErrorIn "ujXML.elText(Value)", Value
End Property

Public Sub elRemoveSubElements()
11060     On Error GoTo errHandler
11070     Do While navFirstChild
11080         If elRemove Then navUP
11090     Loop
11100     Exit Sub
errHandler:
11110     ErrorIn "ujXML.elRemoveSubElements"
End Sub

'Index/Position --------------------------------------------
Public Function elIsFirst() As Boolean
11120     On Error GoTo errHandler
11130     elIsFirst = (mElem.PreviousSibling Is Nothing)
11140     Exit Function
errHandler:
11150     ErrorIn "ujXML.elIsFirst"
End Function

Public Function elIsLast() As Boolean
11160     On Error GoTo errHandler
11170     elIsLast = (mElem.NextSibling Is Nothing)
11180     Exit Function
errHandler:
11190     ErrorIn "ujXML.elIsLast"
End Function

'Element neighbours =====================================================
Public Property Get elNextSibling() As IXMLDOMElement
11200     On Error GoTo errHandler
11210     Set elNextSibling = mElem.NextSibling
11220     Exit Property
errHandler:
11230     ErrorIn "ujXML.elNextSibling"
End Property

Public Property Get elPrevSibling() As IXMLDOMElement
11240     On Error GoTo errHandler
11250     Set elPrevSibling = mElem.PreviousSibling
11260     Exit Property
errHandler:
11270     ErrorIn "ujXML.elPrevSibling"
End Property

Public Property Get elFirstSibling() As IXMLDOMElement
11280     On Error GoTo errHandler
11290     If elParent Is Nothing Then
11300         Set elFirstSibling = mElem
11310         Else
11320         Set elFirstSibling = elParent.FirstChild
11330     End If
11340     Exit Property
errHandler:
11350     ErrorIn "ujXML.elFirstSibling"
End Property

Public Property Get elLastSibling() As IXMLDOMElement
11360     On Error GoTo errHandler
11370     If elParent Is Nothing Then
11380         Set elLastSibling = mElem
11390         Else
11400         Set elLastSibling = elParent.LastChild
11410     End If
11420     Exit Property
errHandler:
11430     ErrorIn "ujXML.elLastSibling"
End Property

Public Property Get elParent() As IXMLDOMElement
          Dim Nd As IXMLDOMNode
11440     On Error GoTo errHandler
11450     Set Nd = mElem.ParentNode
11460     If TypeOf Nd Is IXMLDOMElement Then Set elParent = Nd
11470     Exit Property
errHandler:
11480     ErrorIn "ujXML.elParent"
End Property

Public Property Get elFirstChild() As IXMLDOMElement
11490     On Error GoTo errHandler
11500     Set elFirstChild = mElem.FirstChild
11510     Exit Property
errHandler:
11520     ErrorIn "ujXML.elFirstChild"
End Property

Public Property Get elLastChild() As IXMLDOMElement
11530     On Error GoTo errHandler
11540     Set elLastChild = mElem.LastChild
11550     Exit Property
errHandler:
11560     ErrorIn "ujXML.elLastChild"
End Property

Public Function elLocateXsl(ByVal Path As String, Optional ByVal Values, _
        Optional ByVal Raise As Boolean = True) As Boolean
11570     On Error GoTo errHandler
          Dim tEl As IXMLDOMElement
11580     If Not IsMissing(Values) Then Path = uInsInto(Path, Values)
11590     Set tEl = mElem.SelectSingleNode(Path)
11600     If tEl Is Nothing Then
11610         If Raise Then
11620             Check False, X_UJXML, "Navigation from mElement [%1] using path [%2] failed.", _
                          Array(elPath, Path)
11630         End If
11640         Else
11650         Set mElem = tEl
11660         elLocateXsl = True
11670     End If
11680     Exit Function
errHandler:
11690     ErrorIn "ujXML.elLocateXsl(Path,Values,Raise)", Array(Path, Values, Raise)
End Function

Public Function elGetChildByIndex(ByVal Index As Long) As IXMLDOMElement
11700     On Error GoTo errHandler
11710     elGetChildByIndex = mElem.ChildNodes(Index)
11720     Exit Function
errHandler:
11730     ErrorIn "ujXML.elGetChildByIndex(Index)", Index
End Function

Public Function elGetSiblingByIndex(ByVal Index As Long) As IXMLDOMElement
11740     On Error GoTo errHandler
11750     If elParent Is Nothing Then Exit Function
11760     elGetSiblingByIndex = elParent.ChildNodes(Index)
11770     Exit Function
errHandler:
11780     ErrorIn "ujXML.elGetSiblingByIndex(Index)", Index
End Function

'Moving mElement =============================================================
Public Sub elMoveTo(Optional ByVal BeforeEl As IXMLDOMElement, _
        Optional ByVal AfterEl As IXMLDOMElement, _
        Optional ByVal NewPrntEl As IXMLDOMElement)
11790     On Error GoTo errHandler
11800     Check Not ((BeforeEl Is docRootElem) Or (AfterEl Is docRootElem)), X_UJXML, _
              "Before/After reference mElement may not be document root."
11810     If AfterEl Is Nothing Then
11820         If BeforeEl Is Nothing Then
11830             Check Not (NewPrntEl Is Nothing), X_UJXML, "At least one reference mElement must be provided."
11840             NewPrntEl.appendChild mElem
11850             Else
11860             BeforeEl.ParentNode.InsertBefore mElem, BeforeEl
11870         End If
11880       Else
11890         If AfterEl.NextSibling Is Nothing Then
11900             AfterEl.ParentNode.appendChild mElem
11910             Else
11920             AfterEl.ParentNode.InsertBefore mElem, AfterEl.NextSibling
11930         End If
11940     End If
11950     Exit Sub
errHandler:
11960     ErrorIn "ujXML.elMoveTo(BeforeEl,AfterEl,NewPrntEl)", Array(BeforeEl, AfterEl, NewPrntEl)
End Sub

Public Function elMoveNext() As Boolean
11970     On Error GoTo errHandler
11980     If Not elNextSibling Is Nothing Then
11990         elMoveTo , elNextSibling
12000         elMoveNext = True
12010     End If
12020     Exit Function
errHandler:
12030     ErrorIn "ujXML.elMoveNext"
End Function

Public Function elMovePrev() As Boolean
12040     On Error GoTo errHandler
12050     If Not elPrevSibling Is Nothing Then
12060         elMoveTo elPrevSibling
12070         elMovePrev = True
12080     End If
12090     Exit Function
errHandler:
12100     ErrorIn "ujXML.elMovePrev"
End Function

Public Sub elCreateSibling(Optional ByVal AnElemName As String, Optional InsAfter As Boolean = True)
12110     On Error GoTo errHandler
12120     Check Not mElem Is mTopElem, X_UJXML, "Cannot create sibling for root Element."
12130     If AnElemName = "" Then AnElemName = elName
12140     If InsAfter Then
12150         docCreateElem AnElemName, , mElem
12160         Else
12170         docCreateElem AnElemName, mElem
12180     End If
12190     Exit Sub
errHandler:
12200     ErrorIn "ujXML.elCreateSibling(AnElemName,InsAfter)", Array(AnElemName, InsAfter)
End Sub

'Returns zero-based mElement's index in siblings collection
Public Function elIndex() As Long
12210     On Error GoTo errHandler
          Dim tDoc As MSXML2.DOMDocument60, tEl As MSXML2.IXMLDOMElement, Ind As Long
12220     Set tEl = mElem
12230     While Not tEl.PreviousSibling Is Nothing
12240         Set tEl = tEl.PreviousSibling
12250         Ind = Ind + 1
12260     Wend
12270     elIndex = Ind
12280     Exit Function
errHandler:
12290     ErrorIn "ujXML.elIndex"
End Function

Public Property Get elName() As String
12300     On Error GoTo errHandler
12310     elName = mElem.nodeName
12320     Exit Property
errHandler:
12330     ErrorIn "ujXML.elName"
End Property

Public Property Get elIsTop() As Boolean
12340     On Error GoTo errHandler
12350     elIsTop = (mElem Is mTopElem)
12360     Exit Property
errHandler:
12370     ErrorIn "ujXML.elIsTop"
End Property

Public Property Get elIsRoot() As Boolean
12380     On Error GoTo errHandler
12390     elIsRoot = (mElem Is docRootElem)
12400     Exit Property
errHandler:
12410     ErrorIn "ujXML.elIsRoot"
End Property

Public Property Get elPath() As String
12420     On Error GoTo errHandler
          Dim El As IXMLDOMElement
12430     Set El = mElem
12440     elPath = PathPriv
12450     Set mElem = El
12460     Exit Property
errHandler:
12470     ErrorIn "ujXML.elPath"
End Property

Private Function PathPriv() As String
12480     On Error GoTo errHandler
          Dim Ind As Long, ParPath As String, ElRef As String
12490     If elIsRoot Then
12500         PathPriv = "/" & elName
12510         Else
12520         Ind = elIndex
12530         ElRef = elName
12540         If Ind > 0 Then ElRef = ElRef & "[" & Ind & "]"
12550         navUP
12560         ParPath = PathPriv()
12570         PathPriv = ParPath & "/" & ElRef
12580     End If
12590     Exit Function
errHandler:
12600     ErrorIn "ujXML.PathPriv"
End Function

Public Property Get elXML() As String
12610     On Error GoTo errHandler
12620     elXML = mElem.xml
12630     Exit Property
errHandler:
12640     ErrorIn "ujXML.elXML"
End Property

Public Property Let elXML(ByVal xmlElem As String)
12650     On Error GoTo errHandler
          Dim xTmp As ujXML, ParNd As IXMLDOMNode, NewEl As IXMLDOMElement
12660     If mElem Is docRootElem Then
12670         docLoadXML xmlElem
12680         Else
12690         Set ParNd = mElem.ParentNode
12700         Set xTmp = New ujXML
12710         xTmp.docProgID = docProgID
12720         xTmp.docLoadXML xmlElem
12730         Set NewEl = xTmp.docRootElem
12740         ParNd.replaceChild NewEl, mElem
12750     End If
12760     If mTopElem Is mElem Then Set mTopElem = NewEl
12770     elAfterDelete mElem
12780     Set mElem = NewEl
12790     Exit Property
errHandler:
12800     ErrorIn "ujXML.elXML(xmlElem)", xmlElem
End Property

Private Sub elAfterDelete(ByVal El As IXMLDOMElement)
12810     On Error GoTo errHandler
12820     bkmkRemoveEl El
12830     histRemoveEl El
12840     Exit Sub
errHandler:
12850     ErrorIn "ujXML.elAfterDelete(El)", El
End Sub

Public Property Get elStatus() As XENUM_ELEM_STATUS
12860     On Error GoTo errHandler
          Dim Ch As String
12870     Ch = attrS(S_STATUS)
12880     Select Case UCase(Ch)
              Case "", "E": elStatus = XES_EXISTING
12890         Case "N": elStatus = XES_NEW
12900         Case "M": elStatus = XES_MODIFIED
12910         Case "D": elStatus = XES_DELETED
12920         Case "J": elStatus = XES_JOINED
12930     End Select
12940     Exit Property
errHandler:
12950     ErrorIn "ujXML.elStatus"
End Property

Public Property Let elStatus(ByVal AValue As XENUM_ELEM_STATUS)
12960     On Error GoTo errHandler
12970     Select Case AValue
              Case XES_EXISTING: mElem.removeAttribute S_STATUS
12980         Case XES_NEW: attrS(S_STATUS) = "N"
12990         Case XES_MODIFIED: attrS(S_STATUS) = "M"
13000         Case XES_DELETED: attrS(S_STATUS) = "D"
13010         Case XES_JOINED: attrS(S_STATUS) = "J"
13020     End Select
13030     Exit Property
errHandler:
13040     ErrorIn "ujXML.elStatus(AValue)", AValue
End Property

Public Property Get elVisible() As Boolean
13050     On Error GoTo errHandler
13060     Select Case elStatus
              Case XES_EXISTING, XES_MODIFIED, XES_NEW:  elVisible = True
13070     End Select
13080     Exit Property
errHandler:
13090     ErrorIn "ujXML.elVisible"
End Property

Public Function elCheckValues(ByVal AttrNames, ByVal AttrValues) As Boolean
13100     On Error GoTo errHandler
          Dim arrNames() As String, i As Long
13110     If Not IsArray(AttrValues) Then AttrValues = Array(AttrValues)
13120     If IsArray(AttrNames) Then
13130         arrNames = AttrNames
13140         Else
13150         If AttrNames = "" Then elCheckValues = True: Exit Function
13160         arrNames = Split(AttrNames, ",")
13170         Check UBound(arrNames) = UBound(AttrValues), X_UJXML, _
                  "Attribute names/values count mismatch: Attributes=%1, Values = %2", _
                  Array(UBound(arrNames) + 1, UBound(AttrValues) + 1)
13180     End If
13190     For i = LBound(arrNames) To UBound(arrNames)
13200         If attrS(arrNames(i)) <> AttrValues(i) Then Exit Function  'returning false
13210     Next i
13220     elCheckValues = True
13230     Exit Function
errHandler:
13240     ErrorIn "ujXML.elCheckValues(AttrNames,AttrValues)", Array(AttrNames, AttrValues)
End Function

Public Function elJoin(ByVal xWithList As ujXML, ByVal OnAttr As String, ByVal CopyList As String, _
        Optional ByVal TblBkmk As String)
13250     On Error GoTo errHandler
          Dim SrchAttr As String, SrcAttr As String, arrT() As String, KeyV
13260     arrT = Split(OnAttr, "=")
13270     SrcAttr = arrT(0)
13280     If UBound(arrT) = 0 Then SrchAttr = SrcAttr Else SrchAttr = arrT(1)
13290     KeyV = attrS(SrcAttr)
13300     If Not xWithList.tblLocateElem(TblBkmk, SrchAttr, KeyV) Then Exit Function
13310     attrCopy xWithList, CopyList
13320     Exit Function
errHandler:
13330     ErrorIn "ujXML.elJoin(xWithList,OnAttr,CopyList,TblBkmk)", Array(xWithList, OnAttr, CopyList, _
               TblBkmk)
End Function

Public Sub elForEachElem(ByVal Iterator As Object, Optional Param)
13340     On Error GoTo errHandler
          Dim SkipChildren As Boolean
13350     Iterator.ProcessElement Me, XNAV_START, Param, SkipChildren
13360     If chCount > 0 And Not SkipChildren Then elForEachChild Iterator, Param
13370     Exit Sub
errHandler:
13380     ErrPreserve
13390     ErrorIn "ujXML.elForEachElem(Iterator,Param)", Array(Iterator, Param)
End Sub

Private Sub elForEachChild(ByVal Iterator As Object, Optional Param)
13400     On Error GoTo errHandler
          Dim SkipChildren As Boolean, nav As XENUM_ITER_NAV
13410     If Not navFirstChild Then Exit Sub
13420     nav = XNAV_TO_CHILD
13430     Do
13440         SkipChildren = False
13450         Iterator.ProcessElement Me, nav, Param, SkipChildren
13460         If chCount > 0 And Not SkipChildren Then elForEachChild Iterator, Param
13470         nav = XNAV_TO_NEXT
13480     Loop While navNext
13490     navUP
13500     Iterator.ProcessElement Me, XNAV_TO_PARENT, Param, SkipChildren
13510     Exit Sub
errHandler:
13520     ErrPreserve
13530     ErrorIn "ujXML.elForEachChild(Iterator,Param)", Array(Iterator, Param)
13540     Exit Sub
13550     Resume
End Sub

'Moving/Navigating =============================================================================================
Public Function navNext() As Boolean
13560     On Error GoTo errHandler
          Dim El As MSXML2.IXMLDOMElement, sName As String
13570     Set El = mElem.NextSibling
13580     If Not El Is Nothing Then
13590         Set mElem = El
13600         navNext = True
13610     End If
13620     Exit Function
errHandler:
13630     ErrorIn "ujXML.navNext"
End Function

Public Function navPrevious() As Boolean
13640     On Error GoTo errHandler
          Dim El As MSXML2.IXMLDOMElement, sName As String
13650     Set El = mElem.PreviousSibling
13660     If Not El Is Nothing Then
13670         Set mElem = El
13680         navPrevious = True
13690     End If
13700     Exit Function
errHandler:
13710     ErrorIn "ujXML.navPrevious"
End Function

Public Sub navFirst()
13720     On Error GoTo errHandler
13730     Set mElem = mElem.ParentNode.FirstChild
13740     Exit Sub
errHandler:
13750     ErrorIn "ujXML.navFirst"
End Sub

Public Sub navLast()
13760     On Error GoTo errHandler
13770     Set mElem = mElem.ParentNode.LastChild
13780     Exit Sub
errHandler:
13790     ErrorIn "ujXML.navLast"
End Sub

Public Function navFirstChild() As Boolean
13800     On Error GoTo errHandler
          Dim El As MSXML2.IXMLDOMElement
13810     Set El = mElem.FirstChild
13820     If Not El Is Nothing Then
13830         Set mElem = El
13840         navFirstChild = True
13850     End If
13860     Exit Function
errHandler:
13870     If Err = 13 Then
13880         navFirstChild = False
13890         Exit Function
13900     End If
13910     ErrorIn "ujXML.navFirstChild"
End Function

Public Function navLastChild() As Boolean
13920     On Error GoTo errHandler
          Dim El As MSXML2.IXMLDOMElement
13930     Set El = mElem.LastChild
13940     If Not El Is Nothing Then
13950         Set mElem = El
13960         navLastChild = True
13970     End If
13980     Exit Function
errHandler:
13990     ErrorIn "ujXML.navLastChild"
End Function

Public Sub navChildByIndex(ByVal Index As Long)
14000     On Error GoTo errHandler
          Dim El As MSXML2.IXMLDOMElement
14010     Set El = mElem.ChildNodes(Index)
14020     Check Not El Is Nothing, X_UJXML, "Child not found for index = " & Index
14030     Set mElem = El
14040     Exit Sub
errHandler:
14050     ErrorIn "ujXML.navChildByIndex"
End Sub

'Public Sub navUP()
'    On Error GoTo errHandler
'    Dim Nd As msxml2.IXMLDOMNode
'    Set Nd = mElem.parentNode
'    Check (TypeOf Nd Is IXMLDOMElement), X_UJXML, "Cannot move up from document root"
'    Set mElem = Nd
'    Exit Sub
'errHandler:
'    ErrorIn "ujXML.navUP"
'End Sub
Public Sub navUP()
14060     On Error GoTo errHandler
          Dim Nd As MSXML2.IXMLDOMNode
14070     Set Nd = mElem.ParentNode
14080     Check (TypeOf Nd Is IXMLDOMElement), X_UJXML, "Cannot move up from document root"
14090     Set mElem = Nd
14100     Exit Sub
errHandler:
14110     ErrorIn "ujXML.navUP"
End Sub
Public Sub navTop()
14120     On Error GoTo errHandler
14130     Set mElem = docTopElem
14140     Exit Sub
errHandler:
14150     ErrorIn "ujXML.navTop"
End Sub

Public Function navLocate(ByVal AnElemName As String, Optional ByVal AttrName As String, Optional ByVal Value As String) As Boolean
14160     On Error GoTo errHandler
          Dim El As MSXML2.IXMLDOMElement, Path As String
14170     Path = "//" & AnElemName
14180     If AttrName <> "" Then
14190         Path = Path & "[@" & AttrName & "=" & uSmartQuote(Value) & "]"
14200     End If
14210     Set El = mElem.SelectSingleNode(Path)
14220     If Not El Is Nothing Then
14230         Set mElem = El
14240         navLocate = True
14250     End If
14260     Exit Function
errHandler:
14270     ErrorIn "ujXML.navLocate(ElemName,AttrName,Value)", Array(AnElemName, AttrName, Value)
End Function

Public Function navTo(ByVal Path As String, Optional ByVal Index As Long = -1) As Boolean
14280     On Error GoTo errHandler
          Dim El As MSXML2.IXMLDOMElement
14290     If Index <> -1 Then Path = Path & "[" & Index & "]"
14300     Set El = mElem.SelectSingleNode(Path)
14310     If Not El Is Nothing Then
14320         Set mElem = El
14330         navTo = True
14340     End If
14350     Exit Function
errHandler:
14360     ErrorIn "ujXML.navTo(Path,Index)", Array(Path, Index)
End Function

'Child ===========================================================================================
Public Sub chCreate(ByVal AnElemName As String, Optional ByVal NavigateToIt As Boolean = True, _
        Optional ByVal AsLast As Boolean = True, Optional ByVal MarkAsNew As Boolean = True)
14370     On Error GoTo errHandler
          Dim El As MSXML2.IXMLDOMElement
14380     Set El = mDoc.createElement(AnElemName)
14390     With mElem
14400         If Not .HasChildNodes Then AsLast = True
14410         If AsLast Then
14420             .appendChild El
14430         Else
14440             .InsertBefore El, .FirstChild
14450         End If
14460         If docOptions(XO_TRACK_CHANGES) And MarkAsNew Then El.setAttribute S_STATUS, "N"
14470     End With
14480     If NavigateToIt Then Set mElem = El
14490     Exit Sub
errHandler:
14500     ErrorIn "ujXML.chCreate(AnElemName,NavigateToIt,AsLast,MarkAsNew)", Array(AnElemName, NavigateToIt, _
               AsLast, MarkAsNew)
End Sub

Public Property Get chCount() As Long
14510     On Error GoTo errHandler
14520     chCount = mElem.ChildNodes.Length
14530     Exit Property
errHandler:
14540     ErrorIn "ujXML.chCount"
End Property

Public Function chLocate(Optional ChElemName As String, Optional ByVal AttrNames As String, _
        Optional ByVal AttrValues) As Boolean
14550     On Error GoTo errHandler
          Dim Found As Boolean, NameMismatch As Boolean, arrNames() As String
14560     If AttrNames <> "" Then
14570         arrNames = Split(AttrNames, ",")
14580         If Not IsArray(AttrValues) Then AttrValues = Array(AttrValues)
14590         Check UBound(arrNames) = UBound(AttrValues), X_UJXML, "Invalid parameters: " & _
                  "Number of attributes (%1) doesn't match number of Values (%2)", _
              Array(UBound(arrNames) + 1, UBound(AttrValues) + 1)
14600     End If
14610     If Not navFirstChild Then Exit Function
14620     Do
14630         NameMismatch = (ChElemName <> "") And (elName <> ChElemName)
14640         If (Not NameMismatch) Then
14650             If (AttrNames = "") Then
14660                 chLocate = True
14670                 Exit Function
14680             Else
14690                 If elCheckValues(arrNames, AttrValues) Then chLocate = True: Exit Function
14700             End If '(AttrNames = ....
14710         End If '(Not NameMismatch ....
14720     Loop While navNext
14730     navUP
14740     Exit Function
errHandler:
14750     ErrorIn "ujXML.chLocate(ChElemName,AttrNames,AttrValues)", Array(ChElemName, AttrNames, _
               AttrValues)
End Function

Public Function chLocateByIndex(ByVal Index As Long, Optional ByVal VisOnly As Boolean)
14760     On Error GoTo errHandler
          Dim Nd As IXMLDOMNode, Path As String
14770     Set Nd = mElem.SelectSingleNode("*[" & Index & "]")
14780     Exit Function
errHandler:
14790     ErrorIn "ujXML.chLocateByIndex(Index,VisOnly)", Array(Index, VisOnly)
End Function

Public Property Get chText(ByVal ChildName As String) As String
14800     On Error GoTo errHandler
14810     If chLocate(ChildName) Then
14820         chText = elText
14830         navUP
14840     End If
14850     Exit Property
errHandler:
14860     ErrorIn "ujXML.chText(ChildName)", ChildName
End Property

Public Property Let chText(ByVal ChildName As String, ByVal Value As String)
14870     On Error GoTo errHandler
14880     If chLocate(ChildName) Then
14890         elText = Value
14900         Else
14910         chCreate ChildName
14920         elText = Value
14930     End If
14940     navUP
14950     Exit Property
errHandler:
14960     ErrorIn "ujXML.chText(ChildName,Value)", Array(ChildName, Value)
End Property

Public Sub chAConstruct(ByVal AttrName As String, ByVal Template As String, ByVal PrmAttrList As String)
14970     On Error GoTo errHandler
14980     If navFirstChild Then
14990         Do
15000             attrConstruct AttrName, Template, PrmAttrList
15010         Loop While navNext
15020         navUP
15030     End If
15040     Exit Sub
errHandler:
15050     ErrorIn "ujXML.chAConstruct(AttrName,Template,PrmAttrList)", Array(AttrName, Template, _
               PrmAttrList)
End Sub

Public Function chCalc(ByVal AttrName As String, ByVal Oper As XENUM_OPER, _
        Optional ByVal FilterAttr As String, Optional ByVal FilterValues, _
        Optional ByVal Delim As String = ",") As Variant
15060     On Error GoTo errHandler
          Dim Res, V, dblV As Double
15070     If Not IsArray(FilterValues) Then FilterValues = Array(FilterValues)
15080     If navFirstChild Then
15090         Do
15100             If elCheckValues(FilterAttr, FilterValues) Then
15110                 V = Me.attrV(AttrName)
15120                 dblV = uSafeDbl(V)
15130                 Select Case Oper
                          Case XOP_MIN: If IsEmpty(Res) Then Res = dblV Else Res = IIf(dblV < Res, dblV, Res)
15140                     Case XOP_MAX: If IsEmpty(Res) Then Res = dblV Else Res = IIf(dblV > Res, dblV, Res)
15150                     Case XOP_SUM: If IsEmpty(Res) Then Res = dblV Else Res = Res + dblV
15160                     Case XOP_CONCAT: If IsEmpty(Res) Then Res = V Else Res = Res & Delim & V
15170                     Case XOP_LIST
15180                         If IsEmpty(Res) Then
15190                             Res = V
15200                             Else
15210                             Res = uListAdd(Res, V, Delim)
15220                         End If
15230                     Case XOP_COUNT: If IsEmpty(Res) Then Res = 1 Else Res = Res + 1
15240                 End Select
15250             End If 'elCheckFilter...
15260         Loop While navNext
15270         navUP
15280     End If
15290     chCalc = Res
15300     Exit Function
errHandler:
15310     ErrorIn "ujXML.chCalc(AttrName,Oper,FilterAttr,FilterValues,Delim)", Array(AttrName, Oper, _
               FilterAttr, FilterValues, Delim)
End Function

Public Sub chACopy(ByVal xFrom As ujXML, _
                Optional ByVal AttrList As String = "")
15320     On Error GoTo errHandler
15330     histPush
15340     xFrom.histPush
15350     If navFirstChild And xFrom.navFirstChild Then
15360         Do
15370             attrCopy xFrom, AttrList
15380         Loop While navNext And xFrom.navNext
15390     End If
15400     histGoBack
15410     xFrom.histGoBack
15420     Exit Sub
errHandler:
15430     ErrorIn "ujXML.chACopy(xFrom,AttrList)", Array(xFrom, AttrList)
End Sub

Public Sub chASet(ByVal AttrName As String, ByVal Value, _
        Optional ByVal FilterAtrs As String, Optional FilterValues)
15440     On Error GoTo errHandler
          Dim ToAssign As Boolean, arrAtrs() As String
15450     If FilterAtrs <> "" Then arrAtrs = Split(FilterAtrs, ",")
15460     If navFirstChild Then
15470         Do
15480             If FilterAtrs = "" Then
15490                 ToAssign = True
15500                 Else
15510                 ToAssign = elCheckValues(arrAtrs, FilterValues)
15520             End If
15530             If ToAssign Then attrV(AttrName) = Value
15540         Loop While navNext
15550         navUP
15560     End If
15570     Exit Sub
errHandler:
15580     ErrorIn "ujXML.chASet(AttrName,Value,FilterAtrs,FilterValues)", Array(AttrName, Value, FilterAtrs, _
               FilterValues)
End Sub

Public Sub chASetAsLkpFrom(ByVal AttrName As String, IndexAttr As String, _
        ByVal FromList As String, Optional ByVal Delim As String = ",")
15590     On Error GoTo errHandler
15600     If navFirstChild Then
15610         Do
15620             attrSetAsLkpFrom AttrName, IndexAttr, FromList, Delim
15630         Loop While navNext
15640         navUP
15650     End If
15660     Exit Sub
errHandler:
15670     ErrorIn "ujXML.chASetAsLkpFrom(AttrName,IndexAttr,FromList,Delim)", Array(AttrName, IndexAttr, _
               FromList, Delim)
End Sub

Public Sub chASetAsIndexFrom(ByVal AttrName As String, ByVal ValueAttr As String, _
        ByVal FromList As String, Optional ByVal Delim As String = ",")
15680     On Error GoTo errHandler
15690     If navFirstChild Then
15700         Do
15710             attrSetAsIndexFrom AttrName, ValueAttr, FromList, Delim
15720         Loop While navNext
15730         navUP
15740     End If
15750     Exit Sub
errHandler:
15760     ErrorIn "ujXML.chASetAsIndexFrom(AttrName,ValueAttr,FromList,Delim)", Array(AttrName, ValueAttr, _
               FromList, Delim)
End Sub


Public Sub chCopyElements(ByVal xFrom As ujXML, Optional ClearExisting As Boolean = True)
15770     On Error GoTo errHandler
          Dim El As MSXML2.IXMLDOMElement
15780     If ClearExisting Then chRemoveAll
15790     With xFrom
15800         If .navFirstChild Then
15810             Do
15820                 chCreate "Temp"
15830                 elXML = .elXML
15840                 navUP
15850             Loop While .navNext
15860         End If
15870     End With
15880     Exit Sub
errHandler:
15890     ErrorIn "ujXML.chCopyElements(xFrom,ClearExisting)", Array(xFrom, ClearExisting)
End Sub

Public Sub chJoinLeft(ByVal xWithList As ujXML, ByVal OnAttr As String, _
        Optional ByVal CopyList As String, Optional ByVal TblBkmk As String)
15900     On Error GoTo errHandler
          Dim SrchAttr As String, SrcAttr As String, arrT() As String
15910     arrT = Split(OnAttr, "=")
15920     SrcAttr = arrT(0)
15930     If UBound(arrT) = 0 Then SrchAttr = SrcAttr Else SrchAttr = arrT(1)
15940     If navFirstChild Then
15950         Do
15960             If xWithList.tblLocateElem(TblBkmk, SrchAttr, attrS(SrcAttr)) Then attrCopy xWithList, CopyList
15970         Loop While navNext
15980         navUP
15990     End If
16000     Exit Sub
errHandler:
16010     ErrorIn "ujXML.chJoinLeft(xWithList,OnAttr,CopyList,TblBkmk)", Array(xWithList, OnAttr, _
               CopyList, TblBkmk)
End Sub

Public Sub chJoinRight(ByVal xWithList As ujXML, ByVal OnAttr As String, _
        ByVal CopyList As String, ByVal NewElName As String, _
        Optional ByVal TblBkmk As String, Optional ByVal NewElAttrList As String, _
        Optional ByVal arrNewElValues)
16020     On Error GoTo errHandler
          Dim SrchAttr As String, SrcAttr As String, arrT() As String, KeyV, i As Long
          Dim tListEl As MSXML2.IXMLDOMElement, tLastIns As MSXML2.IXMLDOMElement
16030     arrT = Split(OnAttr, "=")
16040     SrcAttr = arrT(0)
16050     If UBound(arrT) = 0 Then SrchAttr = SrcAttr Else SrchAttr = arrT(1)
16060     Set tListEl = mElem
16070     With xWithList
16080         If TblBkmk <> "" Then .tblLocate TblBkmk
16090         If .navLastChild Then
16100             Do
16110                 Set mElem = tListEl
16120                 If chLocate(, SrcAttr, .attrS(SrchAttr)) Then
16130                     attrCopy xWithList, CopyList
16140                     Else
16150                     Set mElem = tListEl
16160                     chCreate NewElName, , False
16170                     attrCopy xWithList, CopyList
16180                     attrSetArr NewElAttrList, arrNewElValues
16190                     If docOptions(XO_TRACK_CHANGES) Then elStatus = XES_JOINED
16200                 End If
                      'Insert it at the end
16210                 tListEl.RemoveChild mElem
                      'Set mElem.parentNode = Nothing
16220                 If tLastIns Is Nothing Then
16230                     tListEl.appendChild mElem
16240                     Else
16250                     tListEl.InsertBefore mElem, tLastIns
16260                 End If
16270                 Set tLastIns = mElem
16280             Loop While .navPrevious
16290         End If
16300     End With
16310     Set mElem = tListEl
16320     Exit Sub
errHandler:
16330     ErrorIn "ujXML.chJoinRight(xWithList,OnAttr,CopyList,NewElName,TblBkmk,NewElAttrList," & _
              "arrNewElValues)", Array(xWithList, OnAttr, CopyList, NewElName, TblBkmk, NewElAttrList, _
               arrNewElValues)
End Sub

Public Sub chSort(ByVal AttrTypeDirSpecs As String, _
            Optional ByVal DefaultDir As String = "asc")
16340     On Error GoTo errHandler
          Dim arrSpecs() As String, arrNames() As String, arrTypes() As String, arrDirs() As String
          Dim UbndSpecs As Long, arrSDet() As String, iElem As Long, jSpec As Long
          Dim mElemCnt As Long, arrElems(), arrV(), arrElValues()
16350     If AttrTypeDirSpecs = "" Then Exit Sub
16360     AttrTypeDirSpecs = Replace(AttrTypeDirSpecs, vbNewLine, ",")
16370     AttrTypeDirSpecs = Replace(AttrTypeDirSpecs, ",,", ",")
16380     mElemCnt = chCount
16390     If chCount <= 1 Then Exit Sub
16400     Check uInList("asc,desc", DefaultDir), X_UJXML, "Invalid DefaultDir specification: " & DefaultDir
16410     arrSpecs = Split(AttrTypeDirSpecs, ",")     'Unpack AttrTypeDirSpecs
16420     UbndSpecs = UBound(arrSpecs)
16430     ReDim arrNames(0 To UbndSpecs), arrTypes(0 To UbndSpecs), arrDirs(0 To UbndSpecs)
16440     For iElem = 0 To UbndSpecs
16450         arrSDet = Split(arrSpecs(iElem), ":")
16460         arrNames(iElem) = arrSDet(0)
16470         If UBound(arrSDet) < 1 Then arrTypes(iElem) = "string" Else arrTypes(iElem) = arrSDet(1)
16480         If UBound(arrSDet) < 2 Then arrDirs(iElem) = DefaultDir Else arrDirs(iElem) = LCase$(arrSDet(2))
16490         Check uInList("asc,desc", arrDirs(iElem)), X_UJXML, "Invalid Dir specification: " & arrDirs(iElem)
16500     Next iElem
          'Now create mElements array, and array of [Values array]
16510     ReDim arrElems(0 To mElemCnt - 1), arrV(0 To mElemCnt - 1)
16520     navFirstChild
16530     For iElem = 0 To mElemCnt - 1
16540         Set arrElems(iElem) = mElem
16550         arrElValues = uCreateArr(UbndSpecs)
16560         For jSpec = 0 To UbndSpecs
16570             Select Case LCase(arrTypes(jSpec))
                      Case "string": arrElValues(jSpec) = attrS(arrNames(jSpec))
16580                 Case "number": arrElValues(jSpec) = attrF(arrNames(jSpec))
16590                 Case "date": arrElValues(jSpec) = attrDt(arrNames(jSpec))
16600                 Case Else:
16610                     navUP
16620                     Check False, X_UJXML, "Invalid type specification [%1]. Must be one of the following: Date,String,Number", arrTypes(jSpec)
16630             End Select
16640         Next jSpec
16650         arrV(iElem) = arrElValues
16660         navNext
16670     Next iElem
16680     sortShell arrElems, arrV, arrDirs
16690     navUP
          'Now order mElements --------------------------------------------------------
          Dim tLastInsEl As IXMLDOMElement, tCurrEl As IXMLDOMElement
16700     For iElem = mElemCnt - 1 To 0 Step -1
16710         Set tCurrEl = arrElems(iElem)
16720         mElem.RemoveChild tCurrEl
16730         If tLastInsEl Is Nothing Then
16740             mElem.appendChild tCurrEl
16750             Else
16760             mElem.InsertBefore tCurrEl, tLastInsEl
16770         End If
16780         Set tLastInsEl = tCurrEl
16790     Next iElem
16800     Exit Sub
errHandler:
16810     ErrorIn "ujXML.chSort(AttrTypeDirSpecs,DefaultDir)", Array(AttrTypeDirSpecs, DefaultDir)
End Sub

Private Sub sortShell(ByRef arrElems(), ByRef arrCompV(), ByRef arrDirs() As String)
16820     On Error GoTo errHandler
          Dim NewFP As Long, NewLP As Long
          Dim Dist As Long, FwdLimit As Long, cnt As Long
          Dim FPtr As Long, LPtr As Long, ExchFlag As Boolean
16830     cnt = UBound(arrElems)
16840     Dist = cnt / 2
16850     Do Until Dist < 1
16860         FwdLimit = cnt - Dist
16870         FPtr = 0
16880         Do Until FPtr > FwdLimit
16890             NewFP = FPtr
16900             LPtr = FPtr + Dist
16910             If (sortCompare(arrCompV(NewFP), arrCompV(LPtr), arrDirs) = 1) Then
16920                 ExchFlag = True
16930                 Do While ExchFlag
16940                     sortSwap arrElems, NewFP, LPtr
16950                     sortSwap arrCompV, NewFP, LPtr
16960                     NewLP = NewFP
16970                     NewFP = NewFP - Dist
16980                     If NewFP < LBound(arrElems) Then
16990                         ExchFlag = False
17000                     Else
17010                         If (sortCompare(arrCompV(NewFP), arrCompV(NewLP), arrDirs) = 1) Then
17020                             LPtr = NewLP
17030                         Else
17040                             ExchFlag = False
17050                         End If
17060                     End If
17070                 Loop
17080             End If
17090             FPtr = FPtr + 1
17100         Loop
17110     Dist = Dist / 2
17120     Loop
17130     Exit Sub
errHandler:
17140     ErrorIn "ujXML.sortShell(arrElems,arrCompV,arrDirs)", Array(arrElems, arrCompV, arrDirs)
End Sub

Private Function sortCompare(ByRef v1, ByRef v2, arrDirs() As String) As Long
17150     On Error GoTo errHandler
          Dim i As Long
17160     sortCompare = 0
17170     For i = LBound(v1) To UBound(v1)
17180         If v1(i) < v2(i) Then
17190             sortCompare = -1
17200         ElseIf v1(i) > v2(i) Then
17210             sortCompare = 1
17220         End If
17230         If arrDirs(i) = "desc" Then sortCompare = -sortCompare
17240         If sortCompare <> 0 Then
17250             Exit Function
17260         End If
17270     Next i
17280     Exit Function
errHandler:
17290     ErrorIn "ujXML.sortCompare(V1,V2,arrDirs)", Array(v1, v2, arrDirs)
End Function

Private Sub sortSwap(ByRef arr(), ByVal Ind1 As Long, ByVal Ind2 As Long)
17300     On Error GoTo errHandler
          Dim tmp
17310     If IsObject(arr(Ind1)) Then
17320         Set tmp = arr(Ind1)
17330         Set arr(Ind1) = arr(Ind2)
17340         Set arr(Ind2) = tmp
17350         Else
17360         tmp = arr(Ind1)
17370         arr(Ind1) = arr(Ind2)
17380         arr(Ind2) = tmp
17390     End If
17400     Exit Sub
errHandler:
17410     ErrorIn "ujXML.sortSwap(Arr,Ind1,Ind2)", Array(arr, Ind1, Ind2)
End Sub

Public Function chFindDuplicates(ByVal AttrList As String, _
        Optional ByVal VisOnly As Boolean = False) As String
17420     On Error GoTo errHandler
          Dim strDuplicates As String
          Dim arrANms() As String, strElemValues As String, strValues As String, i As Long
17430     strDuplicates = ""
17440     arrANms = Split(AttrList, ",")
17450     histPush
17460     If navFirstChild Then
17470         Do
17480             strElemValues = ""
17490             For i = 0 To UBound(arrANms)
17500                 strElemValues = uCombine(strElemValues, vbTab, attrS(arrANms(i)))
17510                 Next i
17520             If InStr(1, strValues, strElemValues) > 0 Then
                      'Add strElemValues to duplicates if it's not there
17530                 If InStr(1, strDuplicates, strElemValues) = 0 Then strDuplicates = uCombine(strDuplicates, vbNewLine, strElemValues)
17540                 Else
                      'add strElemValues to values list
17550                 strValues = uCombine(strValues, ";", strElemValues)
17560             End If 'InStr(1, strValues....
17570         Loop While navNext
17580     End If
17590     chFindDuplicates = strDuplicates
17600     Exit Function
errHandler:
17610     ErrorIn "ujXML.chFindDuplicates(AttrList,VisOnly)", Array(AttrList, VisOnly)
End Function

Public Sub chRemoveAll()
17620     On Error GoTo errHandler
17630     If navFirstChild Then
17640         While elRemove
17650         Wend
17660     End If
17670     Exit Sub
errHandler:
17680     ErrorIn "ujXML.chRemoveAll"
End Sub


'Document ========================================================================================
'Use docInit when you're going to build xml document from scratch. Use docLoad to when you're
'going to manipulate with xml message that you already have.
Public Sub docInit(ByVal RootElemName As String, Optional ByVal AnOptions)
17690     On Error GoTo errHandler
17700     CheckSetSysHandler
17710     Set mDoc = CreateDocObject
17720     Set mTopElem = mDoc.createElement(RootElemName)
17730     mDoc.appendChild mTopElem
17740     Set mElem = mTopElem
17750     If Not IsMissing(AnOptions) Then mOptions = AnOptions
17760     mDocModified = False
17770     Exit Sub
errHandler:
17780     ErrorIn "ujXML.docInit(RootElemName)", RootElemName
End Sub

Public Function docLoadXML(ByVal xmlData As String) As Boolean
17790     On Error GoTo errHandler
17800     CheckSetSysHandler
17810     Set mDoc = CreateDocObject
        '  MsgBox "Length=" & CStr(Len(xmlData))
       '   LogSaveToFile "xmlData=" & xmlData
17820     Check mDoc.LoadXML(xmlData), ERR_GENERAL, "LoadXML method failed"
17830     Set mTopElem = docRootElem ' mDoc.selectSingleNode("*[nodeType()=1]") ' .firstChild
17840     Set mElem = mTopElem
17850     mDocModified = False
        docLoadXML = True
17860     Exit Function
errHandler:
            docLoadXML = False
            Exit Function
17870     GetFromServer New ujEEDomDoc, mDoc
17880     ErrorIn "ujXML.docLoadXML(xmlData)", xmlData
            docLoadXML = False
End Function

Private Function CreateDocObject() As Object
17890     On Error GoTo errHandler
17900     Set CreateDocObject = CreateObject("MSXML2.DOMDocument.6.0")
17910     Exit Function
errHandler:
17920     ErrorIn "ujXML.CreateDocObject"
End Function

Public Property Get docProgID() As String
17930     docProgID = IIf(mDocProgID = "", S_DEFAULT_PROGID, mDocProgID)
End Property

Public Property Let docProgID(ByVal Value As String)
17940     mDocProgID = Value
End Property

Public Property Get docObject() As IXMLDOMDocument
17950   Set docObject = mDoc
End Property

Public Property Set docObject(ByVal ADoc As IXMLDOMDocument)
17960     On Error GoTo errHandler
17970     Set mDoc = ADoc
17980     Set mTopElem = docRootElem
17990     Set mElem = mTopElem
18000     bkmkClearAll
18010     histClear
18020     Exit Property
errHandler:
18030     ErrorIn "ujXML.docObject(ADoc)", ADoc
End Property

Public Property Get docRootElem() As IXMLDOMElement
18040     On Error GoTo errHandler
          Dim Nd As IXMLDOMNode
18050     Set Nd = mDoc.SelectSingleNode("/*")
18060     Check Not (Nd Is Nothing), X_UJXML, "Failed to find root node." ' .documentElement"
18070     Set docRootElem = Nd
18080     Exit Property
errHandler:
18090     ErrorIn "ujXML.docRootElem"
End Property

Public Property Get docOptions(ByVal mask As XENUM_OPTIONS) As Boolean
18100     On Error GoTo errHandler
18110     docOptions = ((mOptions And mask) <> 0)
18120     Exit Property
errHandler:
18130     ErrorIn "ujXML.docOptions(mask)", mask
End Property

Public Property Let docOptions(ByVal mask As XENUM_OPTIONS, ByVal Value As Boolean)
18140     On Error GoTo errHandler
18150     If Value Then mOptions = mOptions Or mask Else mOptions = mOptions And Not mask
18160     Exit Property
errHandler:
18170     ErrorIn "ujXML.docOptions(mask,Value)", Array(mask, Value)
End Property

Public Sub docSetOptions(ByVal AnOptions As Long)
18180     mOptions = AnOptions
End Sub

Public Property Get docModified() As Boolean
18190     On Error GoTo errHandler
18200     docModified = mDocModified
18210     Exit Property
errHandler:
18220     ErrorIn "ujXML.docModified"
End Property

Public Property Let docModified(ByVal AValue As Boolean)
18230     On Error GoTo errHandler
18240     mDocModified = AValue
18250     Exit Property
errHandler:
18260     ErrorIn "ujXML.docModified(AValue)", AValue
End Property

Public Property Get docXML() As String
18270     On Error GoTo errHandler

       '   docXML = Replace(mDoc.xml, vbCrLf, "") ' mTopElem.xml
18280     docXML = mDoc.xml
18290     Exit Property
errHandler:
18300     ErrorIn "ujXML.docXML"
End Property

Public Sub docReadFromFile(ByVal FilePath As String, Optional ByVal CharSet As String = "UNICODE")
18310     On Error GoTo errHandler
          Dim sXml As String
          Dim s As Object
18320     Set s = CreateObject("ADODB.Stream")
18330     With s
18340         If CharSet <> "" Then .CharSet = CharSet
18350         .Open
              Dim fs As New FileSystemObject
            '  MsgBox fs.FileExists(FilePath)
18360         .LoadFromFile FilePath
18370         sXml = .ReadText
18380         .Close
18390     End With
18400     docLoadXML sXml
18410     mDocModified = False
18420     Exit Sub
errHandler:
18430     ErrorIn "ujXML.docReadFromFile(FilePath,CharSet)", Array(FilePath, CharSet)
End Sub

Public Sub docWriteToFile(ByVal FilePath As String, Optional ByVal CurrElementOnly As Boolean = False, _
                Optional ByVal CharSet As String = "UNICODE", Optional pHeading As String, Optional IncludeBOM As Boolean = True)
18440     On Error GoTo errHandler

          Dim s As Object
18450     Set s = CreateObject("ADODB.Stream")
          Dim sNoBOM As Object
18460     Set sNoBOM = CreateObject("ADODB.Stream")
18470     With s
18480         .Type = adTypeText
18490         If CharSet <> "" Then .CharSet = CharSet
18500         .Open
18510         If IncludeBOM = False Then
18520             If CurrElementOnly Then .WriteText elXML Else .WriteText pHeading & docXML
                  'Skip over UTF-8 BOM.
18530             .Position = 0 'Enable Type change.
18540             .Type = adTypeBinary
18550             .Position = 3 'UTF-8 BOM is in positions 0 through 2.
18560             With sNoBOM
18570               .Open
18580               .Type = adTypeBinary
18590               s.CopyTo sNoBOM
18600               .SaveToFile FilePath, adSaveCreateOverWrite
18610               .Close
18620             End With
              '    .Close
18630         Else
18640             If CurrElementOnly Then .WriteText elXML Else .WriteText pHeading & docXML
18650             .SaveToFile FilePath, 2 'adSaveCreateOverWrite
18660             .Close
                  
18670         End If
18680     End With
18690     Exit Sub
errHandler:
18700    ' ErrorIn "ujXML.docWriteToFile(FilePath,CurrElementOnly,Charset)", Array(FilePath, CurrElementOnly, CharSet)
End Sub

Public Sub docAFormat(ByVal AttrList As String, ByVal Fmt As String)
18710     On Error GoTo errHandler
          Dim Nodes As MSXML2.IXMLDOMNodeList, Nd As MSXML2.IXMLDOMNode, d As Date
18720     Set Nodes = mTopElem.SelectNodes(".//(@" & Replace(AttrList, ",", "|@") & ")")
18730     For Each Nd In Nodes
18740         If IsNull(Nd.NodeValue) Then
                  'Nd.nodeValue = Null
18750             Else
18760             Nd.NodeValue = uSafeFormat(Nd.NodeValue, Fmt)
18770         End If 'Nd.nodeType ...
18780     Next Nd
18790     Exit Sub
errHandler:
18800     ErrorIn "ujXML.docAFormat(AttrList,Fmt)", Array(AttrList, Fmt)
End Sub

Public Sub docAFormatDates(ByVal AttrList As String, ByVal Fmt As String, Optional ByVal ZeroAsEmpty As Boolean = True)
18810     On Error GoTo errHandler
          Dim Nodes As MSXML2.IXMLDOMNodeList, Nd As MSXML2.IXMLDOMNode, d As Date
18820     Set Nodes = mTopElem.SelectNodes(".//(@" & Replace(AttrList, ",", "|@") & ")")
18830     For Each Nd In Nodes
18840         d = uSafeDate(Nd.NodeValue)
18850         If d = 0 And ZeroAsEmpty Then
18860             Nd.NodeValue = ""
18870             Else
18880             Nd.NodeValue = uSafeFormat(d, Fmt)
18890         End If
18900     Next Nd
18910     Exit Sub
errHandler:
18920     ErrorIn "ujXML.docAFormatDates(AttrList,Fmt,ZeroAsEmpty)", Array(AttrList, Fmt, ZeroAsEmpty)
End Sub

Public Sub docAppendXML(ByVal xmlData As String, Optional ByVal bkmkAs As String)
18930     On Error GoTo errHandler
          Dim xTmp As ujXML
18940     Set xTmp = New ujXML
18950     xTmp.docProgID = docProgID
18960     xTmp.docLoadXML xmlData
18970     docRootElem.appendChild xTmp.Element
18980     Set mElem = xTmp.Element
18990     Exit Sub
errHandler:
19000     ErrorIn "ujXML.docAppendXML(xmlData,bkmkAs)", Array(xmlData, bkmkAs)
End Sub

Public Function docXslLocate(ByVal XslPath As String, Optional ByVal Params, _
        Optional ByVal FromTop As Boolean = True) As Boolean
19010     On Error GoTo errHandler
          Dim tEl As IXMLDOMElement, Nd As IXMLDOMNode
19020     Set tEl = IIf(FromTop, mElem, docRootElem)
19030     If Not IsMissing(Params) Then XslPath = uInsInto(XslPath, Params, False)
19040     Set Nd = tEl.SelectSingleNode(XslPath)
19050     If Nd Is Nothing Then Exit Function
19060     Set mElem = Nd
19070     docXslLocate = True
19080     Exit Function
errHandler:
19090     ErrorIn "ujXML.docXslLocate(XslPath,Params,FromTop)", Array(XslPath, Params, _
               FromTop)
End Function

Public Sub docPersistPrepare()
19100     On Error GoTo errHandler
19110     bkmkSaveToA
19120     docTopSetAttr S_DICT, mDict
19130     Exit Sub
errHandler:
19140     ErrorIn "ujXML.docPersistPrepare"
End Sub

Public Sub docPersistRestore()
19150     On Error GoTo errHandler
19160     Set mhistColl = Nothing
19170     Set mbkmkColl = Nothing
19180     Set mSavedEl = Nothing
19190     mBkmkList = ""
19200     bkmkRestoreFromA
19210     mDict = docTopGetAttr(S_DICT)
19220     Exit Sub
errHandler:
19230     ErrorIn "ujXML.docPersistRestore"
End Sub

Public Sub docPersistCleanup()
19240     On Error GoTo errHandler
19250     bkmkClearA
19260     mTopElem.removeAttribute S_DICT
19270     Exit Sub
errHandler:
19280     ErrorIn "ujXML.docPersistClear"
End Sub

Public Function docTopGetAttr(ByVal AttrName As String) As Variant
19290     On Error GoTo errHandler
19300     docTopGetAttr = mTopElem.getAttribute(AttrName)
19310     Exit Function
errHandler:
19320     ErrorIn "ujXML.docTopGetAttr(AttrName)", AttrName
End Function

Public Sub docTopSetAttr(ByVal AttrName As String, ByVal Value)
19330     On Error GoTo errHandler
19340     mTopElem.setAttribute AttrName, Value
19350     Exit Sub
errHandler:
19360     ErrorIn "ujXML.docTopSetAttr(AttrName,Value)", Array(AttrName, Value)
End Sub

Public Function docCreateViewer(Optional ByVal UseCurrentAsTop As Boolean = True, _
            Optional ByVal ADict As String, Optional AnOptions) As ujXML
19370     On Error GoTo errHandler
          Dim x As ujXML
19380     Set x = New ujXML
19390     With x
19400         Set .docObject = docObject
19410         Set .docTopElem = IIf(UseCurrentAsTop, mElem, docTopElem)
19420         Set .Element = .docTopElem
19430         .docSetOptions IIf(IsMissing(AnOptions), mOptions, AnOptions)
19440         If ADict <> "" Then
19450             .dictList = ADict
19460         ElseIf .attrS(S_DICT) <> "" Then
19470             .dictList = .attrS(S_DICT)
19480         Else
19490             .dictList = dictList
19500         End If
19510         .docOptions(XO_ISVIEWER) = True
19520         .docModified = False
19530     End With 'x
19540     Set docCreateViewer = x
19550     Exit Function
errHandler:
19560     ErrorIn "ujXML.docCreateViewer(UseCurrentAsTop,ADict,AnOptions)", Array(UseCurrentAsTop, ADict, _
               AnOptions)
End Function

Public Property Get docIsViewer() As Boolean
19570     docIsViewer = docOptions(XO_ISVIEWER)
End Property

Public Property Get docTopElem() As IXMLDOMElement
19580     Set docTopElem = mTopElem
End Property

Public Property Set docTopElem(ByVal AnElem As IXMLDOMElement)
19590     Set mTopElem = AnElem
End Property

Public Sub docAfterUpdate(Optional ByVal DontDeleteElems As String = "")
19600     On Error GoTo errHandler
          Dim Ndl As IXMLDOMNodeList, i As Long
19610     navTop
19620     Set Ndl = docTopElem.SelectNodes(".//*[@" & S_STATUS & "]")
19630     For i = 0 To Ndl.Length - 1
19640         Set mElem = Ndl.Item(i)
19650         Select Case elStatus
                  Case XES_DELETED:  If uInList(DontDeleteElems, elName) Then elStatus = XES_JOINED Else elRemove
19660             Case XES_MODIFIED, XES_NEW: elStatus = XES_EXISTING
19670             Case XES_EXISTING, XES_JOINED: 'Nothing to do
19680         End Select
19690     Next i
19700     navTop
19710     Exit Sub
errHandler:
19720     ErrorIn "ujXML.docAfterUpdate(DontDeleteElems)", DontDeleteElems
End Sub

Public Sub docCreateElem(ByVal ElemName As String, _
        Optional ByVal BeforeEl As IXMLDOMElement, _
        Optional ByVal AfterEl As IXMLDOMElement, _
        Optional ByVal AsChildOf As IXMLDOMElement)
19730     On Error GoTo errHandler
          Dim NewEl As IXMLDOMElement
19740     Set NewEl = mDoc.createElement(ElemName)
19750     Check Not ((BeforeEl Is docRootElem) Or (AfterEl Is docRootElem)), X_UJXML, _
              "Before/After reference may not be document root."
19760     If AfterEl Is Nothing Then
19770         If BeforeEl Is Nothing Then
19780             Check Not (AsChildOf Is Nothing), X_UJXML, "At least one reference mElement must be provided."
19790             AsChildOf.appendChild NewEl
19800             Else
19810             BeforeEl.ParentNode.InsertBefore NewEl, BeforeEl
19820         End If
19830       Else
19840         If AfterEl.NextSibling Is Nothing Then
19850             AfterEl.ParentNode.appendChild NewEl
19860             Else
19870             AfterEl.ParentNode.InsertBefore NewEl, AfterEl.NextSibling
19880         End If
19890     End If
19900     Set mElem = NewEl
19910     Exit Sub
errHandler:
19920     ErrorIn "ujXML.docCreateElem(ElemName,BeforeEl,AfterEl,AsChildOf)", Array(ElemName, BeforeEl, _
               AfterEl, AsChildOf)
End Sub

'########################################### Bookmarks #########################################
Public Property Get bkmkColl() As Collection
19930     If mbkmkColl Is Nothing Then Set mbkmkColl = New Collection
19940     Set bkmkColl = mbkmkColl
End Property

Public Sub bkmkSet(ByVal ABkmk As String)
19950     On Error GoTo errHandler
19960     If bkmkExists(ABkmk) Then bkmkRemove ABkmk
19970     bkmkColl.Add mElem, ABkmk
19980     mBkmkList = uListAdd(mBkmkList, ABkmk)
19990     Exit Sub
errHandler:
20000     ErrorIn "ujXML.bkmkSet(ABkmk)", ABkmk
End Sub

Public Sub bkmkClearAll()
20010     On Error Resume Next
20020     Set mbkmkColl = Nothing
20030     mBkmkList = ""
End Sub

Public Function bkmkRemove(ByVal ABkmk As String) As Boolean
20040     On Error Resume Next
20050     bkmkColl.Remove ABkmk
20060     mBkmkList = uListRemove(mBkmkList, ABkmk)
End Function

Private Function bkmkRemoveEl(ByVal El As IXMLDOMElement) As Boolean
20070     On Error GoTo errHandler
          Dim i As Long
20080     With bkmkColl
20090         For i = .Count To 1 Step -1
20100             If uIsChildOf(.Item(i), El) Then .Remove i
20110         Next i
20120     End With
20130     Exit Function
errHandler:
20140     ErrorIn "ujXML.bkmkRemoveEl(El)", El
End Function

Private Function bkmkGetElem(ByVal ABkmk As String) As MSXML2.IXMLDOMElement
20150     On Error Resume Next
20160     Set bkmkGetElem = bkmkColl(ABkmk)
End Function

Public Function bkmkExists(ByVal ABkmk As String) As Boolean
20170     On Error GoTo errHandler
20180     bkmkExists = uInList(mBkmkList, ABkmk)
20190     Exit Function
errHandler:
20200     ErrorIn "ujXML.bkmkExists(ABkmk)", ABkmk
End Function

Public Function bkmkGoto(ByVal ABkmk As String, Optional ByVal Raise As Boolean = True) As Boolean
20210     On Error GoTo errHandler
          Dim tEl As MSXML2.IXMLDOMElement
20220     Set tEl = bkmkGetElem(ABkmk)
20230     If tEl Is Nothing Then
20240         Check Not Raise, X_UJXML, "Bookmark [%1] does not exist.", ABkmk
20250         Else
20260         Set mElem = tEl
20270         bkmkGoto = True
20280     End If
20290     Exit Function
errHandler:
20300     ErrorIn "ujXML.bkmkGoto(ABkmk,Raise)", Array(ABkmk, Raise)
End Function

Private Sub bkmkSaveToA()
20310     On Error GoTo errHandler
          Dim arrB() As String, i As Long
20320     arrB = Split(mBkmkList, ",")
20330     For i = LBound(arrB) To UBound(arrB)
20340         If bkmkExists(arrB(i)) Then
20350             bkmkGoto arrB(i)
20360             attrS(S_BKMK) = uListAdd(attrS(S_BKMK), arrB(i))
20370         End If
20380     Next i
20390     Exit Sub
errHandler:
20400     ErrorIn "ujXML.bkmkSaveToA"
End Sub

Private Sub bkmkClearA()
20410     On Error GoTo errHandler
          Dim Ndl As MSXML2.IXMLDOMNodeList, tEl As MSXML2.IXMLDOMElement
20420     Set Ndl = docTopElem.SelectNodes("//*[@" & S_BKMK & "]")
20430     For Each tEl In Ndl
20440         tEl.removeAttribute S_BKMK
20450     Next tEl
20460     Exit Sub
errHandler:
20470     ErrorIn "ujXML.bkmkClearA"
End Sub

Private Sub bkmkRestoreFromA()
20480     On Error GoTo errHandler
20490     Set mbkmkColl = Nothing
20500     mBkmkList = ""
20510     histPush
          Dim Ndl As MSXML2.IXMLDOMNodeList, tEl As MSXML2.IXMLDOMElement, arrB() As String, i As Long
20520     Set Ndl = docTopElem.SelectNodes("//*[@" & S_BKMK & "]")
20530     For Each tEl In Ndl
20540         Set mElem = tEl
20550         arrB = Split(attrS(S_BKMK), ",")
20560         For i = LBound(arrB) To UBound(arrB)
20570             If arrB(i) <> "" Then
20580                 bkmkSet arrB(i)
20590             End If
20600         Next i
20610     Next tEl
20620     histGoBack
20630     Exit Sub
errHandler:
20640     ErrorIn "ujXML.bkmkRestoreFromA"
End Sub



'History =================================================================================================
Private Function histColl() As Collection
20650     On Error GoTo errHandler
20660     If mhistColl Is Nothing Then Set mhistColl = New Collection
20670     Set histColl = mhistColl
20680     Exit Function
errHandler:
20690     ErrorIn "ujXML.histColl"
End Function

Public Sub histPush()
20700     On Error GoTo errHandler
20710     histColl.Add mElem
20720     Exit Sub
errHandler:
20730     ErrorIn "ujXML.elPush"
End Sub

Public Property Get histCanGoBack() As Boolean
20740     On Error GoTo errHandler
20750     histCanGoBack = (histColl.Count > 0)
20760     Exit Property
errHandler:
20770     ErrorIn "ujXML.histCanGoBack"
End Property

Public Sub histGoBack(Optional ByVal Purge As Boolean = True)
20780     On Error GoTo errHandler
20790     Check histCanGoBack, X_UJXML, "History is empty."
20800     With histColl
20810         Set mElem = .Item(.Count)
20820     End With
20830     If Purge Then histPurge
20840     Exit Sub
errHandler:
20850     ErrorIn "ujXML.histGoBack(Purge)", Purge
End Sub

Public Sub histPurge()
20860     On Error GoTo errHandler
20870     Check histCanGoBack, X_UJXML, "History is empty."
20880     With histColl
20890         .Remove .Count
20900     End With
20910     Exit Sub
errHandler:
20920     ErrorIn "ujXML.histPurge"
End Sub

Public Sub histClear()
20930     Set mhistColl = New Collection
End Sub

Private Function histRemoveEl(ByVal El As IXMLDOMElement) As Boolean
20940     On Error GoTo errHandler
          Dim i As Long
20950     With histColl
20960         For i = .Count To 1 Step -1
20970             If uIsChildOf(.Item(i), El) Then .Remove i
20980         Next i
20990     End With
21000     Exit Function
errHandler:
21010     ErrorIn "ujXML.histRemoveEl(El)", El
End Function

'################################################# Dictionary ##############################################
Public Property Get dictList() As String
21020     dictList = mDict
End Property

Public Property Let dictList(ByVal AList As String)
21030     mDict = AList
21040     docOptions(XO_USE_DICT) = (mDict <> "")
End Property

Public Function dictCheckName(ByVal AName As String, _
        Optional Raise As Boolean = True) As Boolean
21050     On Error GoTo errHandler
21060     If uInList(mDict, AName) Then
21070         dictCheckName = True
21080         Else
21090         Check Not Raise, X_UJXML, "Name [%1] is not in XML dictionary; Current mElement: %2.", _
                  Array(AName, elName)
21100     End If
21110     Exit Function
errHandler:
21120     ErrorIn "ujXML.dictCheckName(AName,Raise)", Array(AName, Raise), , , "Dictionary", mDict
End Function

Public Function dictCheckDoc() As String
21130     On Error GoTo errHandler
          Dim Ndl As IXMLDOMNodeList, i As Long, NdName As String, sRes As String
21140     Set Ndl = mTopElem.SelectNodes("//@*")
21150     For i = 0 To Ndl.Length - 1
21160         NdName = Ndl.Item(i).nodeName
21170         If Not uInList(mDict, NdName) Then sRes = uListAdd(sRes, NdName)
21180     Next i
21190     dictCheckDoc = sRes
21200     Exit Function
errHandler:
21210     ErrorIn "ujXML.dictCheckDoc"
End Function

Public Sub dictAdd(ByVal NameList As String)
21220     On Error GoTo errHandler
          Dim arrL() As String, i As Long
21230     arrL = Split(NameList, ",")
21240     For i = 0 To UBound(arrL)
21250         If Not uInList(mDict, arrL(i)) Then mDict = uListAdd(mDict, arrL(i))
21260     Next i
21270     docOptions(XO_USE_DICT) = (mDict <> "")
21280     Exit Sub
errHandler:
21290     ErrorIn "ujXML.dictAdd(NameList)", NameList
End Sub

Public Sub dictRemove(ByVal NameList As String)
21300     On Error GoTo errHandler
          Dim arrL() As String, i As Long
21310     arrL = Split(NameList, ",")
21320     For i = 0 To UBound(arrL)
21330         If uInList(mDict, arrL(i)) Then mDict = uListRemove(mDict, arrL(i))
21340     Next i
21350     docOptions(XO_USE_DICT) = (mDict <> "")
21360     Exit Sub
errHandler:
21370     ErrorIn "ujXML.dictRemove(NameList)", NameList
End Sub

'Table ====================================================================================================
Public Sub tblAdd(ByVal TblBkmk As String, ByVal xmlList As String)
21380     On Error GoTo errHandler
21390     docAppendXML xmlList, TblBkmk
21400     bkmkSet TblBkmk
21410     Exit Sub
errHandler:
21420     ErrorIn "ujXML.tblAdd(TblBkmk,xmlList)", Array(TblBkmk, xmlList)
End Sub

Public Sub tblRemove(ByVal TblBkmk As String)
21430     On Error GoTo errHandler
21440     bkmkGoto TblBkmk
21450     elRemove
21460     Exit Sub
errHandler:
21470     ErrorIn "ujXML.tblRemove(TblBkmk)", TblBkmk
End Sub

Public Sub tblLocate(ByVal TblBkmk As String)
21480     On Error GoTo errHandler
21490     bkmkGoto TblBkmk
21500     Exit Sub
errHandler:
21510     ErrorIn "ujXML.tblLocate(TblBkmk)", TblBkmk
End Sub

Public Function tblLookup(ByVal TblBkmk As String, _
        ByVal KeyAtrs As String, ByVal KeyValues, _
        ByVal ReturnAttr As String) As String
21520     On Error GoTo errHandler
21530     Check tblLocateElem(TblBkmk, KeyAtrs, KeyValues), X_UJXML, "Lookup failed."
21540     tblLookup = attrS(ReturnAttr)
21550     Exit Function
errHandler:
21560     ErrorIn "ujXML.tblLookup(TblBkmk,KeyAtrs,KeyValues,ReturnAttr)", Array(TblBkmk, KeyAtrs, _
               KeyValues, ReturnAttr)
End Function

Public Function tblLocateElem(ByVal TblBkmk As String, ByVal SrchAttrList As String, ByVal SrchValues) As Boolean
21570     On Error GoTo errHandler
21580     tblLocate TblBkmk
21590     tblLocateElem = chLocate(, SrchAttrList, SrchValues)
21600     Exit Function
errHandler:
21610     ErrorIn "ujXML.tblLocateElem(TblBkmk,SrchAttrList,SrchValues)", Array(TblBkmk, SrchAttrList, _
               SrchValues)
End Function

Public Sub tblLoadTables(ByVal Connection, ByVal TblBkmks As String, _
            Optional ByVal SQLs, Optional ByVal ExecInBatch As Boolean = True)
21620     On Error GoTo errHandler
          Dim Conn As Object, SQL As String, sBuf As String
          Dim arrTN, arrDTN, Ubnd As Long, i As Long, rs As Object
21630     arrTN = Split(TblBkmks, ",")
21640     Ubnd = UBound(arrTN)
21650     If IsMissing(SQLs) Then
21660         SQLs = uCreateArr(Ubnd)
21670         For i = 0 To Ubnd
21680            SQLs(i) = uInsInto(SQL_SELECT, arrTN(i))
21690         Next i
21700       Else
21710         If Not IsArray(SQLs) Then SQLs = Array(SQLs)
21720         Check UBound(SQLs) = Ubnd, X_UJXML, "List counts mismatch (tables and SQLs)"
21730     End If 'IsMissing...
21740     If IsObject(Connection) Then
21750         Set Conn = Connection
21760         Else
21770         Set Conn = CreateObject("ADODB.Connection")
21780         Conn.Open Connection
21790     End If
21800     If ExecInBatch Then
21810         Set rs = Conn.execute(Join(SQLs, "; " & vbNewLine))
21820         For i = 0 To Ubnd
21830             sBuf = uExtractRecords(rs, "Item", arrTN(i))
21840             tblAdd arrTN(i), sBuf
21850             Set rs = rs.NextRecordset
21860         Next i
21870       Else
21880         For i = 0 To Ubnd
21890             Set rs = Conn.execute(SQLs(i))
21900             sBuf = uExtractRecords(rs, "Item", arrTN(i))
21910             tblAdd arrTN(i), sBuf
21920         Next i
21930     End If
21940     Exit Sub
errHandler:
21950     ErrorIn "ujXML.tblLoadTables(Connection,TblBkmks,SQLs,ExecInBatch)", Array(Connection, TblBkmks, _
               SQLs, ExecInBatch)
End Sub


'Transactions ================================================================================================
Public Sub tranBegin()
21960     On Error GoTo errHandler
21970     If docIsViewer Then Exit Sub
21980     docPersistPrepare
21990     navTop
22000     mtranSavedXML = elXML
22010     docPersistCleanup
22020     Exit Sub
errHandler:
22030     ErrorIn "ujXML.tranBegin"
End Sub

Public Sub tranCommit()
22040     On Error GoTo errHandler
22050     If docIsViewer Then Exit Sub
22060     mtranSavedXML = ""
22070     Exit Sub
errHandler:
22080     ErrorIn "ujXML.tranCommit"
End Sub

Public Sub tranRollback()
22090     On Error GoTo errHandler
22100     If docIsViewer Then Exit Sub
22110     Check mtranSavedXML <> "", X_UJXML, "No transaction in progress."
22120     docLoadXML mtranSavedXML
22130     docPersistRestore
22140     docPersistCleanup
22150     mtranSavedXML = ""
22160     navTop
22170     Exit Sub
errHandler:
22180     ErrorIn "ujXML.tranRollback"
End Sub

Public Property Get tranSavedXML() As String
22190     tranSavedXML = mtranSavedXML
End Property

'HTML ======================================================================================
Public Function htmElConstruct(ByVal Template As String, ByVal AttrList As String, _
        Optional ByVal SendToIISResponse As Boolean = False, _
        Optional ByVal CtrlChar As String = "%") As String
22200     On Error GoTo errHandler
          Dim arrA() As String, Res As String
22210     Template = Replace(Template, CtrlChar, uChrBk)
22220     arrA = Split(AttrList, ",")
22230     Res = htmElConstrPriv(Template, arrA)
22240     Res = Replace(Res, uChrBk, CtrlChar)
22250     htmElConstruct = Res
22260     If SendToIISResponse And iisInIIS Then iisWrite Res
22270     Exit Function
errHandler:
22280     ErrorIn "ujXML.htmElConstruct(Template,AttrList,CtrlChar)", Array(Template, AttrList, CtrlChar)
End Function

Public Function htmChConstruct(ByVal Template As String, ByVal AttrList As String, _
        Optional ByVal SendToIISResponse As Boolean = False, _
        Optional ByVal SelAttr As String, Optional SelValues As String, _
        Optional ShowSelAs As String = "SELECTED", Optional ByVal CtrlChar As String = "%", _
        Optional ByVal Crlf As String = "||", Optional ByVal Indent As Long = 8) As String
22290     On Error GoTo errHandler
          Dim arrA() As String, Res As String, i As Long
          Dim sTempl As String, s As String, sSel As String, sIndent As String
22300     If Indent > 0 Then
22310         sIndent = Space(Indent)
22320         sTempl = sIndent & sTempl
22330     End If
22340     sTempl = Template
22350     If Crlf <> "" Then sTempl = Replace(sTempl, Crlf, sIndent & vbNewLine)
22360     sTempl = Replace(sTempl, CtrlChar, uChrBk)
22370     arrA = Split(AttrList, ",")
22380     If navFirstChild Then
22390         Do
22400             s = htmElConstrPriv(sTempl, arrA)
22410             If SelAttr <> "" Then
22420                 sSel = IIf(uInList(SelValues, attrS(SelAttr)), ShowSelAs, "")
22430                 s = Replace(s, uChrBk & "0", sSel)
22440             End If
22450             If Res <> "" Then Res = Res & vbNewLine
22460             Res = Res & Space(Indent) & s
22470         Loop While navNext
22480     End If
22490     Res = Replace(Res, uChrBk, CtrlChar)
22500     htmChConstruct = Res
22510     If SendToIISResponse And iisInIIS Then iisWrite Res
22520     Exit Function
errHandler:
22530     ErrorIn "ujXML.htmChConstruct(Template,AttrList,SendToIISResponse,SelAttr,SelValues,ShowSelAs," & _
              "CtrlChar,Crlf,Indent)", Array(Template, AttrList, SendToIISResponse, SelAttr, SelValues, ShowSelAs, _
               CtrlChar, Crlf, Indent)
End Function

Private Function htmElConstrPriv(ByRef Template As String, ByRef arrANms) As String
22540     On Error GoTo errHandler
          Dim i As Long, sV As String, Res As String
22550     Res = Template
22560     For i = 0 To UBound(arrANms)
22570         sV = attrEscHtm(arrANms(i))
22580         Res = Replace(Res, uChrBk & (i + 1), sV)
22590     Next i
22600     htmElConstrPriv = Res
22610     Exit Function
errHandler:
22620     ErrorIn "ujXML.htmElConstrPriv(Template,arrANms)", Array(Template, arrANms)
End Function

'SQL ====================================================================================
Public Function sqlElConstruct(ByVal TableName As String, _
            ByVal FldEqAttrList As String, ByVal KeyFldEqAttrList As String, _
            Optional ByVal SQLType As XENUM_SQLTYPE = XSQL_AUTO) As String
22630     On Error GoTo errHandler
          Dim Ubnd As Long, sRes As String
          Dim arrF, arrA, arrV, arrQ
          Dim arrKF, arrKA, arrKV, arrKQ
22640     If SQLType = XSQL_AUTO Then SQLType = sqlElGetDftSQLType
22650     Ubnd = uUnpackEQList(FldEqAttrList, arrF, arrA, arrQ)
22660     arrV = uCreateArr(Ubnd)
22670     Ubnd = uUnpackEQList(KeyFldEqAttrList, arrKF, arrKA, arrKQ)
22680     arrKV = uCreateArr(Ubnd)
22690     sqlElFillArray arrA, arrQ, arrV
22700     sqlElFillArray arrKA, arrKQ, arrKV
22710     sRes = sqlElConstructPriv(TableName, SQLType, arrF, arrA, arrQ, arrV, arrKF, arrKA, arrKQ, arrKV)
22720     sqlElConstruct = sRes
22730     Exit Function
errHandler:
22740     ErrorIn "ujXML.sqlElConstruct(TableName,FldEqAttrList,KeyFldEqAttrList,SQLType)", Array(TableName, _
               FldEqAttrList, KeyFldEqAttrList, SQLType)
End Function

Public Function sqlChConstruct(ByVal TableName As String, _
            ByVal FldEqAttrList As String, ByVal KeyFldEqAttrList As String, _
            Optional ByVal CascDelTableNames As String, _
            Optional ByVal CascKeyFldEqAttrList As String) As Variant
22750     On Error GoTo errHandler
          Dim Ubnd As Long
          Dim arrF, arrA, arrV, arrQ
          Dim arrKF, arrKA, arrKV, arrKQ
          Dim arrDelF, arrDelA, arrDelV, arrDelQ, arrDelTables
          Dim colRes As Collection, arrRes, i As Long, sSQL As String, SType As XENUM_SQLTYPE, vE
          'Prepare arrays
22760     Set colRes = New Collection
22770     Ubnd = uUnpackEQList(FldEqAttrList, arrF, arrA, arrQ)
22780     arrV = uCreateArr(Ubnd)
22790     Ubnd = uUnpackEQList(KeyFldEqAttrList, arrKF, arrKA, arrKQ)
22800     arrKV = uCreateArr(Ubnd)
22810     If CascDelTableNames <> "" Then
22820         arrDelTables = Split(CascDelTableNames, ",")
22830         If CascKeyFldEqAttrList = "" Then CascKeyFldEqAttrList = KeyFldEqAttrList
22840         Ubnd = uUnpackEQList(CascKeyFldEqAttrList, arrDelF, arrDelA, arrDelQ)
22850         arrDelV = uCreateArr(Ubnd)
22860     End If
          'Now create SQLs
22870     If navFirstChild Then
22880         Do
22890             sqlElFillArray arrA, arrQ, arrV
22900             sqlElFillArray arrKA, arrKQ, arrKV
22910             SType = sqlElGetDftSQLType
22920             If SType = XSQL_DELETE And CascDelTableNames <> "" Then
22930                 sqlElFillArray arrDelA, arrDelQ, arrDelV
22940                 For i = 0 To UBound(arrDelTables)
22950                     sSQL = sqlElConstructPriv(arrDelTables(i), XSQL_DELETE, arrF, arrA, arrQ, arrV, arrDelF, arrDelA, arrDelQ, arrDelV)
22960                     colRes.Add sSQL
22970                 Next i
22980             End If
22990             sSQL = sqlElConstructPriv(TableName, SType, arrF, arrA, arrQ, arrV, arrKF, arrKA, arrKQ, arrKV)
23000             If sSQL <> "" Then colRes.Add sSQL
23010         Loop While navNext
23020     End If
23030     If colRes.Count = 0 Then Exit Function
23040     ReDim arrRes(0 To colRes.Count - 1)
23050     For i = 0 To colRes.Count - 1
23060         arrRes(i) = colRes(i + 1)
23070     Next i
23080     sqlChConstruct = arrRes
23090     Exit Function
errHandler:
23100     ErrorIn "ujXML.sqlChConstruct(TableName,FldEqAttrList,KeyFldEqAttrList)", Array(TableName, FldEqAttrList, _
               KeyFldEqAttrList)
End Function

Private Function sqlElConstructPriv(ByVal TableName As String, ByVal SQLType As Long, _
                ByRef arrF, ByRef arrA, ByRef arrQ, ByRef arrVbuf, _
                ByRef arrKF, ByRef arrKA, ByRef arrKQ, ByRef arrKVbuf) As String
23110     On Error GoTo errHandler
      '    Const _
      '        SQL_INSERT = "INSERT INTO  %1 (%2) VALUES (%3);", _
      '        SQL_UPDATE = "UPDATE  %1 SET %2 WHERE %3;", _
      '        SQL_DELETE = "DELETE FROM %1 WHERE %2;", _
      '        SQL_FIND = "SELECT %2 FROM %1 WHERE %2;"
23120     Select Case SQLType
              Case XSQL_UPDATE
23130             sqlElFillArray arrA, arrQ, arrVbuf
23140             sqlElFillArray arrKA, arrKQ, arrKVbuf
23150             sqlElConstructPriv = uInsInto(SQL_UPDATE, Array(vbNewLine, TableName, uCreateEQList(arrF, arrVbuf), uCreateEQList(arrKF, arrKVbuf, " AND ")))
23160         Case XSQL_INSERT
23170             sqlElFillArray arrA, arrQ, arrVbuf
23180             sqlElFillArray arrKA, arrKQ, arrKVbuf
23190             sqlElConstructPriv = uInsInto(SQL_INSERT, Array(vbNewLine, TableName, Join(arrF, ", "), Join(arrVbuf, ", ")))
23200         Case XSQL_DELETE
23210             sqlElFillArray arrKA, arrKQ, arrKVbuf
23220             sqlElConstructPriv = uInsInto(SQL_DELETE, Array(vbNewLine, TableName, uCreateEQList(arrKF, arrKVbuf, " AND ")))
23230     End Select
23240     Exit Function
errHandler:
23250     ErrorIn "ujXML.sqlElConstructPriv(TableName,SQLType,arrF,arrA,arrQ,arrVbuf,arrKF,arrKA,arrKQ," & _
              "arrKVbuf)", Array(TableName, SQLType, arrF, arrA, arrQ, arrVbuf, arrKF, arrKA, arrKQ, arrKVbuf)
End Function

Private Sub sqlElFillArray(ByRef arrA, ByRef arrQ, ByRef arrV)
23260     On Error GoTo errHandler
          Dim i As Long, SA As String
23270     For i = LBound(arrA) To UBound(arrA)
23280         SA = arrA(i)
23290         If attrIsNull(SA) Then
23300             arrV(i) = "Null"
23310         ElseIf arrQ(i) Then
23320             arrV(i) = attrQt(SA)
23330         Else
23340             arrV(i) = attrS(SA)
23350         End If
23360     Next i
23370     Exit Sub
errHandler:
23380     ErrorIn "ujXML.sqlElFillArray(arrA,arrQ,arrV)", Array(arrA, arrQ, arrV)
End Sub

Public Function sqlElGetDftSQLType() As XENUM_SQLTYPE
23390     On Error GoTo errHandler
          Dim T As Long
23400     Select Case elStatus
              Case XES_JOINED, XES_EXISTING: T = XSQL_NONE
23410         Case XES_MODIFIED: T = XSQL_UPDATE
23420         Case XES_NEW: T = XSQL_INSERT
23430         Case XES_DELETED: T = XSQL_DELETE
23440     End Select
23450     sqlElGetDftSQLType = T
23460     Exit Function
errHandler:
23470     ErrorIn "ujXML.sqlElGetDftSQLType"
End Function

'Private Function sqlTblConstructForCrossTable(ByVal TblBkmk As String, ByVal CrossTableName As String, _
'            ByVal LkpTableKey As String, ByVal MainTableKey As String, ByVal MainKeyValue, _
'            ByVal OldLkpKeyValues As String, ByVal NewLkpKeyValues As String) As Variant
'    On Error GoTo ErrHandler
'    If TblBkmk <> "" Then tblLocate TblBkmk
'    Check False, X_UJXML, "Method Not implemented."
'    Exit Function
'ErrHandler:
'    ErrorIn "ujXML.sqlTblConstructForCrossTable(TblBkmk,CrossTableName,LkpTableKey,MainTableKey," & _
'        "MainKeyValue,OldLkpKeyValues,NewLkpKeyValues)", Array(TblBkmk, CrossTableName, LkpTableKey, _
'         MainTableKey, MainKeyValue, OldLkpKeyValues, NewLkpKeyValues)
'End Function

'################################# Recordset-related functions ################################################
Public Sub rsElCopy(ByVal AdoRs As Object, _
            Optional ByVal Specs As String, Optional ByVal AddToDict As Boolean = False)
23480     On Error GoTo errHandler
          Dim arrF, arrA, arrQ
23490     If Specs = "" Then Specs = rsGetFieldList(AdoRs)
23500     uUnpackEQList Specs, arrA, arrF, arrQ
23510     If AddToDict Then dictAdd Join(arrA, ",")
23520     rsElCopyPriv AdoRs, arrA, arrF, arrQ
23530     Exit Sub
errHandler:
23540     ErrorIn "ujXML.rsElCopy(AdoRs,Specs,AddToDict)", Array(AdoRs, Specs, AddToDict)
End Sub

Private Sub rsElCopyPriv(ByVal AdoRs As Object, arrA, arrF, arrQ)
23550     On Error GoTo errHandler
          Dim i As Long, fName As String, f As Object
23560     For i = 0 To UBound(arrF)
23570         If arrQ(i) Then
23580             attrV(arrA(i)) = arrF(i)
23590             Else
23600             Set f = uSafeGetFieldObj(AdoRs, arrF(i))
23610             Check Not f Is Nothing, X_UJXML, "Field [%1] not found.", arrF(i)
23620             attrV(arrA(i)) = f.Value
23630         End If
23640     Next i
23650     Exit Sub
errHandler:
23660     ErrorIn "ujXML.rsElCopyPriv(AdoRs,arrA,arrF,arrQ)", Array(AdoRs, arrA, arrF, arrQ)
End Sub

Public Sub rsAddAsChildren(ByVal AdoRs As Object, _
                          Optional ByVal RecElemName As String = "Record", _
                          Optional ByVal AddToDict As Boolean = False)
23670     On Error GoTo errHandler
          Dim s As String, xBuf As ujXML, Ndl As IXMLDOMNodeList, El As IXMLDOMElement, i As Long
23680     If AddToDict Then rsDictAdd AdoRs
23690     s = uExtractRecords(AdoRs, RecElemName)
23700     If s = "" Then Exit Sub
23710     s = uEnclose(s, "Recordset")
23720     Set xBuf = New ujXML
23730     xBuf.docLoadXML s
23740     Set Ndl = xBuf.docRootElem.SelectNodes("*")
23750     For i = 0 To Ndl.Length - 1
23760         Set El = Ndl.Item(i)
23770         mElem.appendChild El
23780     Next i
23790     Exit Sub
errHandler:
23800     ErrorIn "ujXML.rsAddAsChildren(AdoRS,RecElemName,AddToDict)", Array(AdoRs, RecElemName, _
               AddToDict)
End Sub
    
Public Sub rsAddAsChildrenExt(ByVal AdoRs As Object, ByVal RecElemName As String, _
            Optional ByVal Specs As String, Optional ByVal AddToDict As Boolean = False)
23810     On Error GoTo errHandler
          Dim arrF, arrA, arrQ
23820     If Specs = "" Then Specs = rsGetFieldList(AdoRs)
23830     uUnpackEQList Specs, arrA, arrF, arrQ
23840     If AddToDict Then dictAdd Join(arrA, ",")
23850     With AdoRs
23860         If .BOF And .eof Then Exit Sub
23870         .First
23880         While Not .eof
23890             chCreate RecElemName
23900             rsElCopyPriv AdoRs, arrA, arrF, arrQ
23910             navUP
23920             .MoveNext
23930         Wend
23940     End With
23950     Exit Sub
errHandler:
23960     ErrorIn "ujXML.rsAddAsChildrenExt(AdoRS,RecElemName,Specs)", Array(AdoRs, RecElemName, _
               Specs)
End Sub
    
Public Sub rsDictAdd(ByVal AdoRs As Object)
23970     On Error GoTo errHandler
23980     dictAdd rsGetFieldList(AdoRs)
23990     Exit Sub
errHandler:
24000     ErrorIn "ujXML.rsDictAdd(AdoRS)", AdoRs
End Sub
    
Private Function rsGetFieldList(ByVal AdoRs As Object) As String
24010     On Error GoTo errHandler
          Dim f As Object, Res As String
24020     For Each f In AdoRs.Fields
24030         Res = uCombine(Res, ",", f.Name)
24040     Next f
24050     rsGetFieldList = Res
24060     Exit Function
errHandler:
24070     ErrorIn "ujXML.rsGetFieldList(AdoRs)", AdoRs
End Function

'Controls ====================================================================================================
Public Property Get wfCopyDir() As XENUM_COPY_DIR
24080     wfCopyDir = mCopyDir
End Property

Public Property Let wfCopyDir(ByVal Value As XENUM_COPY_DIR)
24090     mCopyDir = Value
End Property

Public Sub wfCopy(ByVal AttrList As String, ByVal arrControls, _
        Optional ByVal ListCopyText As Boolean = False)
24100     On Error GoTo errHandler
          Dim i As Long, arrA() As String
24110     If Not IsArray(arrControls) Then arrControls = Array(arrControls)
24120     arrA = Split(AttrList, ",")
24130     Check UBound(arrA) = UBound(arrControls), X_UJXML, _
              "Count mismatch: AttrList: %1; arrControls: %2.", Array(UBound(arrA) + 1, UBound(arrControls) + 1)
24140     For i = LBound(arrControls) To UBound(arrControls)
24150         wfCopyOne arrA(i), arrControls(i), ListCopyText
24160     Next i
24170     Exit Sub
errHandler:
24180     ErrorIn "ujXML.wfCopy(AttrList,arrControls,ListCopyText)", Array(AttrList, arrControls, _
               ListCopyText)
End Sub

Public Sub wfFill(ByVal ctrl As Object, _
                       ByVal TextAttr As String, Optional ByVal DataAttr As String, _
                       Optional ByVal TblBkmk As String, _
                       Optional ByVal FilterAtrs As String, Optional ByVal FilterValues As String, _
                       Optional ByVal SelDataOrIndexes As String)
24190     On Error GoTo errHandler
          Dim blnAdd As Boolean, lngD As Long, blnSorted As Boolean, Ind As Long
24200     blnSorted = uSafeListSorted(ctrl)
24210     Check Not blnSorted, X_UJXML, "wfFillCtrlCtrl doesn't support SORTED list/combo boxes."
24220     ctrl.Clear
24230     If TblBkmk <> "" Then bkmkGoto TblBkmk
24240     If navFirstChild Then
24250         Do
24260             If FilterAtrs = "" Then
24270                 blnAdd = True
24280                 Else
24290                 blnAdd = elCheckValues(FilterAtrs, FilterValues)
24300             End If
24310             If blnAdd Then
24320                 ctrl.AddItem attrS(TextAttr)
24330                 Ind = ctrl.ListCount - 1
24340                 If DataAttr = "" Then
24350                     If SelDataOrIndexes <> "" Then
24360                         If uInList(SelDataOrIndexes, Ind) Then uListSafeSelect ctrl, Ind, True
24370                     End If 'Not IsMissing(...
24380                     Else
24390                     lngD = attrI(DataAttr)
24400                     ctrl.ItemData(ctrl.ListCount - 1) = lngD
24410                     If SelDataOrIndexes <> "" Then
24420                         If uInList(SelDataOrIndexes, lngD) Then uListSafeSelect ctrl, Ind, True
24430                     End If 'Not IsMissing(...
24440                 End If 'DataAttr ....
24450             End If 'blnAdd ....
24460         Loop While navNext
24470     End If
      '    If Not IsMissing(SelectedIndex) Then
      '        If SelectedIndex <= Ctrl.ListCount - 1 Then Ctrl.ListIndex = SelectedIndex
      '    End If
24480     Exit Sub
errHandler:
24490     ErrorIn "ujXML.wfFill(Ctrl,TextAttr,DataAttr,TblBkmk,FilterAtrs,FilterValues,SelDataOrIndexes)", _
               Array(ctrl, TextAttr, DataAttr, TblBkmk, FilterAtrs, FilterValues, SelDataOrIndexes)
End Sub

Public Function wfSyncWith(ByVal ctrl As Object, Optional ByVal TblBkmk As String, _
        Optional ByVal DataAttr As String) As Boolean
24500     On Error GoTo errHandler
          Dim Ind As Long, d As Long
24510     If TblBkmk <> "" Then bkmkGoto TblBkmk
24520     Ind = ctrl.ListIndex
24530     If Ind < 0 Then
24540         Else
24550         If DataAttr = "" Then
24560             navChildByIndex Ind
24570             wfSyncWith = True
24580             Else
24590             wfSyncWith = chLocate(DataAttr, ctrl.ItemData(Ind))
24600         End If
24610     End If
24620     Exit Function
errHandler:
24630     ErrorIn "ujXML.wfSyncWith(Ctrl,TblBkmk,DataAttr)", Array(ctrl, TblBkmk, DataAttr)
End Function

Private Sub wfCopyOne(ByVal AttrName As String, ByVal ctrl As Object, Optional ByVal ListCopyText As Boolean)
24640     On Error GoTo errHandler
24650     Select Case mCopyDir
              Case XDIR_IN: wfCopyIn AttrName, ctrl, ListCopyText
24660         Case XDIR_OUT: wfCopyOut AttrName, ctrl, ListCopyText
24670         Case XDIR_XMLCLEAR: wfClearXML AttrName, ctrl, ListCopyText
24680         Case XDIR_CTRLCLEAR: wfClearCtrl AttrName, ctrl, ListCopyText
24690     End Select
24700     Exit Sub
errHandler:
24710     ErrorIn "ujXML.wfCopyOne(AttrName,Ctrl,ListCopyText)", Array(AttrName, ctrl, ListCopyText)
End Sub
    
Private Sub wfCopyIn(ByVal AttrName As String, ByVal ctrl As Object, Optional ByVal ListCopyText As Boolean)
24720     On Error GoTo errHandler
24730     Select Case TypeName(ctrl)
              Case "Label": attrS(AttrName) = ctrl.Caption
24740         Case "TextBox": attrS(AttrName) = ctrl.text
24750         Case "ComboBox"
24760             Select Case uSafeCtrlStyle(ctrl)
                      Case 0, 1: attrS(AttrName) = ctrl.text   'combodropdown, simplecombo
24770                 Case 2 'dropdown list
24780                         attrV(AttrName) = uListGetValue(ctrl)
24790                         If ListCopyText Then attrV("_" & AttrName & "_Text") = ctrl.text
24800             End Select
24810         Case "ListBox"
24820             If uSafeListMultiSelect(ctrl) Then
24830                 attrV(AttrName) = uListGetMultiSelect(ctrl)
24840                 Else
24850                 attrV(AttrName) = uListGetValue(ctrl)
24860                 If ListCopyText Then attrV("_" & AttrName & "-Text") = ctrl.text
24870             End If
24880         Case "OptionButton"
24890             If ctrl.Index <> "" Then
24900                 If ctrl.Value Then attrI(AttrName) = ctrl.Index
24910             End If
24920         Case "CheckBox":  attrB(AttrName) = uValueIn(ctrl.Value, vbChecked, True)
24930         Case "FlatScrollBar", "Slider", "ProgressBar", "HScrollBar", "VScrollBar", "ScrollBar"
24940             attrF(AttrName) = ctrl.Value
24950         Case "ToggleButton": attrB(AttrName) = ctrl.Value
24960     End Select
24970     Exit Sub
errHandler:
24980     ErrorIn "ujXML.wfCopyIn(AttrName,Ctrl,ListCopyText)", Array(AttrName, ctrl, ListCopyText)
End Sub

Private Sub wfCopyOut(ByVal AttrName As String, ByVal ctrl As Object, Optional ByVal ListCopyText As Boolean)
24990     On Error GoTo errHandler
25000     Select Case TypeName(ctrl)
              Case "Label":  ctrl.Caption = attrS(AttrName)
25010         Case "TextBox": ctrl.text = attrS(AttrName)
25020         Case "ComboBox"
25030             Select Case uSafeCtrlStyle(ctrl)
                      Case 0, 1: ctrl.text = attrS(AttrName)           'combodropdown, simplecombo
25040                 Case 2: uListSetValue ctrl, attrI(AttrName)      'dropdown list
25050             End Select
25060         Case "ListBox"
25070             If uSafeListMultiSelect(ctrl) Then
25080                 uListSetMultiSelect ctrl, attrS(AttrName)
25090                 Else
25100                 uListSetValue ctrl, attrI(AttrName)
25110             End If
25120         Case "OptionButton"
25130             If ctrl.Index <> "" Then
25140                  ctrl.Value = (ctrl.Index = attrI(AttrName))
25150             End If
25160         Case "CheckBox":        ctrl.Value = IIf(attrB(AttrName), vbChecked, vbUnchecked)
25170         Case "FlatScrollBar", "Slider", "ProgressBar", "HScrollBar", "VScrollBar", "ScrollBar"
25180             ctrl.Value = attrF(AttrName)
25190         Case "ToggleButton": ctrl.Value = attrB(AttrName)
25200     End Select
25210     Exit Sub
errHandler:
25220     ErrorIn "ujXML.wfCopyOut(AttrName,Ctrl,ListCopyText)", Array(AttrName, ctrl, ListCopyText)
End Sub

Private Sub wfClearCtrl(ByVal AttrName As String, ByVal ctrl As Object, Optional ByVal ListCopyText As Boolean = False)
25230     On Error GoTo errHandler
          Dim i As Long
25240     Select Case TypeName(ctrl)
              Case "Label":  ctrl.Caption = ""
25250         Case "TextBox": ctrl.text = ""
25260         Case "ComboBox"
25270             Select Case uSafeCtrlStyle(ctrl)
                      Case 0, 1: ctrl.text = ""           'combodropdown, simplecombo
25280                 Case 2: uListSetValue ctrl, -1      'dropdown list
25290             End Select
25300         Case "ListBox"
25310             If uSafeListMultiSelect(ctrl) Then
25320                 For i = 0 To ctrl.ListCount - 1
25330                     ctrl.Selected(i) = False
25340                 Next i
                      'uListSetValue Ctrl, -1
25350                 Else
25360                 uListSetValue ctrl, -1
25370             End If
25380         Case "OptionButton":  ctrl.Value = False
25390         Case "CheckBox":        ctrl.Value = False
25400         Case "FlatScrollBar", "Slider", "ProgressBar", "HScrollBar", "VScrollBar", "ScrollBar"
25410             ctrl.Value = ctrl.Min
25420         Case "ToggleButton": ctrl.Value = False
25430     End Select
25440     Exit Sub
errHandler:
25450     ErrorIn "ujXML.wfClearCtrl(AttrName,Ctrl,ListCopyText)", Array(AttrName, ctrl, ListCopyText)
End Sub

Private Sub wfClearXML(ByVal AttrName As String, ByVal ctrl As Object, Optional ByVal ListCopyText As Boolean)
25460     On Error GoTo errHandler
25470     Select Case TypeName(ctrl)
              Case "Label": attrS(AttrName) = ctrl.Caption
25480         Case "TextBox": attrS(AttrName) = ctrl.text
25490         Case "ComboBox"
25500             Select Case uSafeCtrlStyle(ctrl)
                      Case 0, 1: attrS(AttrName) = ""   'combodropdown, simplecombo
25510                 Case 2 'dropdown list
25520                         attrV(AttrName) = "-1"
25530                         If ListCopyText Then attrV("_" & AttrName & "_Text") = ""
25540             End Select
25550         Case "ListBox"
25560             If uSafeListMultiSelect(ctrl) Then
25570                 attrV(AttrName) = ""
25580                 Else
25590                 attrV(AttrName) = -1
25600                 If ListCopyText Then attrV("_" & AttrName & "-Text") = ""
25610             End If
25620         Case "OptionButton": attrV(AttrName) = Null
25630         Case "CheckBox":  attrB(AttrName) = False
25640         Case "FlatScrollBar", "Slider", "ProgressBar", "HScrollBar", "VScrollBar", "ScrollBar"
25650             attrF(AttrName) = ctrl.Min
25660         Case "ToggleButton": attrB(AttrName) = False
25670     End Select
25680     Exit Sub
errHandler:
25690     ErrorIn "ujXML.wfClearXML(AttrName,Ctrl,ListCopyText)", Array(AttrName, ctrl, ListCopyText)
End Sub

'IIS =========================================================================================================
Public Property Get iisInIIS() As Boolean
25700     On Error Resume Next
25710     iisInIIS = Not (iisGetObject("Application") Is Nothing)
End Property

Private Function iisCheckor() As Boolean
25720     If Err.Number = 0 Then
25730         iisCheckor = True
25740         Else
25750         miisReport = uCombine(miisReport, vbNewLine & vbNewLine, ErrReport)
25760     End If
End Function

Public Property Get iisReport() As String
25770     iisReport = miisReport
End Property

Public Property Get iisReportHTML() As String
25780     iisReportHTML = "<pre>" & uEscapeHTML(miisReport) & "</pre>"
End Property

Public Sub iisReportClear()
25790     miisReport = ""
End Sub

Public Function iisSaveDocSafe(ByVal DocName As String) As Boolean
25800     On Error Resume Next
25810     iisSaveDoc DocName
25820     iisSaveDocSafe = iisCheckor
End Function

Public Sub iisSaveDoc(ByVal DocName As String)
25830     On Error GoTo errHandler
          Dim iisApp As Object
25840     Set iisApp = iisGetObject("Application")
25850     Check Not (iisApp Is Nothing), X_UJXML, "Failed to retrieve IIS Application object."
25860     docPersistPrepare
25870     With iisApp
25880         .Lock
25890         Set .Contents(DocName) = mDoc
25900         .UnLock
25910     End With
25920     Exit Sub
errHandler:
25930     ErrorIn "ujXML.iisSaveDocPriv(DocName)", DocName
End Sub

Public Function iisClearDocSafe(ByVal DocName As String) As Boolean
25940     On Error Resume Next
25950     iisClearDoc DocName
25960     iisClearDocSafe = iisCheckor
End Function

Public Sub iisClearDoc(ByVal DocName As String)
25970     On Error GoTo errHandler
          Dim iisApp As Object, V
25980     Set iisApp = iisGetObject("Application")
25990     Check Not (iisApp Is Nothing), X_UJXML, "Failed to retrieve IIS Application object."
26000     With iisApp
26010         .Lock
26020         .Contents(DocName) = V
26030         .UnLock
26040     End With
26050     Exit Sub
errHandler:
26060     ErrorIn "ujXML.iisClearDocPriv(DocName)", DocName
End Sub

Public Function iisRestoreDocSafe(ByVal DocName As String) As Boolean
26070     On Error Resume Next
26080     iisRestoreDoc DocName
26090     iisRestoreDocSafe = iisCheckor
End Function

Public Sub iisRestoreDoc(ByVal DocName As String)
26100     On Error GoTo errHandler
          Dim iisApp As Object, obj As Object
26110     docOptions(XO_IIS_RETRIEVED) = False
26120     Set iisApp = iisGetObject("Application")
26130     Check Not (iisApp Is Nothing), X_UJXML, "Failed to retrieve IIS Application object "
26140     Set obj = iisGetAppContObjSafe(iisApp, DocName)
26150     Check Not (obj Is Nothing), X_UJXML, _
                  "Failed to retrieve document from Application.Contents; DocName=[%1]", DocName
26160     Check TypeOf obj Is IXMLDOMDocument, _
                  X_UJXML, "Invalid type of retrieved object, expected IXMLDOMDocument, retrieved [%1]", TypeName(obj)
26170     Set docObject = obj
26180     docPersistRestore
26190     docOptions(XO_IIS_RETRIEVED) = True
26200     Exit Sub
errHandler:
26210     ErrorIn "ujXML.iisRestoreDoc(DocName)", DocName
End Sub

Private Function iisGetAppContObjSafe(ByVal iisApp As Object, ByVal ObjName As String) As Object
26220     On Error Resume Next
26230     Set iisGetAppContObjSafe = iisApp.Contents(ObjName)
End Function

Public Function iisRestoreLoadTablesSafe(ByVal DocName As String, ByVal Connect, _
        ByVal TableNames As String, Optional ByVal SQLs, Optional ByVal ExecInBatch As Boolean, _
        Optional ByVal MaxAgeSec As Double = 0) As Boolean
26240     On Error Resume Next
26250     iisRestoreLoadTables DocName, Connect, TableNames, SQLs, ExecInBatch, MaxAgeSec
26260     iisRestoreLoadTablesSafe = iisCheckor
End Function

Public Sub iisRestoreLoadTables(ByVal DocName As String, ByVal Connect, _
        ByVal TableNames As String, Optional ByVal SQLs, Optional ByVal ExecInBatch As Boolean, _
        Optional ByVal MaxAgeSec As Double = 0)
26270     On Error GoTo errHandler
          Const ONE_SEC As Double = 1# / 86400
          Dim ToReload As Boolean, Age As Double
26280     ToReload = True
26290     If iisRestoreDocSafe(DocName) Then
26300         ToReload = False
26310         If MaxAgeSec > 0 Then
26320             navTop
26330             Age = (Now - attrDt(S_DATETIME)) / ONE_SEC
26340             If Age > MaxAgeSec Then ToReload = True
26350         End If 'ReloadIfOlderThanSec...
26360     End If
26370     docOptions(XO_IIS_RETRIEVED) = Not ToReload
26380     If ToReload Then
26390         If mDocProgID = "" Then mDocProgID = PROGID_FREETHREADEDDOC ' "msxml2.FreeThreadedDOMDocument"
26400         docInit "Tables"
26410         tblLoadTables Connect, TableNames, SQLs, ExecInBatch
26420         navTop
26430         attrDt(S_DATETIME) = Now
26440         iisSaveDoc DocName
26450     End If
26460     Exit Sub
errHandler:
26470     ErrorIn "ujXML.iisRestoreLoadTables(DocName,Connect,TableNames,SQLs,ExecInBatch,MaxAgeSec)", _
               Array(DocName, Connect, TableNames, SQLs, ExecInBatch, MaxAgeSec)
End Sub

Public Function iisCopySafe(Optional ByVal AttrList As String) As Boolean
26480     On Error Resume Next
26490     iisCopy AttrList
26500     iisCopySafe = iisCheckor
End Function

Public Sub iisCopy(Optional ByVal AttrList As String)
26510     On Error GoTo errHandler
26520     If iisPosted Then
26530         iisCopyForm AttrList
26540         Else
26550         iisCopyQuery AttrList
26560     End If
26570     Exit Sub
errHandler:
26580     ErrorIn "ujXML.iisCopy(AttrList)", AttrList
End Sub
   
Public Function iisCopyFormSafe(Optional ByVal AttrList As String) As Boolean
26590     On Error Resume Next
26600     iisCopyForm AttrList
26610     iisCopyFormSafe = iisCheckor
End Function
   
Public Sub iisCopyForm(Optional ByVal AttrList As String)
26620     On Error GoTo errHandler
          Dim arrA, arrF, i As Long, Req As Object
26630     If AttrList = "" Then AttrList = iisFormFields
26640     uUnpackEQList AttrList, arrA, arrF
26650     Set Req = iisGetObject("Request")
26660     For i = 0 To UBound(arrA)
26670         attrS(arrA(i)) = Req.Form(arrF(i))
26680     Next i
26690     Exit Sub
errHandler:
26700     ErrorIn "ujXML.iisCopyForm(AttrList)", AttrList
End Sub
   
Public Function iisCopyQuerySafe(Optional AttrList As String) As Boolean
26710     On Error Resume Next
26720     iisCopyQuery AttrList
26730     iisCopyQuerySafe = iisCheckor
End Function
   
Public Sub iisCopyQuery(Optional ByVal AttrList As String)
26740     On Error GoTo errHandler
          Dim arrA, arrF, i As Long, Req As Object
26750     If AttrList = "" Then AttrList = iisQueryParams
26760     uUnpackEQList AttrList, arrA, arrF
26770     Set Req = iisGetObject("Request")
26780     For i = 0 To UBound(arrA)
26790         attrS(arrA(i)) = Req.QueryString.Item(arrF(i)) '.QueryString.Item(arrF(i))
26800     Next i
26810     Exit Sub
errHandler:
26820     ErrorIn "ujXML.iisCopyQuery(AttrList)", AttrList
End Sub
   
Public Function iisPosted() As Boolean
26830     On Error Resume Next
26840     iisPosted = (UCase$(iisRequestMethod) = "POST")
End Function

Public Function iisQueryParamsSafe() As String
26850     On Error Resume Next
26860     iisQueryParamsSafe = iisQueryParams
26870     iisCheckor
End Function
   
Public Function iisQueryParams() As String
26880     On Error GoTo errHandler
          Dim s As String, i As Long, Req As Object 'As ASPTypeLibrary.Request
26890     Set Req = iisGetObject("Request")
26900     For i = 1 To Req.QueryString.Count
26910         s = uCombine(s, ",", Req.QueryString.Key(i))
26920     Next i
26930     iisQueryParams = s
26940     Exit Function
errHandler:
26950     ErrorIn "ujXML.iisQueryParams"
End Function
   
Public Function iisFormFieldsSafe() As String
26960     On Error Resume Next
26970     iisFormFieldsSafe = iisFormFields
26980     iisCheckor
End Function
   
Public Function iisFormFields() As String
26990     On Error GoTo errHandler
          Dim s As String, i As Long, Req As Object
27000     Set Req = iisGetObject("Request")
27010     For i = 1 To Req.Form.Count
27020         s = uCombine(s, ",", Req.Form.Key(i))
27030     Next i
27040     iisFormFields = s
27050     Exit Function
errHandler:
27060     ErrorIn "ujXML.iisFormFields"
End Function
   
Public Function iisRequestMethodSafe() As String
27070     On Error Resume Next
27080     iisRequestMethodSafe = iisRequestMethod
27090     iisCheckor
End Function
   
Public Function iisRequestMethod() As String
27100     On Error GoTo errHandler
          Dim objReq As Object 'ASPTypeLibrary.Request
27110     Set objReq = iisGetObject("Request")
27120     Check Not objReq Is Nothing, X_UJXML, "Failed to retrieve IIS Request object."
27130     iisRequestMethod = objReq.ServerVariables("REQUEST_METHOD")
27140     Exit Function
errHandler:
27150     ErrorIn "ujXML.iisRequestMethod"
End Function

Public Function iisWriteSafe(ByVal HtmlText As String) As Boolean
27160     On Error Resume Next
27170     iisWrite HtmlText
27180     iisWriteSafe = iisCheckor
End Function

Public Sub iisWrite(ByVal HtmlText As String)
27190     On Error GoTo errHandler
          Dim objResp As Object 'ASPTypeLibrary.Request
27200     Set objResp = iisGetObject("Response")
27210     Check Not objResp Is Nothing, X_UJXML, "Failed to retrieve IIS Response object."
27220     objResp.Write HtmlText
27230     Exit Sub
errHandler:
27240     ErrorIn "ujXML.iisWrite(HtmlText)", "..."
End Sub

Public Function iisReportWriteSafe() As Boolean
27250     On Error Resume Next
27260     iisReportWrite
27270     iisReportWriteSafe = iisCheckor
End Function

Public Sub iisReportWrite()
27280     On Error GoTo errHandler
27290     iisWrite iisReportHTML
27300     Exit Sub
errHandler:
27310     ErrorIn "ujXML.iisReportWrite"
End Sub

Public Function iisGetObjectSafe(ByVal ObjectName As String) As Object
27320     On Error Resume Next
27330     Set iisGetObjectSafe = iisGetObject(ObjectName)
27340     iisCheckor
End Function
   
Public Function iisGetObject(ByVal ObjectName As String) As Object
27350     On Error GoTo errHandler
          Dim Ctx As ObjectContext
27360     Set Ctx = GetContext
27370     Set iisGetObject = Ctx(ObjectName)
27380     Exit Function
errHandler:
27390     ErrorIn "ujXML.iisGetObject"
End Function

Private Function GetContext() As ObjectContext
27400     On Error Resume Next
27410     Set GetContext = GetObjectContext
End Function

Private Sub CheckSetSysHandler()
27420     On Error GoTo errHandler
27430     If (Not docOptions(XO_NO_SYS_HANDLER)) And (Not errSysHandlerWasSet) Then
27440         errSysHandlerSet
27450     End If
27460     Exit Sub
errHandler:
27470     ErrorIn "ujXML.CheckSetSysHandler"
End Sub

'Version group ==============================================================
Public Property Get verMajor() As Long
27480     verMajor = App.Major
End Property

Public Property Get verMinor() As Long
27490     verMinor = App.Minor
End Property

Public Property Get verRevision() As Long
27500     verRevision = App.Revision
End Property

Public Function verVersion() As String
27510     verVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

'Utilities ================================================================
Public Function uEscapeXML(ByVal s As String) As String
27520     s = Replace(s, "&", "&amp;")
27530     s = Replace(s, "<", "&lt;")
27540     s = Replace(s, ">", "&gt;")
27550     s = Replace(s, """", "&quot;")
27560     s = Replace(s, "'", "&apos;")
27570     uEscapeXML = s
End Function

Public Function uEscapeHTML(ByVal s As String) As String
27580     s = Replace(s, "&", "&amp;")
27590     s = Replace(s, "<", "&lt;")
27600     s = Replace(s, ">", "&gt;")
27610     s = Replace(s, """", "&quot;")
27620     uEscapeHTML = s
End Function


'Utilities, Private
Private Function uListGetValue(ByVal ctrl As Object)
27630     On Error GoTo errHandler
          Dim i As Long
27640     With ctrl
27650         If uListHasData(ctrl) Then
27660             If .ListIndex = -1 Then uListGetValue = 0 Else uListGetValue = .ItemData(.ListIndex)
27670             Else
27680             uListGetValue = .ListIndex
27690         End If
27700     End With
27710     Exit Function
errHandler:
27720     ErrorIn "ujXML.uListGetValue(Ctrl)", ctrl
End Function

Private Sub uListSetValue(ByVal ctrl As Object, ByVal Value As Long)
27730     On Error GoTo errHandler
          Dim i As Long
27740     With ctrl
27750         If uListHasData(ctrl) Then
27760             For i = 0 To .ListCount - 1
27770                 If .ItemData(i) = Value Then .ListIndex = i: Exit Sub
27780             Next i
27790             .ListIndex = -1
27800             Else
27810             .ListIndex = Value
27820         End If
27830     End With
27840     Exit Sub
errHandler:
27850     ErrorIn "ujXML.uListSetValue(Ctrl,Value)", Array(ctrl, Value)
End Sub

Private Sub uListSafeSelect(ByVal ctrl As Object, ByVal Ind As Long, ByVal Sel As Boolean)
27860     On Error Resume Next
27870     ctrl.Selected(Ind) = Sel
27880     If Err.Number <> 0 And Sel Then ctrl.ListIndex = Ind
27890     Err.Clear
End Sub

Private Function uListHasData(ByVal ctrl As Object) As Boolean
27900     On Error GoTo errHandler
          Dim i As Long
27910     uListHasData = True
27920     With ctrl
27930         For i = 0 To .ListCount - 1
27940             If .ItemData(i) <> 0 Then Exit Function
27950         Next i
27960     End With
27970     uListHasData = False
27980     Exit Function
errHandler:
27990     uListHasData = False
          'orIn "ujXML.uListHasData(Ctrl)", Ctrl
End Function

Private Function uListGetMultiSelect(ByVal ctrl As Object) As String
28000     On Error GoTo errHandler
          Dim i As Long, s As String, blnHasData As Boolean
28010     blnHasData = uListHasData(ctrl)
28020     With ctrl
28030         For i = 0 To .ListCount - 1
28040             If .Selected(i) Then
28050                 If s <> "" Then s = s & ","
28060                 If blnHasData Then s = s & .ItemData(i) Else s = s & i
28070             End If
28080         Next i
28090     End With
28100     uListGetMultiSelect = s
28110     Exit Function
errHandler:
28120     ErrorIn "ujXML.uListGetMultiSelect(Ctrl)", ctrl
End Function

Private Sub uListSetMultiSelect(ByVal ctrl As Object, ByVal DataList As String)
28130     On Error GoTo errHandler
          Dim i As Long, V As Long, blnHasData As Boolean
28140     blnHasData = uListHasData(ctrl)
28150     With ctrl
28160         For i = 0 To .ListCount - 1
28170             If blnHasData Then V = .ItemData(i) Else V = i
28180             .Selected(i) = uInList(DataList, V)
28190         Next i
28200     End With
28210     Exit Sub
errHandler:
28220     ErrorIn "ujXML.uListSetMultiSelect(Ctrl,DataList)", Array(ctrl, DataList)
End Sub

Private Function uListAdd(ByVal sList As String, ByVal sValue As String, Optional ByVal Delim As String = ",") As String
28230     On Error GoTo errHandler
28240     If Not uInList(sList, sValue, Delim) Then sList = uCombine(sList, Delim, sValue)
28250     uListAdd = sList
28260     Exit Function
errHandler:
28270     ErrorIn "modXUtils.uListAdd(sList,sValue)", Array(sList, sValue)
End Function

Private Function uListRemove(ByVal sList As String, ByVal sValue As String) As String
28280     On Error GoTo errHandler
          Dim SL As String, sV As String, p As Long
28290     SL = "," & sList & ","
28300     sV = "," & sValue & ","
28310     SL = Replace(SL, sV, ",")
28320     uListRemove = Mid$(SL, 2, Len(SL) - 2) 'cut-off commas
28330     Exit Function
errHandler:
28340     ErrorIn "modXUtils.uListRemove(sList,sValue)", Array(sList, sValue)
End Function

Private Function uListRemoveLast(ByVal sList As String) As String
28350     On Error GoTo errHandler
          Dim p As Long
28360     p = InStrRev(sList, ",")
28370     If p = 0 Then sList = "" Else sList = Left$(sList, p - 1)
28380     uListRemoveLast = sList
28390     Exit Function
errHandler:
28400     ErrorIn "modXUtils.uListRemoveLast(sList)", sList
End Function

Private Function uInList(ByVal sList As String, ByVal sValue As String, Optional ByVal Delim As String = ",") As Boolean
          Dim SL As String, sV As String
28410     SL = Delim & sList & Delim
28420     sV = Delim & sValue & Delim
28430     uInList = (InStr(1, SL, sV) > 0)
End Function

Private Function uInsInto(ByVal IntoStr As String, ByVal Values, _
        Optional XMLEscape As Boolean = False) As String
28440     On Error GoTo errHandler
          Dim i As Long, ChrX As String, strV As String
28450     ChrX = Chr$(vbKeyBack) 'Spec char to act instead of % during manipulations
28460     IntoStr = Replace(IntoStr, "%", ChrX)
28470     If Not IsArray(Values) Then Values = Array(Values)
28480     For i = LBound(Values) To UBound(Values)
28490         strV = uSafeStr(Values(i))
28500         If XMLEscape Then strV = uEscapeXML(strV)  'Replace(strV, "'", "''")  '
28510         IntoStr = Replace(IntoStr, ChrX & (i + 1), strV)
28520     Next i
28530     uInsInto = Replace(IntoStr, ChrX, "%") 'replace back
28540     Exit Function
errHandler:
28550     ErrorIn "ujXML.uInsInto(IntoStr,Values)", Array(IntoStr, Values)
End Function


Private Function uDBValue(Value)
28560   uDBValue = IIf(IsEmpty(Value), Null, Value)
End Function

Private Function uVBValue(Value)
          Dim V
28570     uVBValue = IIf(IsNull(Value), V, Value)
End Function

Private Function uGetField(ByVal AdoRs As Object, ByVal FldName As String)
28580     On Error GoTo errHandler
          Dim f As Object
28590     Set f = AdoRs.Fields(FldName)
28600     If Not IsNull(f.Value) Then uGetField = f.Value
28610     Exit Function
errHandler:
28620     ErrorIn "ujXML.uGetField(AdoRS,FldName)", Array(AdoRs, FldName)
End Function

Private Function uSafeGetFieldObj(ByVal AdoRs As Object, ByVal FldName As String) As Object
28630     On Error Resume Next
28640     Set uSafeGetFieldObj = AdoRs.Fields(FldName)
End Function


Private Function uIsElement(ByVal Node As MSXML2.IXMLDOMNode) As Boolean
28650     On Error GoTo errHandler
28660     uIsElement = (Node.NodeType = NODE_ELEMENT)
28670     Exit Function
errHandler:
28680     ErrorIn "ujXML.uIsElement(Node)", Node
End Function

Private Function uIsAttribute(ByVal Node As MSXML2.IXMLDOMNode) As Boolean
28690     On Error GoTo errHandler
28700     uIsAttribute = (Node.NodeType = NODE_ATTRIBUTE)
28710     Exit Function
errHandler:
28720     ErrorIn "ujXML.uIsAttribute(Node)", Node
End Function

Private Function uCompare(ByVal v1, ByVal v2) As Long
28730     On Error GoTo errHandler
28740     v1 = uToEmpty(v1)
28750     v2 = uToEmpty(v2)
28760     If IsEmpty(v1) Then
28770         uCompare = IIf(IsEmpty(v2), 0, -1)
28780     ElseIf IsEmpty(v2) Then
28790         uCompare = IIf(IsEmpty(v1), 0, 1)
28800     Else
28810         If IsNumeric(v1) And IsNumeric(v2) Then
28820             v1 = uSafeDbl(v1)
28830             v2 = uSafeDbl(v2)
28840         ElseIf IsDate(v1) And IsDate(v2) Then
28850             v1 = uSafeDate(v1)
28860             v2 = uSafeDate(v2)
28870         End If
28880         If v1 = v2 Then
28890             uCompare = 0
28900         ElseIf v1 < v2 Then
28910             uCompare = -1
28920         Else
28930             uCompare = 1
28940         End If
28950     End If
28960     Exit Function
errHandler:
28970     ErrorIn "ujXML.uCompare(V1,V2)", Array(v1, v2)
End Function

Private Property Get uEmpty() As Variant
          Dim V
28980     uEmpty = V
End Property

Private Function uToEmpty(ByVal Value)
28990     uToEmpty = Value
29000     If IsNull(Value) Or IsEmpty(Value) Then
29010         uToEmpty = uEmpty
29020     ElseIf VarType(Value) = vbString Then
29030         If Value = "" Then uToEmpty = uEmpty
29040     End If
End Function

Private Function uSafeListSorted(ByVal c As Object) As Boolean
29050     On Error Resume Next
29060     uSafeListSorted = c.Sorted
End Function

Private Function uSafeListMultiSelect(ByVal c As Object) As Boolean
29070     On Error Resume Next
29080     uSafeListMultiSelect = ((c.Style <> 0) Or (c.MultiSelect <> 0))
29090     If Err.Number <> 0 Then
29100         uSafeListMultiSelect = ((c.ListStyle <> 0) Or (c.MultiSelect <> 0))
29110     End If
End Function

Private Function uSafeCtrlStyle(ByVal c As Object) As Long
29120     On Error Resume Next
29130     uSafeCtrlStyle = c.Style
End Function

Private Property Get uChrBk() As String
29140     uChrBk = Chr$(vbKeyBack)
End Property

Private Function uUnpackEQList(ByVal ListAeqB As String, ByRef arrA, ByRef arrB, Optional ByRef arrQ) As Long
29150     On Error GoTo errHandler
          Dim arrAB() As String, arrT() As String, i As Long, SA As String, SB As String, Ubnd As Long
          Dim arr
29160     arrAB = Split(ListAeqB, ",")
29170     Ubnd = UBound(arrAB)
29180     uUnpackEQList = Ubnd
29190     arrA = uCreateArr(Ubnd)
29200     arrB = uCreateArr(Ubnd)
29210     If Not IsMissing(arrQ) Then arrQ = uCreateArr(Ubnd)
29220     For i = 0 To Ubnd
29230         If Not IsMissing(arrQ) Then
29240             arrQ(i) = (InStr(arrAB(i), "'") > 0)
29250             If arrQ(i) Then arrAB(i) = Replace(arrAB(i), "'", "")
29260         End If
29270         arrT = Split(arrAB(i), "=")
29280         arrA(i) = arrT(0)
29290         If UBound(arrT) > 0 Then arrB(i) = arrT(1) Else arrB(i) = arrT(0)
29300     Next i
29310     Exit Function
errHandler:
29320     ErrorIn "modXUtils.uUnpackEQList(ListAeqB,arrA,arrB,arrQ)", Array(ListAeqB, arrA, arrB, arrQ)
End Function


Private Function uCreateArr(ByVal ArrUbnd As Long)
          Dim arr()
29330     ReDim arr(0 To ArrUbnd)
29340     uCreateArr = arr
End Function

Private Function uSafeFormat(ByVal V, ByVal Fmt As String) As String
29350     On Error Resume Next
29360     uSafeFormat = Format(V, Fmt)
End Function

Private Function uCreateEQList(ByRef arrN, ByRef arrV, Optional ByVal Delim = ", ") As String
29370     On Error GoTo errHandler
          Dim arrRes() As String, i As Long
29380     ReDim arrRes(0 To UBound(arrN))
29390     For i = 0 To UBound(arrN)
29400         arrRes(i) = arrN(i) & " = " & arrV(i)
29410     Next i
29420     uCreateEQList = Join(arrRes, Delim)
29430     Exit Function
errHandler:
29440     ErrorIn "ujXML.uCreateEQList(arrN,arrV,Delim)", Array(arrN, arrV, Delim)
End Function

'Recordset =====================================================================================
Private Function uSaveRecordset(ByVal rs As Object) As String
29450     On Error GoTo errHandler
          Dim s As Object
29460     Set s = CreateObject("ADODB.Stream")
29470     s.Open
29480     rs.Save s, 1 'adPersistXML
29490     s.Position = 0
29500     uSaveRecordset = s.ReadText
29510     Exit Function
errHandler:
29520     ErrorIn "modXUtils.uSaveRecordset(RS)", rs
End Function

'Extracts z:row elements, changes [z:row] to ZRowElem
Private Function uExtractRecords(ByVal rs As Object, ByVal ZRowElem As String, _
            Optional ByVal RsElem As String = "") As String
29530     On Error GoTo errHandler
          Dim p1 As Long, P2 As Long, xml As String
29540     xml = uSaveRecordset(rs)
29550     p1 = InStr(1, xml, "<rs:data>")
29560     P2 = InStr(1, xml, "</rs:data>")
29570     If P2 > p1 And p1 > 0 Then
29580         p1 = p1 + Len("<rs:data>")
29590         xml = Mid$(xml, p1, P2 - p1)
29600         xml = uChangeElemName(xml, "z:row", ZRowElem)
29610         xml = Replace(xml, "<rs:insert>", "")
29620         xml = Replace(xml, "</rs:insert>", "")
29630         Else
29640         xml = ""
29650     End If
29660     If RsElem <> "" Then xml = uEnclose(xml, RsElem)
29670     uExtractRecords = xml
29680     Exit Function
errHandler:
29690     Err.Raise Err.Number, Err.Source, "(uExtractRecords) " & Err.Description
End Function

'We can use string matching, because patterns include special symbols that
'are always escaped in XML attribute values
Private Function uChangeElemName(ByVal xml As String, _
                                ByVal OldName As String, _
                                ByVal NewName As String) As String
29700     xml = Replace(xml, "<" & OldName & " ", "<" & NewName & " ")
29710     xml = Replace(xml, "<" & OldName & ">", "<" & NewName & ">")
29720     xml = Replace(xml, "<" & OldName & "/>", "<" & NewName & "/>")
29730     xml = Replace(xml, "</" & OldName & ">", "</" & NewName & ">")
29740     uChangeElemName = xml
End Function

Private Function uMakeElement(ByVal Elem As String, ByVal attrS As String, ByVal Content As String) As String
29750     uMakeElement = "<" & Elem & attrS & IIf(Content = "", "/>", ">" & Content & "</" & Elem & ">") & vbNewLine
End Function

Private Function uMakeAttr(ByVal Attr As String, ByVal Value As String) As String
29760     uMakeAttr = " " & Attr & "='" & uEscapeXML(Value) & "'"
End Function

Private Function uEnclose(ByVal xmlSrc As String, ByVal ElemOrTag As String, _
        Optional ByVal ClosingTag As String) As String
          Dim xmlRes As String
29770     If ElemOrTag = "" Then
29780         xmlRes = xmlSrc
29790         Else
29800         If InStr(1, ElemOrTag, "<") = 0 Then
29810             xmlRes = "<" & ElemOrTag & ">" & xmlSrc
29820             Else
29830             xmlRes = ElemOrTag & xmlSrc
29840         End If
29850         If ClosingTag = "" Then
29860             xmlRes = xmlRes & "</" & ElemOrTag & ">"
29870             Else
29880             xmlRes = xmlRes & ClosingTag
29890         End If
29900     End If
29910     uEnclose = xmlRes
End Function

'=========================================================================
Private Function uSafeStr(ByVal V, Optional ByVal Default) As String
29920   On Error Resume Next
29930   uSafeStr = IIf(IsMissing(Default), "", Default)
29940   uSafeStr = Trim$(V)
End Function

Private Function uSafeLng(ByVal V, Optional ByVal Default = 0) As Long
29950   On Error Resume Next
29960   uSafeLng = IIf(IsMissing(Default), 0, CLng(Default))
29970   If IsObject(V) Then uSafeLng = V.Value Else uSafeLng = CLng(V)
  End Function

Private Function uSafeDbl(ByVal V, Optional ByVal Default = 0#) As Double
29980   On Error Resume Next
29990   uSafeDbl = IIf(IsMissing(Default), 0#, Default)
30000   uSafeDbl = CDbl(V)
  End Function

Private Function uSafeDate(ByVal V, Optional ByVal Default) As Date
30010     On Error Resume Next
30020     uSafeDate = IIf(IsMissing(Default), CDate(0), Default)
30030     uSafeDate = CDate(V)
30040     If Err.Number <> 0 And VarType(V) = vbString Then
30050         If V Like "*T??:??:??" Then uSafeDate = CDate(Replace(V, "T", " "))
30060     End If
  End Function

Private Function uSafeBool(ByVal V) As Boolean
30070     On Error Resume Next
30080     uSafeBool = CBool(V)
End Function

Function uCombine(ByVal St1 As String, Delim As String, ByVal St2 As String) As String
30090     If St2 = "" Then
30100         uCombine = St1
30110     ElseIf St1 = "" Then
30120         uCombine = St2
30130     Else
30140         uCombine = St1 & Delim & St2
30150     End If
End Function

Private Function uRemoveObject(ByVal Col As Collection, ByVal obj As Object) As Boolean
30160     On Error GoTo errHandler
        Dim colObj As Object, Index As Long
30170   uRemoveObject = True
30180   For Index = 1 To Col.Count
30190     If IsObject(Col(Index)) Then
30200       Set colObj = Col(Index)
30210       If colObj Is obj Then
30220         Col.Remove Index
30230         Exit Function
30240         End If
30250       End If ' IsObject...
30260     Next Index
30270   uRemoveObject = False
30280     Exit Function
errHandler:
30290     ErrorIn "modXUtils.uRemoveObject(col,Obj)", Array(Col, obj)
End Function
  
Private Function uValueIn(ByVal Value, ParamArray Values()) As Boolean
        Dim i As Integer
30300   uValueIn = True
30310   For i = LBound(Values) To UBound(Values)
30320     If Value = Values(i) Then Exit Function
30330     Next i
30340   uValueIn = False
  End Function

Private Function uQuote(ByVal s As String) As String
30350     uQuote = "'" & Replace(s, "'", "''") & "'"
End Function

Private Function uDQuote(ByVal s As String) As String
30360     uDQuote = """" & Replace(s, """", """""") & """"
End Function

Private Function uSmartQuote(ByVal s As String) As String
30370     If InStr(1, s, "'") = 0 Then
30380         uSmartQuote = uQuote(s)
30390         Else
30400         uSmartQuote = uDQuote(s)
30410     End If
End Function

Private Function uIndexOf(ByVal SubStr As String, ByVal strList As String, _
        Optional ByVal Delim As String = ",") As Long
30420     On Error GoTo errHandler
          Dim arrList() As String, i As Long
30430     uIndexOf = -1
30440     arrList = Split(strList, Delim)
30450     For i = 0 To UBound(arrList)
30460         uIndexOf = i + 1
30470         If LCase$(arrList(i)) = LCase(SubStr) Then
30480             uIndexOf = i
30490             Exit Function
30500         End If
30510     Next i
30520     Exit Function
errHandler:
30530     ErrorIn "ujXML.uIndexOf(SubStr,strList,Delim)", Array(SubStr, strList, Delim)
End Function

Private Function uSubstrByIndex(ByVal strList As String, ByVal Index As Long, _
            Optional ByVal Delim As String = ",") As String
30540     On Error GoTo errHandler
          Dim arrList() As String
30550     arrList = Split(strList, Delim)
30560     Check uInRange(Index, LBound(arrList), UBound(arrList)), X_UJXML, "Substring index (%1) out of range.", Index
30570     uSubstrByIndex = arrList(Index)
30580     Exit Function
errHandler:
30590     ErrorIn "ujXML.uSubstrByIndex(strList,Index,Delim)", Array(strList, Index, Delim)
End Function

Private Function uInRange(ByVal V, ByVal v1, ByVal v2) As Boolean
30600   On Error Resume Next
30610   uInRange = (V >= v1) And (V <= v2)
End Function

Private Function uIsChildOf(ByVal ChEl As IXMLDOMElement, ByVal ParEl As IXMLDOMElement) As Boolean
30620     On Error GoTo errHandler
          Dim ChNd As IXMLDOMNode, ParNd As IXMLDOMNode
30630     Set ChNd = ChEl
30640     Set ParNd = ParEl
30650     Do
30660         If ChNd Is ParNd Then
30670             uIsChildOf = True
30680             Exit Function
30690         End If
30700         Set ChNd = ChNd.ParentNode
30710     Loop Until ChNd Is Nothing
30720     Exit Function
errHandler:
30730     ErrorIn "ujXML.uIsChildOf(ChEl,ParEl)", Array(ChEl, ParEl)
End Function

'Fix List
'10/05/03  Method elGetAttrList: fixed search criteria


Private Sub Class_Initialize()

End Sub




