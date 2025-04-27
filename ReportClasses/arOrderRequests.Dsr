VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arOrderRequests 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   16425
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   28972
   _ExtentY        =   13996
   SectionData     =   "arOrderRequests.dsx":0000
End
Attribute VB_Name = "arOrderRequests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long
Dim oC As c_Exchanges
Private xMLDoc As ujXML
Dim SortedDoc As ujXML
Dim Res As Boolean
Dim rs As ADODB.Recordset
Dim strAcnoKey As String

Sub component(p As c_Exchanges)
    Set oC = p
    i = 1
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 12600
    Me.Height = 6000
    
    Set rs = New ADODB.Recordset
    rs.Fields.Append "ExchangeID", adGUID
    rs.Fields.Append "CustomerID", adInteger
    rs.Fields.Append "CustomerName", adVarChar, 200
    rs.Fields.Append "CustomerAcno", adVarChar, 50
    rs.Fields.Append "CustomerTitle", adVarChar, 10
    rs.Fields.Append "CustomerInitials", adVarChar, 50
    rs.Fields.Append "CustomerAllDetails", adVarChar, 1000
    rs.Fields.Append "CustomerPhone", adVarChar, 50
    rs.Fields.Append "CustomerEmail", adVarChar, 200
    rs.Fields.Append "CustomerAddress", adVarChar, 1000

    rs.Fields.Append "CustomerContactDetails", adVarChar, 200
    rs.Fields.Append "OrderNotes", adVarChar, 1000
    rs.Fields.Append "TotalDeposit", adVarChar, 200
    rs.Fields.Append "ItemEAN", adVarChar, 20
    rs.Fields.Append "ItemDescription", adVarChar, 200
    rs.Fields.Append "ItemPrice", adVarChar, 20
    rs.Fields.Append "ItemDeposit", adVarChar, 20
    rs.Open , , adOpenDynamic, adLockOptimistic
    For i = 1 To oC.Count
        If Len(oC(i).Note) > 0 Then
            rs.AddNew
            rs.Fields("ExchangeID") = oC(i).ID
            LoadFromXML oC(i).Note
        End If
    Next
    rs.MoveFirst
    Set DC1.Recordset = rs
End Sub

'Private Sub Detail_Format()
'Dim ar() As String
'Dim ar2() As String
'Dim tmp As String
'
'    If i <= oC.Count Then
'        If Len(oC(i).Note) > 0 Then
'            LoadFromXML oC(i).Note
'        End If
'        tDeposit = oC(i).TotalPayableF
'        Detail.PrintSection
'        i = i + 1
'    End If
'End Sub

Private Sub LoadFromXML(pXML As String)
    On Error GoTo errHandler
Dim strAcno As String
Dim strTitle As String
Dim strInitials As String
Dim strCustname As String


        Set xMLDoc = New ujXML
       
        xMLDoc.docLoadXML pXML
        xMLDoc.navTop
        
        xMLDoc.navLocate "CustomerID"
        rs.Fields("CustomerID") = FNN(xMLDoc.Element.text)
        
        xMLDoc.navLocate "CustomerAcno"
        rs.Fields("CustomerAcno") = Left(xMLDoc.Element.text, 50)
        strAcnoKey = FNS(rs.Fields("CustomerAcno"))
        
        xMLDoc.navLocate "CustomerTitle"
        rs.Fields("CustomerTitle") = Left(xMLDoc.Element.text, 10)
        
        xMLDoc.navLocate "CustomerName"
        rs.Fields("CustomerName") = Left(xMLDoc.Element.text, 200)
        
        xMLDoc.navLocate "CustomerInitials"
        rs.Fields("CustomerInitials") = Left(xMLDoc.Element.text, 50)
       
        Res = xMLDoc.navLocate("CustomerPhone")
        rs.Fields("CustomerPhone") = Left(xMLDoc.Element.text, 50)
        
        xMLDoc.navLocate "CustomerEmail"
        rs.Fields("CustomerEmail") = Left(xMLDoc.Element.text, 200)
        
        xMLDoc.navLocate "CustomerAddress"
        rs.Fields("CustomerAddress") = Replace(xMLDoc.Element.text, Chr(10), vbCrLf)
        
        rs.Fields("CustomerAllDetails") = LTrim$(FNS(rs.Fields("CustomerTitle")) & " " & FNS(rs.Fields("CustomerInitials")) & " " & FNS(rs.Fields("CustomerName")) & " " & FNS(rs.Fields("CustomerPhone")) & " " & FNS(rs.Fields("CustomerEmail")) & " " & FNS(rs.Fields("CustomerAddress")))
        
        xMLDoc.navLocate "Notes"
        rs.Fields("OrderNotes") = Replace(xMLDoc.Element.text, Chr(10), vbCrLf)
        
        xMLDoc.navLocate "Deposit"
        rs.Fields("TotalDeposit") = Replace(xMLDoc.Element.text, Chr(10), vbCrLf)
        
        Res = xMLDoc.navLocate("ItemList")
        Set SortedDoc = xMLDoc.docCreateViewer(True)
        SortedDoc.navTop
        If SortedDoc.chCount > 0 Then
            SortedDoc.elForEachElem Me
        End If
        
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147216306 Then   'invalid XML
        rs.Fields("CustomerAllDetails") = "Invalid XML"
        Exit Sub
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arOrderRequests.LoadFromXML(pXML)", pXML
End Sub
Public Sub ProcessElement(ByVal xObj As ujXML, ByVal NavAction As XENUM_ITER_NAV, ByRef Param As Variant, ByRef SkipChildren As Boolean)
    On Error GoTo errHandler
Dim s As String
Dim sEAN As String
Dim sDep As String
Dim sDescr As String
Dim sPrice As String

    If IsMissing(Param) Then Param = ""
    If xObj.Element.nodeName = "Item" Then
        If NavAction <> XNAV_TO_PARENT Then
            If FNS(rs.Fields("ItemEAN")) > "" Then
                rs.AddNew
                rs.Fields("CustomerAcno") = strAcnoKey
            End If
            xObj.navFirstChild
            rs.Fields("ItemEAN") = xObj.Element.text
            Res = xObj.navNext
            rs.Fields("ItemPrice") = xObj.Element.text
            Res = xObj.navNext
            rs.Fields("ItemDescription") = xObj.Element.text
            Res = xObj.navNext
            rs.Fields("ItemDeposit") = xObj.Element.text
            xObj.navUP
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arOrderRequests.ProcessElement(xObj,NavAction,Param,SkipChildren)", Array(xObj, _
         NavAction, Param, SkipChildren), EA_NORERAISE
End Sub

