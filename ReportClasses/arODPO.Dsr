VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arODPO 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   13996
   SectionData     =   "arODPO.dsx":0000
End
Attribute VB_Name = "arODPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ar As XArrayDB
Dim i As Long
Dim rs As ADODB.Recordset

Sub component(par As XArrayDB)
    Set ar = par
    i = 1
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 8000
    Me.Height = 8000

    Set rs = New ADODB.Recordset
    rs.Fields.Append "SupplierName", adVarChar, 200
    rs.Fields.Append "DocCode", adVarChar, 50
    rs.Fields.Append "EANF", adVarChar, 25
    rs.Fields.Append "Title", adVarChar, 200
    rs.Fields.Append "DocDate", adVarChar, 200
    rs.Fields.Append "Qty", adVarChar, 200
    rs.Fields.Append "QtyRec", adVarChar, 200
    rs.Fields.Append "QtyOS", adVarChar, 200
    rs.Fields.Append "ETA", adVarChar, 25
    rs.Fields.Append "ProductStatus", adVarChar, 200
    rs.Fields.Append "Action", adVarChar, 200
    rs.Open
    For i = par.LowerBound(1) To par.UpperBound(1)
        rs.AddNew
            rs.Fields("SupplierName") = Left(par.Value(i, 1), 200)
            rs.Fields("DocCode") = par.Value(i, 2)
            rs.Fields("EANF") = par.Value(i, 3)
            rs.Fields("Title") = par.Value(i, 4)
            rs.Fields("DocDate") = par.Value(i, 5)
            rs.Fields("Qty") = par.Value(i, 9)
            rs.Fields("QtyRec") = par.Value(i, 10)
            rs.Fields("QtyOS") = par.Value(i, 11)
            rs.Fields("ETA") = par.Value(i, 6)
            rs.Fields("ProductStatus") = par.Value(i, 8)
            rs.Fields("Action") = IIf(par.Value(i, 12) = 1, "T", "") & " " & IIf(par.Value(i, 13) = 1, "T", "") & " " & CStr(par.Value(i, 14))
        
        rs.Update
    Next i
    rs.MoveFirst
    Set DC1.Recordset = rs

End Sub

'Private Sub Detail_Format()
'    If i <= ar.Count(1) Then
'        tTitle = ar.Value(i, 3) & " " & ar.Value(i, 4)
'        tCode = ar.Value(i, 2)
'        tDate = ar.Value(i, 5)
'        tQty = ar.Value(i, 6)
'        tRecd = ar.Value(i, 7)
'        tOS = ar.Value(i, 8)
'        tAction = ar.Value(i, 9)
'        Detail.PrintSection
'        GroupHeader1.GroupValue = ar.Value(i, 1)
'        i = i + 1
'    End If
'End Sub
'
'Private Sub GroupHeader1_Format()
'    If i <= ar.UpperBound(1) And i > 0 Then
'        tSupplier = ar.Value(i, 1)
'        GroupHeader1.GroupValue = ar.Value(i, 1)
'    End If
'End Sub
'
'Private Sub ReportHeader_Format()
''    GroupHeader1.GroupValue = ar.Value(i, 1)
'End Sub
