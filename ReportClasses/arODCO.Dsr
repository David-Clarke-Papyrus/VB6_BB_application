VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arODCO 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   18465
   MDIChild        =   -1  'True
   _ExtentX        =   32570
   _ExtentY        =   13996
   SectionData     =   "arODCO.dsx":0000
End
Attribute VB_Name = "arODCO"
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
    rs.Fields.Append "DocCode", adVarChar, 50
    rs.Fields.Append "EANF", adVarChar, 25
    rs.Fields.Append "Title", adVarChar, 200
    rs.Fields.Append "DocDate", adVarChar, 200
    rs.Fields.Append "ColRef", adVarChar, 25
    rs.Fields.Append "Qty", adVarChar, 200
    rs.Fields.Append "QtySent", adVarChar, 200
    rs.Fields.Append "QtyOS", adVarChar, 200
    rs.Fields.Append "ETA", adVarChar, 25
    rs.Fields.Append "ProductStatus", adVarChar, 200
    rs.Fields.Append "Action", adVarChar, 200
    rs.Fields.Append "CustName", adVarChar, 200
    rs.Open
    For i = par.LowerBound(1) To par.UpperBound(1)
        rs.AddNew
            rs.Fields("DocCode") = Left(par.Value(i, 2), 50)
            rs.Fields("EANF") = par.Value(i, 4)
            rs.Fields("Title") = Left(par.Value(i, 5), 200)
            rs.Fields("DocDate") = par.Value(i, 6)
            rs.Fields("Qty") = par.Value(i, 9)
            rs.Fields("QtySent") = par.Value(i, 10)
            rs.Fields("QtyOS") = par.Value(i, 11)
            rs.Fields("ETA") = par.Value(i, 7)
            rs.Fields("ProductStatus") = par.Value(i, 8)
            rs.Fields("COLREF") = Left(par.Value(i, 3), 25)
            rs.Fields("CustName") = Left(par.Value(i, 1), 200)
        
        rs.Update
    Next i
    rs.MoveFirst
    Set DC1.Recordset = rs
End Sub

