Attribute VB_Name = "POSGlobals"
Option Explicit



Global Const SI_TITLE = 1
Global Const SI_AUTHOR = 2
Global Const SI_UNITPR = 3
Global Const SI_QTY = 4
Global Const SI_DISC = 5
Global Const SI_PRICE = 6
Global Const SI_PID = 7
Public oGD As z_GetData

Public sTillCode As String

'Public Sub Main()
'    Dim oSale As New clsSaleLineCol
'
'End Sub

Public Sub LoadCombo(oRS As ADODB.Recordset, oCBO As ComboBox)
    oRS.MoveFirst
    oCBO.Clear
    
    With oRS
        Do While Not oRS.EOF
            oCBO.AddItem !SP_Code
            oCBO.ItemData(oCBO.NewIndex) = !SalesPerson_ID
            .MoveNext
        Loop
        .MoveFirst
    End With
    
End Sub

