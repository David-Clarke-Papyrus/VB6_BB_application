VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Retrieval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetAdvances(pCNote As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
On Error GoTo Errh

    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM vAdvances WHERE GD_CNote = '" & pCNote & "'", oPC.CO, adOpenKeyset
    Set GetAdvances = rs
    Exit Function
Errh:
    oError.SetError Err, Error, Now(), "z_Retrieval:GetAdvances", "", ""
End Function


Public Function GetFinals(pCNote As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
On Error GoTo Errh

    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM vFinals WHERE PD_CNote = '" & pCNote & "'", oPC.CO, adOpenKeyset
    Exit Function
Errh:
    oError.SetError Err, Error, Now(), "z_Retrieval:GetFinals", "", ""
    Set GetFinals = rs
    
End Function

Public Function GetAddDeds(pCNote As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
On Error GoTo Errh

    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tAD WHERE AD_CNote = '" & pCNote & "'", oPC.CO, adOpenKeyset
    Set GetAddDeds = rs
    Exit Function
Errh:
    oError.SetError Err, Error, Now(), "z_Retrieval:GetAddDeds", "", ""
    
End Function


Public Function GetBoth(pCNote As String) As ADODB.Recordset
Dim rsA As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim rsC As New ADODB.Recordset
Dim oldKey As String
Dim finFruitSpec As String
Dim advFruitSpec As String
Dim tempFruitKey As String
Dim c As Integer
On Error GoTo Errh

    rsA.CursorLocation = adUseClient
    rsA.Open "SELECT * FROM vAdvances WHERE GD_CNote = '" & pCNote & "'", oPC.CO, adOpenKeyset
    If rsA.EOF Then
        advFruitSpec = "ZZZZZZZZZZZZZ"
    Else
        advFruitSpec = rsA!FruitSpec
    End If
    
    rsF.CursorLocation = adUseClient
    rsF.Open "SELECT * FROM vFinals WHERE PD_CNote = '" & pCNote & "'", oPC.CO, adOpenKeyset
    If rsF.EOF Then
        finFruitSpec = "ZZZZZZZZZZZZZ"
    Else
        finFruitSpec = rsF!FruitSpec
    End If
    Set rsC = New ADODB.Recordset
    rsC.Fields.Append "Colour", adVarChar, 30
    rsC.Fields.Append "CNote", adVarChar, 30
    rsC.Fields.Append "Variety", adVarChar, 30
    rsC.Fields.Append "Pack", adVarChar, 30
    rsC.Fields.Append "Grade", adVarChar, 30
    rsC.Fields.Append "Brand", adVarChar, 30
    rsC.Fields.Append "Cnt", adVarChar, 30
    rsC.Fields.Append "RailDate", adVarChar, 30
    rsC.Fields.Append "Stmnt", adVarChar, 30
    rsC.Fields.Append "StatementF", adVarChar, 30
    rsC.Fields.Append "SumOfQty_Adv", adInteger
    rsC.Fields.Append "StatementDateF", adVarChar, 20
    rsC.Fields.Append "StatementRef", adVarChar, 20
    rsC.Fields.Append "SumOfAmt_Adv", adDouble
    rsC.Fields.Append "Rate_Adv", adDouble
    rsC.Fields.Append "SumOfQty_Fin", adInteger
    rsC.Fields.Append "SumOfAmt_Fin", adDouble
    rsC.Fields.Append "Rate_Fin", adDouble
    rsC.Fields.Append "SumOfFinAdvRec", adDouble
    rsC.Fields.Append "SumOfDiffs", adDouble
    rsC.Fields.Append "AddDed", adDouble
    rsC.Open
    c = 1
    Do While Not (rsA.EOF And rsF.EOF)
        rsC.AddNew
        rsC!SumOfQty_Adv = 0
        rsC!SumOfAmt_Adv = 0
        rsC!SumOfQty_Fin = 0
        rsC!SumOfAmt_Fin = 0
        rsC!Rate_Adv = 0
        rsC!Rate_Fin = 0
        rsC!SumOfFinAdvRec = 0
        rsC!SumOfDiffs = 0
        rsC!AddDed = 0
        tempFruitKey = oldKey
        If advFruitSpec < finFruitSpec Then
            oldKey = advFruitSpec
        Else
            oldKey = finFruitSpec
        End If
        If tempFruitKey <> oldKey Then
            c = c * -1
        End If
        rsC!Colour = c
        If advFruitSpec <= oldKey And Not rsA.EOF Then
            rsC!CNote = rsA!GD_CNote
            rsC!Variety = rsA!GD_Variety
            rsC!Pack = rsA!GD_Pack
            rsC!Grade = rsA!GD_Grade
            rsC!Brand = rsA!GD_Brand
            rsC!Cnt = rsA!GD_Count
            rsC!RailDate = rsA!RailDate
            rsC!Stmnt = rsA!StatementDate
            rsC!SumOfQty_Adv = rsA!SumOfQty_Adv
            rsC!SumOfAmt_Adv = rsA!SumOfAmt_Adv
            If FNN(rsA.Fields("SumOfQty_Adv")) > 0 Then
              rsC!Rate_Adv = FNDBL(rsA!SumOfAmt_Adv) / FNN(rsA.Fields("SumOfQty_Adv"))
            End If
            rsA.MoveNext
            If rsA.EOF Then
                advFruitSpec = "ZZZZZZZZZZZZZ"
            Else
                advFruitSpec = rsA!FruitSpec
            End If
        End If
        If finFruitSpec <= oldKey And Not rsF.EOF Then
            rsC!CNote = rsF!PD_CNote
            rsC!Variety = rsF!PD_Variety
            rsC!Pack = rsF!PD_Pack
            rsC!Grade = rsF!PD_Grade
            rsC!Brand = rsF!PD_Brand
            rsC!Cnt = rsF!PD_Count
       '     rsC!RailDate = rsF!PD_Variety
            rsC!SumOfQty_Fin = FNN(rsF!SumOfQty_Fin)
            rsC!SumOfAmt_Fin = FNDBL(rsF!SumOfAmt_Fin)
            rsC!SumOfDiffs = FNDBL(rsF!SumOfDiffs)
            rsC!SumOfFinAdvRec = FNDBL(rsF!SumOfFinAdvRec)
            rsC!StatementDateF = FNS(rsF!StatementDateF)
            rsC!StatementRef = FNS(rsF!PD_StatementRef)
            rsC!AddDed = rsF!AddDed
            If FNN(rsF.Fields("SumOfQty_Fin")) > 0 Then
              rsC!Rate_Fin = FNDBL(rsF!SumOfAmt_Fin) / FNN(rsF.Fields("SumOfQty_Fin"))
            End If
            rsF.MoveNext
            If rsF.EOF Then
                finFruitSpec = "ZZZZZZZZZZZZZ"
            Else
                finFruitSpec = rsF!FruitSpec
            End If
        End If
        rsC.Update
    Loop
    
    Set GetBoth = rsC
    
    Exit Function
Errh:
    oError.SetError Err, Error, Now(), "z_Retrieval:GetBoth", "", ""
End Function

Public Function GetOS(pEntity As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
On Error GoTo Errh
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM vOS WHERE Entity = '" & pEntity & "' ORDER BY VARIETY,WEEK,Grade,Pack,Cnt ", oPC.CO, adOpenKeyset
    Set GetOS = rs
    Exit Function
Errh:
    oError.SetError Err, Error, Now(), "z_Retrieval:GetOS", "", ""
End Function

