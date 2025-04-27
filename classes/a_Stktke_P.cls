VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Stktke_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function Fetch(TransactionID As Long) As String
Dim rs As ADODB.Recordset
  
Dim strSQL As String
Dim udtProps As StockTakeProps
Dim udtData As StocktakeData
Dim objPB As PropertyBag
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    rs.Open "SELECT tTR.*,tSTKTKE.* FROM tTR INNER JOIN tSTKTKE On tSTKTKE.STKTKE_ID = tTR.TR_ID WHERE TR_ID = " & TransactionID, oPC.COShort, adOpenKeyset, adLockOptimistic
    
    On Error GoTo ERRH
    
    If Not rs.EOF Then
        With udtProps
            .TRID = rs.Fields("Tr_ID")
            .Code = FNS(rs.Fields("TR_Code"))
            .NominalDate = FND(rs.Fields("TR_Date"))
            .CaptureDate = FND(rs.Fields("TR_CaptureDate"))
            .Status = FNB(rs.Fields("TR_STatus"))
            .Note = FNS(rs.Fields("TR_Note"))
            .OperatorID = FNN(rs.Fields("TR_StAFFID"))
            .ValueOfStockRetail = FNN(rs.Fields("STKTKE_ValueOfStock"))
            .ValueOfStockCost = FNN(rs.Fields("STKTKE_ValueOfStockAtCost"))
            .ProductsWithoutPrices = FNN(rs.Fields("STKTKE_TitleswithoutPrices"))
            .ItemsWithoutPrices = FNN(rs.Fields("STKTKE_BookswithoutPrices"))
            .TotalItems = FNN(rs.Fields("STKTKE_TotalBooks"))
            .TotalProducts = FNN(rs.Fields("STKTKE_TotalTitles"))
            .AvgDiscount = FNDBL(rs.Fields("STKTKE_AvgDiscount"))
            .CutoffDate = FND(rs.Fields("STKTKE_CutoffDate"))

            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
            rs.Close
            Set rs = Nothing
            LSet udtData = udtProps
            Set objPB = New PropertyBag
            objPB.WriteProperty "State", udtData.buffer
            Fetch = objPB.Contents
            Set objPB = Nothing
            End With
    Else
        ' force an error
        rs.MoveNext
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
Exit Function
  
ERRH:
  rs.Close
  Set rs = Nothing
    Err.Raise Err.Number
  Resume
End Function
Public Function Save(ByVal buffer As String) As String
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim objCode As New z_Code
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim objPersist As a_Stktke_P
Dim udtProps As StockTakeProps
Dim udtData As StocktakeData
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.BeginTrans
    Set objPB = New PropertyBag
    arBuffer = buffer
    With objPB
      .Contents = arBuffer
      udtData.buffer = .ReadProperty("State")
    End With
    LSet udtProps = udtData
    Set rs = New ADODB.Recordset
   
    rs.Open "SELECT * FROM tTR WHERE TR_ID = " & udtProps.TRID, oPC.COShort, adOpenKeyset, adLockOptimistic
    On Error GoTo ERRH
        With udtProps
            If .IsNew Then
                rs.AddNew
                .Code = "ST" & objCode.GetNextCode(enStockTake, "GetandWrite")
            End If
            rs.Fields("TR_Type") = 7 'StockTake
            If IsDate(.NominalDate) Then rs.Fields("TR_Date") = .NominalDate Else rs.Fields("TR_Date") = Date
            rs.Fields("TR_CaptureDate") = .CaptureDate
            rs.Fields("TR_Code") = FNS(.Code)
            rs.Fields("TR_Status") = FNN(.Status)
            rs.Fields("TR_Note") = FNS(.Note)
            rs.Fields("TR_StaffID") = FNN(.OperatorID)
            rs.Update
            If .IsNew Then
                rs.MoveLast
                udtProps.TRID = rs.Fields("TR_ID")
            End If
            rs.Close
        End With

    strSQL = "SELECT * FROM tSTKTKE WHERE STKTKE_ID = " & udtProps.TRID
    rs.Open strSQL, oPC.COShort, adOpenDynamic, adLockOptimistic
    With udtProps
        If udtProps.IsNew Then rs.AddNew
        rs.Fields("STKTKE_ID") = udtProps.TRID
        rs.Fields("STKTKE_CutoffDate") = .CutoffDate
        rs.Fields("STKTKE_ValueOfStock") = .ValueOfStockRetail
        rs.Fields("STKTKE_ValueOfStockAtCost") = .ValueOfStockCost
        rs.Fields("STKTKE_TitleswithoutPrices") = .ProductsWithoutPrices
        rs.Fields("STKTKE_BookswithoutPrices") = .ItemsWithoutPrices
        rs.Fields("STKTKE_TotalTitles") = .TotalProducts
        rs.Fields("STKTKE_TotalBooks") = .TotalItems
        rs.Update
        rs.Close
    End With
    oPC.COShort.CommitTrans
    Set objPBOut = New PropertyBag
    LSet udtData = udtProps
    objPBOut.WriteProperty "State", udtData.buffer
    Set objPB = Nothing
    
    Save = objPBOut.Contents
    Set objPBOut = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
  
EXIT_HANDLER:
    Set rs = Nothing
    Exit Function
  
ERRH:
    oPC.COShort.RollbackTrans
    rs.Close
    Set rs = Nothing
    Err.Raise Err.Number
    GoTo EXIT_HANDLER
    Resume
End Function




Private Sub DeleteStockTake(SAID As Long, TRID As Long)
End Sub

Public Sub DeleteObject(TRID As Long)
    oPC.COShort.BeginTrans
    oPC.COShort.execute "DELETE FROM tSTKTKEL WHERE STKTKEL_TR_ID=" & TRID
    oPC.COShort.execute "DELETE FROM tSTKTKE WHERE STKTKE_ID=" & TRID
    oPC.COShort.execute "DELETE FROM tTR WHERE Tr_ID=" & TRID
    oPC.COShort.CommitTrans
End Sub




