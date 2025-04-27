VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_PT_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function Fetch(Optional PTID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
'Dim cmd As New ADODB.Command
Dim strSQL As String
Dim strWHERE As String
Dim udtProps As PTProps
Dim udtData As PTData
Dim objPB As PropertyBag
Dim OpenResult As Integer

    strSQL = "SELECT * FROM tPT WHERE PT_ID = " & PTID '& " AND LEFT(PT_CODE,1) <> '*'"

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open strSQL, oPC.COShort, adOpenDynamic, adLockOptimistic
    If Not rs.eof Then
        With rs
          udtProps.ID = !PT_ID
          udtProps.Active = FNB(!PT_Active)
          udtProps.code = FNS(!PT_Code)
          udtProps.Number = FNS(!PT_Number)
          udtProps.Round = FNN(rs!PT_ROUNDINGMARGIN)
          udtProps.Discount = FNDBL(rs!PT_Discount)
          udtProps.B1Min = FNN(rs!PT_B1Min)
          udtProps.B1Max = FNN(rs!PT_B1Max)
          udtProps.B1MU = FNN(rs!PT_B1MU)
          udtProps.B2Min = FNN(rs!PT_B2Min)
          udtProps.B2Max = FNN(rs!PT_B2Max)
          udtProps.B2MU = FNN(rs!PT_B2MU)
          udtProps.B3Min = FNN(rs!PT_B3Min)
          udtProps.B3Max = FNN(rs!PT_B3Max)
          udtProps.B3MU = FNN(rs!PT_B3MU)
          udtProps.CRSALES = FNS(rs!PT_CRSALES)
          udtProps.CRSALES_CONTRA = FNS(rs!PT_CRSALES_CONTRA)
          udtProps.CASHSALES = FNS(rs!PT_CASHSALES)
          udtProps.CASHSALES_CONTRA = FNS(rs!PT_CASHSALES_CONTRA)
          udtProps.PURCHASES = FNS(rs!PT_PURCHASES)
          udtProps.PURCHASES_CONTRA = FNS(rs!PT_PURCHASES_CONTRA)
          udtProps.VAT = FNS(rs!PT_VAT)
          udtProps.SaleOrReturn = FNB(rs!PT_Seesafe)
          udtProps.SystemCode = FNS(rs!PT_SYSTEM)
          .Close
        End With
        Set rs = Nothing
        LSet udtData = udtProps
        
        Set objPB = New PropertyBag
        With objPB
          .WriteProperty "State", udtData.buffer
          Fetch = .Contents
        End With
        Set objPB = Nothing
  Else
    ' force an or
    Fetch = ""
  End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT_P.Fetch(PTID)", PTID
End Function

Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim udtProps As PTProps
Dim udtData As PTData
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim OpenResult As Integer
Dim strSQL As String

    Set objPB = New PropertyBag
    arBuffer = buffer
    With objPB
      .Contents = arBuffer
      udtData.buffer = .ReadProperty("State")
    End With
    LSet udtProps = udtData
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    With udtProps
        strSQL = "SELECT * FROM tPT WHERE PT_ID = " & .ID
        rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
        If .IsNew Then rs.AddNew
        rs!PT_Code = FNS(.code)
        rs!PT_Number = FNS(.Number)
        rs!PT_Active = FNB(.Active)
        rs!PT_ROUNDINGMARGIN = FNN(.Round)
        rs!PT_Discount = .Discount
        rs!PT_B1Min = FNN(.B1Min)
        rs!PT_B1Max = FNN(.B1Max)
        rs!PT_B1MU = FNN(.B1MU)
        rs!PT_B2Min = FNN(.B2Min)
        rs!PT_B2Max = FNN(.B2Max)
        rs!PT_B2MU = FNN(.B2MU)
        rs!PT_B3Min = FNN(.B3Min)
        rs!PT_B3Max = FNN(.B3Max)
        rs!PT_B3MU = FNN(.B3MU)
        rs!PT_CRSALES = FNS(udtProps.CRSALES)
        rs!PT_CRSALES_CONTRA = FNS(udtProps.CRSALES_CONTRA)
        rs!PT_CASHSALES = FNS(udtProps.CASHSALES)
        rs!PT_CASHSALES_CONTRA = FNS(udtProps.CASHSALES_CONTRA)
        rs!PT_PURCHASES = FNS(udtProps.PURCHASES)
        rs!PT_PURCHASES_CONTRA = FNS(udtProps.PURCHASES_CONTRA)
        rs!PT_VAT = FNS(udtProps.VAT)
          
        rs!PT_Seesafe = FNB(.SaleOrReturn)
        If FNB(.IsVoucher) = True Then
            rs!PT_SYSTEM = "V1"
        End If
        udtProps.dbactionStatus = 0
        rs.Update
        If .dbactionStatus = 0 Then
            If .IsNew Then
              rs.Bookmark = rs.Bookmark
              .ID = rs.fields("PT_ID")
            End If
            rs.Close
        End If
        Set rs = Nothing
    End With
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Set objPBOut = New PropertyBag
    
    LSet udtData = udtProps
    
    With objPBOut
        .WriteProperty "State", udtData.buffer
    End With
    Set objPB = Nothing
    
    Save = objPBOut.Contents
    Set objPBOut = Nothing
    
    Exit Function
        
EXITH:
  Exit Function
errHandler:
    ErrPreserve
    If Err = -2147217873 Then
        udtProps.dbactionStatus = 22
        Resume Next
    End If
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT_P.Save(buffer)", buffer
End Function

Public Sub DeleteObject(ByVal ID As Long)
    On Error GoTo errHandler
  Dim strSQL As String
    strSQL = "DELETE FROM tPT WHERE PT_ID=" & ID
    oPC.COShort.execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT_P.DeleteObject(ID)", ID
End Sub






