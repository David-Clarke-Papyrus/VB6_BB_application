VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_CustomerPOS_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(Optional CustomerID As Long, Optional pACNO As String, Optional pDefaultPhone As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim cmd As New ADODB.Command
Dim strSQL As String
Dim strWHERE As String
Dim udtProps As CustomerPOSProps
Dim udtData As CustomerPOSData
Dim objPB As PropertyBag
  
    If CustomerID > 0 Then
        strWHERE = " Customer_ID = " & CustomerID
    ElseIf Len(pACNO) > 0 Then
        strWHERE = " TP_ACNO = '" & pACNO & "'"
    ElseIf Len(pDefaultPhone) > 0 Then
        strWHERE = " TP_Phone = " & pDefaultPhone
    End If
    
    strSQL = "SELECT DISTINCT tCustomer.* FROM tCustomer WHERE " & strWHERE
'-------------------------------
    oPC.OpenLocalDatabase
'-------------------------------

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, oPC.DBLocalConn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        With rs
          udtProps.ID = !Customer_ID
          udtProps.CustomerTypeID = FNS(!C_Type)
          udtProps.Initials = FNS(!C_Initials)
          udtProps.title = FNS(!C_Title)
          udtProps.Address = FNS(!C_Address)
          udtProps.Name = FNS(!C_Name)
          udtProps.Phone = FNS(!C_Phone)
          udtProps.VATable = FNB(!C_VATABLE)
          udtProps.AcNo = FNS(!C_Acno)
          udtProps.DefaultDiscount = FNDBL(!C_DefaultDiscount)
          udtProps.Balance = FNN(!C_BALANCE)
          udtProps.CreditLimit = FNN(!C_CREDITLIMIT)
          udtProps.Balances = FNS(!C_BALANCES)
          .Close
        End With
        Set rs = Nothing
        LSet udtData = udtProps
        
        Set objPB = New PropertyBag
        With objPB
          .WriteProperty "State", udtData.Buffer
          Fetch = .Contents
        End With
        Set objPB = Nothing
  Else
    ' force an error
        Fetch = ""
  End If
    
    oPC.CloseLocalDatabase
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CustomerPOS_P.Fetch(CustomerID,pACNO,pDefaultPhone)", Array(CustomerID, pACNO, _
         pDefaultPhone)
End Function

Public Function Save(ByVal Buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim rsRole As ADODB.Recordset
Dim udtProps As CustomerPOSProps
Dim udtData As CustomerPOSData
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim tmpAddressBuffer As String
Dim strSema As String

Dim cmd As New ADODB.Command
Dim strSQL As String
    
    Set objPB = New PropertyBag
    arBuffer = Buffer
    With objPB
      .Contents = arBuffer
      udtData.Buffer = .ReadProperty("State")
    End With
    LSet udtProps = udtData
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    With rs
        .Open "SELECT * FROM tTP WHERE TP_ID = " & udtProps.ID, oPC.DBLocalConn, adOpenDynamic, adLockOptimistic
        If udtProps.IsNew Then rs.AddNew
        !C_Initials = FNS(udtProps.Initials)
        !C_Title = FNS(udtProps.title)
        !C_Name = FNS(udtProps.Name)
        !C_Note = FNS(udtProps.Note)
        !C_VATABLE = udtProps.VATable
        !C_DefaultDiscount = udtProps.DefaultDiscount
        !C_Acno = FNS(udtProps.AcNo)
        !C_Phone = FNS(udtProps.Phone)
        !C_SearchPhone = FNS(udtProps.SearchPhone)
        !C_Type = FNN(udtProps.CustomerTypeID)
        .Update
        If udtProps.IsNew Then
            rs.Bookmark = rs.Bookmark
          .MoveLast
          udtProps.ID = FNN(.Fields("Customer_ID"))
        End If
        
    Set objPBOut = New PropertyBag
  
    LSet udtData = udtProps
  
    With objPBOut
        .WriteProperty "State", udtData.Buffer
    End With
    Set objPB = Nothing

    Save = objPBOut.Contents
    Set objPBOut = Nothing
    rs.Close
    Set rs = Nothing
        
   End With
   
EXITH:
    Exit Function
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CustomerPOS_P.Save(buffer)", Buffer
End Function

Public Sub DeleteObject(ByVal ID As Long)
    On Error GoTo errHandler

    oPC.DBLocalConn.Execute "DELETE FROM tCustomer WHERE Customer_ID= " & ID

  Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CustomerPOS_P.DeleteObject(ID)", ID
End Sub




