Attribute VB_Name = "mSecurity"
Option Explicit
Public Const M_NEWPRODUCT As Integer = 1
Public Const M_NEWPRODUCTINADHOCFORM As Integer = 2
'Const M_DISCOUNT As Integer = 4
'Const M_CREDITNOTE As Integer = 8
'Const M_ISSUEAPPRO As Integer = 16
'Const M_REFUNDDEPOSIT As Integer = 32
'Const M_ACCEPTACPAYMENT As Integer = 64
'Const M_ISSUEPOSCREDITNOTE As Integer = 128
'Const M_ISSUEPOSREFUND As Integer = 256
'Const M_ACCEPTDIRECTDEPOSIT As Integer = 512
'Const M_PETTYCASH As Integer = 1024
'Const M_POSPRICECHANGE As Integer = 2048
'Const M_POSDISCOUNT As Integer = 4096
'Const M_CLOSEAPPLICATION As Integer = 8192
'Const M_DELETELINE As Integer = 16384

Public Function SecurityControlforSupervisor() As Boolean
    On Error GoTo errHandler
Dim frmS As New frmSecurity
Dim lngSMID As Long
    frmS.component "Supervisor signature"
    frmS.Caption = "Security: please enter your signature"
    frmS.Show vbModal
    If frmS.Cancelled Then
        SecurityControlforSupervisor = False
        Unload frmS
        Exit Function
    End If
    SecurityControlforSupervisor = True
    If oPC.Configuration.Staff.IsSupervisor(frmS.GetSignature, , lngSMID) = False Then
        MsgBox "You do not have supervisor status.", vbExclamation, "Action denied"
        SecurityControlforSupervisor = False
    Else
        gSTAFFID = lngSMID
    End If
    Unload frmS
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "mSecurity.SecurityControlforSupervisor"
End Function


Public Function SecurityControl(pNode As enumSecurityNode, Optional pCancelled As Boolean, _
        Optional pMsg As String, Optional pErrMsg As String, Optional IsSupervisor As Boolean, _
        Optional pName As String, Optional pSMID As Long, Optional pFullsignature As String) As Boolean
    On Error GoTo errHandler
Dim frmS As New frmSecurity
Dim strName As String
Dim lngSMID As Long
Dim strFullsignature As String

    frmS.component pMsg
    frmS.Caption = "Security: please enter your signature"
    frmS.Show vbModal
    If frmS.Cancelled Then
        SecurityControl = False
        Unload frmS
        Exit Function
    End If
    SecurityControl = True
    If Not oPC.Configuration.Staff.IsSecurityOK(pNode, frmS.GetSignature, strName, lngSMID, strFullsignature) Then
        If pErrMsg = "" Then pErrMsg = "You do not have security authority."
        MsgBox pErrMsg, vbExclamation, "Action denied"
        SecurityControl = False
    Else
        If Not IsMissing(pName) Then pName = strName
        If Not IsMissing(pSMID) Then pSMID = lngSMID
        If Not IsMissing(pFullsignature) Then pFullsignature = strFullsignature
        gSTAFFID = lngSMID
    End If
    Unload frmS
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "mSecurity.SecurityControl(pNode,pCancelled,pMsg,pErrMsg,IsSupervisor,pName,pSMID," & _
        "pFullsignature)", Array(pNode, pCancelled, pMsg, pErrMsg, IsSupervisor, pName, pSMID, pFullsignature)
End Function

Public Function CheckThisPoint(CheckPoint As Long) As Boolean
    On Error GoTo errHandler
    If (oPC.fSecurity And CheckPoint) = CheckPoint Then
        CheckThisPoint = True
    Else
        CheckThisPoint = False
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "mSecurity.CheckThisPoint(CheckPoint)", CheckPoint
End Function

