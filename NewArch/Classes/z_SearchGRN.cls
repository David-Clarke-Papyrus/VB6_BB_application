VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_SearchGRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim cDel As c_DELs
Dim dteDate1 As Date
Dim dteDate2 As Date
Dim enSince As enumSince

'Friend Sub Find()
'    On Error GoTo errHandler
'Dim bNotFound As Boolean
'Dim frm As frmBrowseSUppliers2
'Dim lngTPID As Long
'Dim byear As Boolean
'Dim yr As String
'Dim mth As String
'Dim strDate1 As String
'Dim strDate2 As String
'Dim lngCount As Long
'
'    bNotFound = False
'  '  txtArg = Replace(txtArg, " ", "")
'    If UCase(Left(txtArg, 3)) = "YR=" Then byear = True
'    If txtArg > " " And Not (byear) Then
'        If ArgIsProductCode Then
'            'Search for product code
'            enSince = 1
'            cbSince.Caption = TranslateSince(1)
'            Set cDel = Nothing
'            Set cDel = New c_DELs
'            cDel.Load bNotFound, 0, "", "", , , , txtArg
'            GoTo EXIT_Handler
'        End If
'        'Search for Reference
'        Set cDel = Nothing
'        Set cDel = New c_DELs
'        cDel.Load bNotFound, 0, "", txtArg ', dteDate1, dteDate2
'        If bNotFound Then
'            'Search for customer by ACCNO
'            Set cDel = Nothing
'            Set cDel = New c_DELs
'            cDel.Load bNotFound, 0, txtArg, "", dteDate1, dteDate2
'            If bNotFound Then
'                'Search for Invoice by supplier invoice no
'                Set cDel = Nothing
'                Set cDel = New c_DELs
'                SetDateArgs
'                cDel.Load bNotFound, 0, "", "", dteDate1, dteDate2, , , , txtArg
'                If bNotFound Then
'                    Set frm = New frmBrowseSUppliers2
'                    frm.component txtArg, lngCount
'                    If lngCount > 1 Then
'                        frm.Show vbModal
'                        lngTPID = frm.SupplierID
'                    ElseIf lngCount = 1 Then
'                        lngTPID = frm.SupplierID
'                    End If
'                    Unload frm
'                    If lngTPID > 0 Then
'                        Set cDel = Nothing
'                        Set cDel = New c_DELs
'                        SetDateArgs
'                        cDel.Load bNotFound, lngTPID, "", "", dteDate1, dteDate2
'                    End If
'                End If
'
'            Else
'                enSince = 1
'                cbSince.Caption = TranslateSince(1)
'            End If
'        End If
'    Else
'        If byear Then
'            yr = Mid(txtArg, 4, 4)
'            mth = Mid(txtArg, 9, 2)
'            If mth > "" Then
'                strDate1 = yr & "-" & mth & "-01"
'                strDate2 = yr & "-" & mth & "-" & LastDayOfMonth(yr & "-" & mth & "-01")
'            Else
'                strDate1 = yr & "-01-01"
'                strDate2 = yr & "-12-31"
'            End If
'            If Not (IsDate(strDate1) And IsDate(strDate2)) Then
'                SetDateArgs
'            Else
'                dteDate1 = CDate(strDate1)
'                dteDate2 = CDate(strDate2)
'            End If
'        Else
'            SetDateArgs
'        End If
'        cDel.Load bNotFound, 0, "", "", dteDate1, dteDate2
'    End If
'    Grid.Visible = True
'
'EXIT_Handler:
'    mSetfocus Grid
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseDels.Find"
'End Sub

Friend Sub RecentSearch()
Dim bNotFound As Boolean
    enSince = enQuarter
    SetDateArgs
    Set cDel = New c_DELs
    cDel.Load bNotFound, 0, "", "", dteDate1, dteDate2

End Sub

Private Sub SetDateArgs()
    Select Case enSince
    Case enAny
        dteDate1 = CDate("1995-01-01")
        dteDate2 = DateAdd("d", 1, Date)
    Case enWeek
        dteDate1 = DateAdd("d", -7, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case enMonth
        dteDate1 = DateAdd("m", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case enQuarter
        dteDate1 = DateAdd("q", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case enYear
        dteDate1 = DateAdd("yyyy", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    End Select

    Exit Sub
End Sub

Public Property Get BrowsedGRNs() As c_DELs
    Set BrowsedGRNs = cDel
End Property
