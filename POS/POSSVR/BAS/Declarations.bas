Attribute VB_Name = "Declarations"
Option Explicit
Public oPC As PapyConn
Public strSQL As String
Public frm As frmPosServerMain
Public oMS As z_ManageStations
Public oMF As Z_ManageFolders
Public bDebug As Boolean
Public frmWS As frmWaitStatus

'Dim i As ListBox
Public Const CLIENTFILENAME = "\Client.dat"
Public Type ProductProps
    ID As Long
    Code As String * 10
    EAN As String * 13
    Availability As String * 5
    Description As String * 300
    CategoryID As Long
    BFClassification As String * 10
    UKPrice As Currency
    UKPoundsExch As Double
    USPrice As Currency
    USDollarExch As Double
    LastExchUpdate As Date
    LocalPrice As Currency
    Cost As Currency
    Title As String * 255
    BindingCode As String * 5
    SeriesTitle As String * 255
    SubTitle As String * 255
    Author As String * 255
    Publisher As String * 50
    Note As String * 40
    PublisherID As Long
    StockBalance As Long
    PublicationDate As String * 50
    PublicationPlace As String * 100
    MainSupplierName As String * 30
    Edition As String * 100
    LastSupplierID As Long
    LastDEalID As Long
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type


Public Sub TestLength()
    Dim T As ProductProps
    MsgBox LenB(T)
End Sub

Public Sub HandleError()
Dim frmErr As frmError
Dim strMsg As String
Dim strPos As String

    If InException Then
        strMsg = Err.Description
        Set frmErr = New frmError
        frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & oPC.SharedFolderRoot & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
        frmErr.Show vbModal
    Else
    '    If ErrInIDE Then frmShowError.ErrorReport = ErrReport
        ErrSaveToFile
        If oPC.ConnectionString > "" Then
            SaveToPC oPC.ConnectionString, 3, "spErrorLogInsert"
        End If
    End If
End Sub

Public Sub HandleError2()
    On Error Resume Next
Dim strMsg As String
Dim frmErr As frmError
Dim strPos As String

    ErrSaveToFile
    SaveToPC oPC.ConnectionString, 1, "spErrorLogInsert"
    If frmWS Is Nothing Then
    Else
        Unload frmWS
    End If
    If InException Then
        Select Case ErrNumber
            Case EXC_GENERAL
                strMsg = Err.Description
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & oPC.SharedFolderRoot & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            Case EXC_CANCELLED
                      'nothing to do - it is silent exception.
            Case EXC_MULTIPLE
                strMsg = Err.Description
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & oPC.SharedFolderRoot & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            Case EXC_VALIDATION
                strMsg = Err.Description
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & oPC.SharedFolderRoot & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            Case EXC_NOSERVER
                MsgBox "Server cannot be reached, closing application. ", vbOKOnly, "Exception"
            Case Else
                Set frmErr = New frmError
                strMsg = Description
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & oPC.SharedFolderRoot & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            
        End Select
    Else
        If ErrInIDE Then
            frmShowError.ErrorReport = ErrReport
        Else
            Screen.MousePointer = vbDefault
            If UCase(Left(ErrReport, 15)) = "TIMEOUT EXPIRED" Then
                MsgBox " A timeout error has occurred. Probably a record is being used by another user." & vbCrLf & "Try Again or cancel your action.", vbInformation, "Error in application"
            Else
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & oPC.SharedFolderRoot & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            
            End If
        End If
    End If
    
    Forms(0).ForceClose = True
    Unload Forms(0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "oMainc.HandleError"
End Sub


