VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Reconciliation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim xMLDoc As ujXML
Dim dteReconDate As Date

Public Sub Component(pRs As ADODB.Recordset, pDate As Date)
    Set rs = pRs
    dteReconDate = pDate
End Sub

Public Function ExportToXML(DispatchMode As enTransmitType, pExVAT As Boolean, pStaffName As String) As Boolean
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim strFOFile As String
Dim strPDFFile As String
Dim strXML As String
Dim strCommand As String
Dim i As Integer
Dim strHTML As String
Dim fs As New FileSystemObject
Dim objXSL As New MSXML2.DOMDocument60
Dim opXMLDOC As New MSXML2.DOMDocument60
Dim objXMLDOC  As New MSXML2.DOMDocument60
Dim strWorkingFolder As String
Dim strDispatchMethod As String
Dim strExecutable As String
Dim cnt As Integer
Dim strPath As String

    If DispatchMode = enPrint Then
        If Not fs.FolderExists(oPC.SharedFolderRoot & "\Printing\") Then
            fs.CreateFolder oPC.SharedFolderRoot & "\Printing\"
        End If
        strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
        strDispatchMethod = "Printing"
    ElseIf DispatchMode = enView Then
        If Not fs.FolderExists(oPC.SharedFolderRoot & "\PDF\") Then
            fs.CreateFolder oPC.SharedFolderRoot & "\PDF\"
        End If
        strWorkingFolder = oPC.SharedFolderRoot & "\PDF\"
        strDispatchMethod = ""
    ElseIf DispatchMode = enXML Then
        If Not fs.FolderExists(oPC.SharedFolderRoot & "\XML\") Then
            fs.CreateFolder oPC.SharedFolderRoot & "\XML\"
        End If
        strWorkingFolder = oPC.SharedFolderRoot & "\XML\"
        strDispatchMethod = ""
    Else
        If Not fs.FolderExists(oPC.SharedFolderRoot & "\Printing\") Then
            fs.CreateFolder oPC.SharedFolderRoot & "\Printing\"
        End If
        strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
        strDispatchMethod = "Printing"
    End If
                        p 1

    Set xMLDoc = New ujXML
                        p 21
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "SOH_RECON_DOC"
            .chCreate "MessageType"
                .elText = "SOH_RECON_DOC"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
                        p 23
            .elCreateSibling "TemplateName"
                .elText = "SOH_RECON_DOC"
            .elCreateSibling "LogoPath"
                .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
            .elCreateSibling "DocPrintedDate", True
                .elText = Format(Now, "dd-mm-yyyy Hh:Nn")
            .elCreateSibling "DateOfReconcilation", True
                .elText = Format(dteReconDate, "dd-mm-yyyy")
            .elCreateSibling "PrintedBy", True
                .elText = pStaffName
            .elCreateSibling "VATMsg", True
                .elText = IIf(pExVAT = False, "Values include VAT", "Values exclude VAT")
            .elCreateSibling "CompanyName", True
                .elText = oPC.Configuration.DefaultCompany.CompanyName
            .elCreateSibling "StoreName", True
                .elText = oPC.Configuration.DefaultStore.Description
            If Not pExVAT Then
                .elCreateSibling "gBF"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("BF_STAT_OnHand_QtyItems")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("BF_STAT_ValueofStock_Cost")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("BF_STAT_ValueofStock_Retail")), "###,##0")
                    .elCreateSibling "PQTY", True
                        .elText = Format(FNN(rs.Fields("BF_STAT_OnHand_QtyProducts")), "###,##0")
                    .navUP
                .elCreateSibling "gREC"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_DEL_QtyItems_mm")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_DEL_Value_COST_mm")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_DEL_Value_Retail_mm")), "###,##0")
                    .navUP
                .elCreateSibling "gSAL"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_INV_QtyItems_mm")) + FNN(rs.Fields("STAT_CS_QtyItems_mm")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_INV_Value_Cost_mm")) + FNN(rs.Fields("STAT_CS_Value_Cost_mm")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_INV_Value_Retail_mm")) + FNN(rs.Fields("STAT_CS_Value_Retail_mm")), "###,##0")
                    .navUP
                .elCreateSibling "gRET"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_CN_QtyItems_mm")) + FNN(rs.Fields("STAT_CSR_QtyItems_mm")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_CN_Value_Cost_mm")) + FNN(rs.Fields("STAT_CSR_Value_Cost_mm")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_CN_Value_Retail_mm")) + FNN(rs.Fields("STAT_CSR_Value_Retail_mm")), "###,##0")
                    .navUP
                .elCreateSibling "gAPP"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_APP_QtyItems_mm")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_APP_Value_Cost_mm")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_APP_Value_Retail_mm")), "###,##0")
                    .navUP
                .elCreateSibling "gAPPR"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_APPR_QtyItems_mm")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_APPR_Value_Cost_mm")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_APPR_Value_Retail_mm")), "###,##0")
                    .navUP
                .elCreateSibling "gTFROUT"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_TFROUT_QtyItems_mm")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_TFROUT_Value_Cost_mm")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_TFROUT_Value_Retail_mm")), "###,##0")
                    .navUP
                .elCreateSibling "gTFRIN"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_TFRIN_QtyItems_mm")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_TFRIN_Value_Cost_mm")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_TFRIN_Value_Retail_mm")), "###,##0")
                    .navUP
                .elCreateSibling "gRETS"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_RTN_QtyItems_mm")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_RTN_Value_Cost_mm")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_RTN_Value_Retail_mm")), "###,##0")
                    .navUP
                .elCreateSibling "gCF"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_OnHand_QtyItems")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_ValueofStock_Cost")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_ValueofStock_Retail")), "###,##0")
                     .elCreateSibling "PQTY", True
                        .elText = Format(FNN(rs.Fields("STAT_OnHand_QtyProducts")), "###,##0")
                   .navUP
                 .elCreateSibling "gOSAPP"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_APPROS_OS_QtyItems")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_APPROS_OS_Value_Cost")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_APPROS_OS_Value_Retail")), "###,##0")
                    .navUP
                 .elCreateSibling "gTOT"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_OnHand_QtyItems")) + FNN(rs.Fields("STAT_APPROS_OS_QtyItems")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_ValueofStock_Cost")) + FNN(rs.Fields("STAT_APPROS_OS_Value_Cost")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_ValueofStock_Retail")) + FNN(rs.Fields("STAT_APPROS_OS_Value_Retail")), "###,##0")
                    .navUP
           Else
                .elCreateSibling "gBF"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("BF_STAT_OnHand_QtyItems")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("BF_STAT_ValueofStock_CostExVAT")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("BF_STAT_ValueofStock_RetailExVAT")), "###,##0")
                    .elCreateSibling "PQTY", True
                        .elText = Format(FNN(rs.Fields("BF_STAT_OnHand_QtyProducts")), "###,##0")
                    .navUP
                .elCreateSibling "gREC"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_DEL_QtyItems_mm")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_DEL_Value_COST_mmExVAT")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_DEL_Value_Retail_mmExVAT")), "###,##0")
                    .navUP
                .elCreateSibling "gSAL"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_INV_QtyItems_mm")) + FNN(rs.Fields("STAT_CS_QtyItems_mm")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_INV_Value_Cost_mmExVAT")) + FNN(rs.Fields("STAT_CS_Value_Cost_mmExVAT")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_INV_Value_Retail_mmExVAT")) + FNN(rs.Fields("STAT_CS_Value_Retail_mmExVAT")), "###,##0")
                    .navUP
                .elCreateSibling "gRET"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_CN_QtyItems_mm")) + FNN(rs.Fields("STAT_CSR_QtyItems_mm")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_CN_Value_Cost_mmExVAT")) + FNN(rs.Fields("STAT_CSR_Value_Cost_mmExVAT")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_CN_Value_Retail_mmExVAT")) + FNN(rs.Fields("STAT_CSR_Value_Retail_mmExVAT")), "###,##0")
                    .navUP
                .elCreateSibling "gAPP"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_APP_QtyItems_mm")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_APP_Value_Cost_mmExVAT")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_APP_Value_Retail_mmExVAT")), "###,##0")
                    .navUP
                .elCreateSibling "gAPPR"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_APPR_QtyItems_mm")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_APPR_Value_Cost_mmExVAT")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_APPR_Value_Retail_mmExVAT")), "###,##0")
                    .navUP
                .elCreateSibling "gTFROUT"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_TFROUT_QtyItems_mm")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_TFROUT_Value_Cost_mmExVAT")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_TFROUT_Value_Retail_mmExVAT")), "###,##0")
                    .navUP
                .elCreateSibling "gTFRIN"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_TFRIN_QtyItems_mm")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_TFRIN_Value_Cost_mmExVAT")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_TFRIN_Value_Retail_mmExVAT")), "###,##0")
                    .navUP
                .elCreateSibling "gRETS"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_RTN_QtyItems_mm")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_RTN_Value_Cost_mmExVAT")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_RTN_Value_Retail_mmExVAT")), "###,##0")
                    .navUP
                .elCreateSibling "gCF"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_OnHand_QtyItems")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_ValueofStock_CostExVAT")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_ValueofStock_RetailExVAT")), "###,##0")
                     .elCreateSibling "PQTY", True
                        .elText = Format(FNN(rs.Fields("STAT_OnHand_QtyProducts")), "###,##0")
                   .navUP
                 .elCreateSibling "gOSAPP"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_APPROS_OS_QtyItems")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_APPROS_OS_Value_CostExVAT")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_APPROS_OS_Value_RetailExVAT")), "###,##0")
                    .navUP
                 .elCreateSibling "gTOT"
                    .chCreate "IQTY"
                        .elText = Format(FNN(rs.Fields("STAT_OnHand_QtyItems")) + FNN(rs.Fields("STAT_APPROS_OS_QtyItems")), "###,##0")
                    .elCreateSibling "Cost", True
                        .elText = Format(FNN(rs.Fields("STAT_ValueofStock_CostExVAT")) + FNN(rs.Fields("STAT_APPROS_OS_Value_CostExVAT")), "###,##0")
                    .elCreateSibling "Retail", True
                        .elText = Format(FNN(rs.Fields("STAT_ValueofStock_RetailExVAT")) + FNN(rs.Fields("STAT_APPROS_OS_Value_RetailExVAT")), "###,##0")
                    .navUP
            End If
    End With
                            p 26
'FINALLY PRODUCE THE .XML FILE
    strXML = strWorkingFolder & "SOH_RECON_" & Format(dteReconDate, "dd-mm-yyyy") & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With
    
                        p 27
'WRITE THE .PDF FILE IF NECESSARY
'Stage 1 apply the .XSLT style sheet to the .XML and produce the .FO file
    If DispatchMode = enView Then
        Set objXSL = Nothing
        Set objXSL = New MSXML2.DOMDocument60
        objXSL.async = False
        objXSL.ValidateOnParse = False
        objXSL.resolveExternals = False
        strPath = oPC.SharedFolderRoot & "\Templates\SOH_RECON_DOC_FO.xsl"
        Set fs = New FileSystemObject
        If fs.FileExists(strPath) Then
            objXSL.Load strPath
        End If
                        p 6
        Set opXMLDOC = New MSXML2.DOMDocument60
        opXMLDOC.async = False
        opXMLDOC.ValidateOnParse = False
        opXMLDOC.resolveExternals = False
        xMLDoc.docObject.transformNodeToObject objXSL, opXMLDOC
        
        strFOFile = strWorkingFolder & "SOH_RECON" & Format(dteReconDate, "dd-mm-yyyy") & ".FO"
        strPDFFile = strWorkingFolder & "SOH_RECON" & Format(dteReconDate, "dd-mm-yyyy") & ".PDF"
                                p 7
        docWriteTostream strFOFile, opXMLDOC, "UNICODE"
        
'Stage 2 Convert the .FO file to .PDF and clean up
       ''' strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & strFOFile & " " & strPDFFile
        strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & strFOFile & " " & strPDFFile
        F_7_AB_1_ShellAndWaitSimple strCommand
        'IF IT FAILS HERE CHECK THAT JAVA IS INSTALLED
        'We do the following because the batch file executes and is not waited for by the F_7_AB_1_ShellAndWaitSimple routine
        'so the FO file is sometimes deleted before the .PDF file is produced.
        cnt = 0
        Do While Not fs.FileExists(strPDFFile) And cnt < 40
            MsgWaitObj 1000
            cnt = cnt + 1
        Loop
        If fs.FileExists(strFOFile) Then
            fs.DeleteFile strFOFile
        End If
    End If
    oTF.CloseTextFile
    ExportToXML = True
    
    '
    If DispatchMode = enView Then
        strExecutable = GetPDFExecutable(strPDFFile)
        If strExecutable = "" Then
            MsgBox "There is no application set on this computer to open the file: " & strPDFFile & ". The document cannot be displayed", vbOKOnly, "Can't do this"
        Else
            Shell strExecutable & " " & strPDFFile
        End If
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Reconciliation.ExportToXML(DispatchMode)", DispatchMode
End Function
Public Sub docWriteTostream(ByVal FilePath As String, obj As MSXML2.DOMDocument60, _
                Optional ByVal CharSet As String = "UNICODE")
    On Error GoTo errHandler
    Dim s As Object
    Set s = CreateObject("ADODB.Stream")
    With s
        If CharSet <> "" Then .CharSet = CharSet
        .Open
        .WriteText obj.xml
        .SaveToFile FilePath, 2 'adSaveCreateOverWrite
        .Close
    End With
    Exit Sub
errHandler:
    ErrorIn "a_PO.docWriteToFile(FilePath,Charset)", Array(FilePath, CharSet)
End Sub

