VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox DirList 
      Height          =   765
      Left            =   3795
      TabIndex        =   3
      Top             =   60
      Width           =   2910
   End
   Begin VB.DriveListBox drvList 
      Height          =   315
      Left            =   510
      TabIndex        =   2
      Top             =   150
      Width           =   3105
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2490
      Left            =   420
      TabIndex        =   1
      Top             =   990
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   4392
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdBuildXMLDoc 
      Caption         =   "Prepare document"
      Height          =   465
      Left            =   6525
      TabIndex        =   0
      Top             =   4110
      Width           =   2040
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   75
      Top             =   4020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuildXMLDoc_Click()
    BuildXMLDoc
End Sub

'Dim XMLDoc as int
Private Sub drvList_Change()
   On Error GoTo DriveHandler
   ' If new drive was selected, the Dir1 box
   ' updates its display.
   DirList.Path = drvList.Drive
   DirList.Path = DirList.List(-4)
   Exit Sub
' If there is an error, reset drvList.Drive with the
' drive from dirList.Path.
DriveHandler:
   drvList.Drive = DirList.Path
   
   Exit Sub
End Sub


Private Sub BuildXMLDoc()
    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "PO_DOC"
            .chCreate "MessageType"
                .elText = "PURCHASEORDER"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "DestinationAddress"
                If oPC.TestMode Then
                    .elText = oPC.EmailAddressForTesting
                    pDestinationEmailAddress = oPC.EmailAddressForTesting
                Else
                    If Me.Supplier Is Nothing Then
                        .elText = ""
                        pDestinationEmailAddress = ""
                    Else
                        .elText = Me.Supplier.OrderToAddress.Email
                        pDestinationEmailAddress = Me.Supplier.OrderToAddress.Email
                    End If
                End If
                        p 23
            .elCreateSibling "TemplateName"
                .elText = "PO_DOC"
            .elCreateSibling "SendersEmail"
                If oPC.EmailFrom > "" Then
                    .elText = oPC.EmailFrom
                Else
                    .elText = Me.SendersEmail
                End If
            .elCreateSibling "CopyCount"
                .elText = pQtyCopies
            .elCreateSibling "Printer"
                If Not oDC Is Nothing Then .elText = oDC.PrinterName
            .elCreateSibling "Status"
                .elText = StatusForPrinting
            .elCreateSibling "AccompanyingMessage"
                .elText = oPC.Configuration.EmailPOMsg
            .elCreateSibling "DispatchMethod"
                .elText = Me.Supplier.DispatchModes.ItemByOrdinalIndex(Me.Supplier.DispatchModes.FindIndexByKey(Me.DispatchModeID))
            .elCreateSibling "LogoPath"
                If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP") Then
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP"
                Else
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
                End If
            .elCreateSibling "DocCode"
                .elText = Me.DOCCode
            .elCreateSibling "DocDate", True
                .elText = Me.DOCDate
            .elCreateSibling "Sender", True
                .elText = oPC.Configuration.DefaultCompany.CompanyName
            .elCreateSibling "SenderAddress", True
                .elText = Replace(oPC.Configuration.DefaultCompany.StreetAddress, Chr(13) & Chr(10), Chr(10))
            .elCreateSibling "SupplierName", True
                .elText = Supplier.NameAndCode(35)
            .elCreateSibling "SupplierWithAddress", True
                If Supplier.OrderToAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = Replace(Supplier.OrderToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
                End If
            .elCreateSibling "SupplierPhone", True
                If Supplier.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = IIf(Me.Supplier.BillTOAddress.Phone > "", "Phone: " & Supplier.BillTOAddress.Phone, "")
                End If
            .elCreateSibling "SupplierFax", True
                If Supplier.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = IIf(Supplier.BillTOAddress.Fax > "", "Fax: " & Supplier.BillTOAddress.Fax, "")
                End If
            .elCreateSibling "ACNO"
                .elText = IIf(Me.Supplier.AcNo > "", "Ac/no. " & Me.Supplier.AcNo, "")
            .elCreateSibling "BillTo", True
                If Me.BillToAddressID > 0 Then
                    If Not oPC.Configuration.Stores.FindStoreByID(Me.BillToAddressID) Is Nothing Then
                        .elText = Replace(oPC.Configuration.Stores.FindStoreByID(Me.BillToAddressID).BillAddress, Chr(13) & Chr(10), Chr(10))
                    Else
                        .elText = Replace(oPC.Configuration.DefaultStore.BillAddress, Chr(13) & Chr(10), Chr(10))
                    End If
                End If
            .elCreateSibling "DelTo", True
                If Me.DELTOStoreID > 0 Then
                    If Not oPC.Configuration.Stores.FindStoreByID(Me.DELTOStoreID) Is Nothing Then
                        .elText = Replace(oPC.Configuration.Stores.FindStoreByID(Me.DELTOStoreID).DelAddress, Chr(13) & Chr(10), Chr(10))
                    Else
                        .elText = Replace(oPC.Configuration.DefaultStore.DelAddress, Chr(13) & Chr(10), Chr(10))
                    End If
                End If
                            p 24
            
            For i = 1 To Me.POLines.Count
                    .elCreateSibling "DetailLine", True
                    .chCreate "SKU"
                    .elText = POLines.Item(i).ProductCodeForExport
                    .elCreateSibling "Title", True
                    If POLines(i).Fulfilled <> "CAN" Then
                       .elText = POLines(i).TitleAuthor
                    Else
                       .elText = "***CANCELLED***" & POLines(i).TitleAuthor
                    End If
                    .elCreateSibling "QtyFirm", True
                        .elText = POLines(i).QtyFirmF
                    .elCreateSibling "QtySS", True
                        .elText = POLines(i).QtySSF
                    .elCreateSibling "Price", True
                        .elText = POLines(i).PriceF(bForeign)
                    .elCreateSibling "DiscountRate", True
                        .elText = POLines(i).DiscountF
                    .elCreateSibling "Reference", True
                        .elText = POLines(i).Ref
                    .elCreateSibling "Extension", True
                        .elText = POLines(i).PLessDiscExtF(bForeign)
                    .elCreateSibling "Note", True
                        .elText = POLines(i).Note
                    .navUP
            Next i
                            p 25
            .elCreateSibling "TotalNumberOfLines", True
                .elText = CStr(Me.POLines.Count)
            .elCreateSibling "TotalText", True
                .elText = "Total"
            .elCreateSibling "TotalNumbers", True
                .elText = TotalPayableF(bForeign)
            .elCreateSibling "Memo", True
                .elText = Memo
            .elCreateSibling "CompanyRegistration", True
                .elText = oPC.Configuration.DefaultCompany.CoRegistrationNumber
            .elCreateSibling "VATNumber", True
                .elText = oPC.Configuration.DefaultCompany.VatNumber
            .elCreateSibling "StaffMember", True
                If oPC.Configuration.Staff.FindStaffByID(Me.StaffID) Is Nothing Then
                    .elText = ""
                Else
                    .elText = oPC.Configuration.Staff.FindStaffByID(Me.StaffID).StaffName
                End If
            .elCreateSibling "OrderMessage", True
                .elText = oPC.Configuration.OrderText
    End With
                            p 26
'FINALLY PRODUCE THE .XML FILE
    strXML = strWorkingFolder & "PO_" & Me.DOCCode & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With

End Sub
