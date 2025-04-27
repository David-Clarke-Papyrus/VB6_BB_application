VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Schools ordering management"
   ClientHeight    =   6180
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   11610
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5925
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14817
            MinWidth        =   14817
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuBrowse 
      Caption         =   "&Browse"
      Begin VB.Menu mnuInvoice 
         Caption         =   "&Invoice"
      End
      Begin VB.Menu mnuImportOrdersheet 
         Caption         =   "&Order sheet"
      End
      Begin VB.Menu mnuInvoices 
         Caption         =   "&Invoices"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fs As FileSystemObject
Dim frmBrowseInvoices As frmBrowseInvoices


Private Sub Form_Load()
    
'    Set fs = New FileSystemObject
'    strPath = oPC.SharedFolderRoot & IIf(Right(oPC.SharedFolderRoot, 1) = "\", "", "\") & "Stocktke"
'    If Not fs.FolderExists(strPath) Then
'        fs.CreateFolder (strPath)
'    End If
'    LoadExisting
'    bCaptureQuantities = False
'    mnuQty.Checked = 0

End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

'Private Sub mnuImportCustomerList_Click()
'Dim frm As frmSchoolsCustomerList
'
'    Set frm = New frmSchoolsCustomerList
'    frm.Show
'    Screen.MousePointer = vbDefault
'
'End Sub
'
'Private Sub mnuImportOrdersheet_Click()
'Dim frm As frmSchoolsOrderList
'
'    Set frm = New frmSchoolsOrderList
'    frm.Show
'    Screen.MousePointer = vbDefault
'End Sub

Private Sub mnuInvoice_Click()
BrowseInvoices
End Sub

'Private Sub mnuInvoices_Click()
'Dim frm As frmCustomerOrderList
'
'    Set frm = New frmCustomerOrderList
'    frm.Show
'    Screen.MousePointer = vbDefault
'End Sub
Private Sub BrowseInvoices()
    If frmBrowseInvoices Is Nothing Then
       Set frmBrowseInvoices = New frmBrowseInvoices
    End If
    frmBrowseInvoices.ZOrder 0
End Sub
