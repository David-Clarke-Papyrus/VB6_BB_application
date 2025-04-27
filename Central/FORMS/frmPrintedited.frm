VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintedited 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Edited loyalty customer files"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   6510
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   4755
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   1380
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2970
      Left            =   150
      TabIndex        =   1
      Top             =   60
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   5239
      View            =   3
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483635
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   4304
      EndProperty
   End
End
Attribute VB_Name = "frmPrintedited"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oTF As z_TextFile
Dim arEDCust As arEditedCustomers

Private Sub cmdPrint_Click()
Dim rs As New ADODB.Recordset
Dim ar() As String
Dim strLine As String

    ReDim ar(30)
    rs.Fields.Append "E_ACNO", adVarChar, 100
    rs.Fields.Append "E_LastName", adVarChar, 100
    rs.Fields.Append "E_FirstName", adVarChar, 100
    rs.Fields.Append "E_Title", adVarChar, 100
    rs.Fields.Append "E_Phone", adVarChar, 100
    rs.Fields.Append "E_Phone2", adVarChar, 100
    rs.Fields.Append "E_Cell", adVarChar, 100
    rs.Fields.Append "E_Address", adVarChar, 100
    rs.Fields.Append "E_PostCode", adVarChar, 100
    rs.Fields.Append "E_Country", adVarChar, 100
    rs.Fields.Append "E_Email", adVarChar, 100
    rs.Fields.Append "E_Launch", adVarChar, 100
    rs.Fields.Append "E_PromotionYN", adVarChar, 100
    rs.Fields.Append "E_SaleYN", adVarChar, 100
    rs.Open
    Set oTF = New z_TextFile
    oTF.OpenTextFileToRead oPC.SharedFolderRoot & "\Data\Loyalty\Edited\" & lvw.SelectedItem.Text
    Do While Not oTF.IsEOF
        strLine = oTF.ReadLinefromTextFile
        ar = Split(strLine, vbTab)
        rs.AddNew
        rs.Fields("E_ACNO") = FNS(ar(17))
        rs.Fields("E_LastName") = FNS(ar(0))
        rs.Fields("E_FirstName") = FNS(ar(1))
        rs.Fields("E_Title") = FNS(ar(2))
        rs.Fields("E_Phone") = FNS(ar(15))
        rs.Fields("E_Phone2") = FNS(ar(16))
        rs.Fields("E_Cell") = FNS(ar(5))
        rs.Fields("E_Address") = Trim(ar(10) & vbCrLf & ar(11) & vbCrLf & ar(12) & vbCrLf & ar(13))
        rs.Fields("E_PostCode") = FNS(ar(8))
        rs.Fields("E_Country") = ar(22)
        rs.Fields("E_Email") = FNS(ar(18))
        rs.Fields("E_Launch") = IIf(ar(19) = "1", "Y", "N")
        rs.Fields("E_PromotionYN") = IIf(ar(20) = "1", "Y", "N")
        rs.Fields("E_SaleYN") = IIf(ar(21) = "1", "Y", "N")
        rs.Update
    Loop
    oTF.CloseTextFile
    Set oTF = Nothing
    rs.MoveFirst
    Set arEDCust = New arEditedCustomers
    arEDCust.Component rs
    arEDCust.Show vbModal
    rs.Close
    Set rs = Nothing
    Set arEDCust = Nothing
    
End Sub

Private Sub Form_Load()
Dim oFSO As FileSystemObject
Dim fol, fc, f
Dim itm As ListItem

    Set oFSO = New FileSystemObject
    Set fol = oFSO.GetFolder(oPC.SharedFolderRoot & "\Data\Loyalty\Edited")
    Set fc = fol.Files
    For Each f In fc
        Set itm = lvw.ListItems.Add
        itm.Text = f.Name
        itm.SubItems(1) = f.DateCreated
        Set itm = Nothing
    Next
End Sub

Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub
