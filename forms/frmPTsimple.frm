VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPTsimple 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Product types"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   3345
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      Picture         =   "frmPTsimple.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4380
      Width           =   1000
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   4170
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   7355
      View            =   3
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Description"
         Object.Width           =   4304
      EndProperty
   End
End
Attribute VB_Name = "frmPTsimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oPT As a_PT
Attribute oPT.VB_VarHelpID = -1
Dim tlProductTypes As z_TextList
Dim strCRSALES As String
Dim strCRSALES_CONTRA As String
Dim strCASHSALES As String
Dim strCASHSALES_CONTRA As String
Dim strPURCHASES As String
Dim strPURCHASES_CONTRA As String
Dim strVAT As String

Public Property Get CRSALES() As String
    CRSALES = strCRSALES
End Property
Public Property Get CRSALES_CONTRA() As String
    CRSALES_CONTRA = strCRSALES_CONTRA
End Property
Public Property Get CASHSALES() As String
    CASHSALES = strCASHSALES
End Property
Public Property Get CASHSALES_CONTRA() As String
    CASHSALES_CONTRA = strCASHSALES_CONTRA
End Property
Public Property Get PURCHASES() As String
    PURCHASES = strPURCHASES
End Property
Public Property Get PURCHASES_CONTRA() As String
    PURCHASES_CONTRA = strPURCHASES_CONTRA
End Property
Public Property Get VAT() As String
    VAT = strVAT
End Property
Private Sub cmdClose_Click()
    If lvw.SelectedItem.Index < 1 Then
        Me.Hide
        Exit Sub
    End If
    Set oPT = New a_PT
    oPT.Load tlProductTypes.Key(lvw.SelectedItem.text)
    If oPT.PTID > 0 Then
        strCRSALES = oPT.CRSALES
        strCRSALES_CONTRA = oPT.CRSALES_CONTRA
        strCASHSALES = oPT.CASHSALES
        strCASHSALES_CONTRA = oPT.CASHSALES_CONTRA
        strPURCHASES = oPT.PURCHASES
        strPURCHASES_CONTRA = oPT.PURCHASES_CONTRA
        strVAT = oPT.VAT
    Else
        MsgBox "Select an item.", vbInformation, "Status"
    End If
    Set oPT = Nothing
    Me.Hide
End Sub

Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub


Private Sub Form_Load()
    If Me.WindowState <> 2 Then
        TOP = 850
        Left = 550
        Width = 3400
        Height = 5400
    End If
    Set tlProductTypes = New z_TextList
    tlProductTypes.Load ltProductType
    LoadListView
End Sub

Private Sub LoadListView()
    Set tlProductTypes = Nothing
    Set tlProductTypes = New z_TextList
    tlProductTypes.Load ltProductType
    LoadList
End Sub
Private Sub LoadList()
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvw.ListItems.Clear
    For i = 1 To tlProductTypes.Count
        Set objItm = Me.lvw.ListItems.Add
        With objItm
            .text = tlProductTypes.ItemByOrdinalIndex(i)
            .Bold = tlProductTypes.f4ByOrdinalIndex(i) = "True"
        End With
    Next i
    
End Sub

