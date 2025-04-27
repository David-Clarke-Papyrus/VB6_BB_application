VERSION 5.00
Begin VB.Form frmProductNB 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Edit general (non-book)  product"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9690
   ControlBox      =   0   'False
   Icon            =   "frmProductNB.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleMode       =   0  'User
   ScaleWidth      =   12782.56
   Begin VB.CheckBox chkExSales 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Exclude from sales reporting"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   5010
      TabIndex        =   42
      Top             =   6420
      Width           =   2955
   End
   Begin VB.CommandButton cmdSupplier 
      BackColor       =   &H00C4BCA4&
      Caption         =   "· · ·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8925
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4080
      Width           =   570
   End
   Begin VB.CommandButton cmdChangeType 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Change this product type to a book"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   -30
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   7275
      Width           =   2955
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Sections"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1350
      Left            =   5040
      TabIndex        =   36
      Top             =   4920
      Width           =   4215
      Begin VB.TextBox txtSection 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   165
         MultiLine       =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   720
         Width           =   3870
      End
      Begin VB.CommandButton cmdAddSection 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Add"
         Height          =   315
         Left            =   2655
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   330
         Width           =   750
      End
      Begin VB.ComboBox cboSection 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   165
         TabIndex        =   18
         Top             =   330
         Width           =   2490
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Codes and numbers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1800
      Left            =   270
      TabIndex        =   32
      Top             =   240
      Width           =   8730
      Begin VB.CheckBox chkNonStock 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Non stock-take item (e.g. newspaper )"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   480
         Left            =   1110
         TabIndex        =   2
         Top             =   1215
         Width           =   3525
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   990
         Left            =   3135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "frmProductNB.frx":030A
         Top             =   390
         Width           =   5415
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1110
         TabIndex        =   1
         Top             =   810
         Width           =   1680
      End
      Begin VB.TextBox txtEAN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   1110
         TabIndex        =   0
         Top             =   420
         Width           =   1680
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   390
         TabIndex        =   34
         Top             =   840
         Width           =   660
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "E.A.N."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   180
         TabIndex        =   33
         Top             =   465
         Width           =   870
      End
   End
   Begin VB.TextBox txtCost 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6015
      TabIndex        =   15
      Top             =   3000
      Width           =   1380
   End
   Begin VB.TextBox txtSP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6015
      TabIndex        =   14
      Top             =   2625
      Width           =   1380
   End
   Begin VB.TextBox txtRRP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6015
      TabIndex        =   13
      Top             =   2250
      Width           =   1380
   End
   Begin VB.CheckBox chkObsolete 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Obsolete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   480
      Left            =   1875
      TabIndex        =   12
      Top             =   6405
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Suppliers' status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1305
      Left            =   1470
      TabIndex        =   8
      Top             =   5025
      Width           =   2280
      Begin VB.OptionButton optRP 
         BackColor       =   &H00D3D3CB&
         Caption         =   "On backorder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   270
         TabIndex        =   11
         Top             =   945
         Width           =   1575
      End
      Begin VB.OptionButton optOOP 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Unavailable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   270
         TabIndex        =   10
         Top             =   630
         Width           =   1575
      End
      Begin VB.OptionButton optIP 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Available"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   270
         TabIndex        =   9
         Top             =   315
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdSetDefault 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Default V.A.T. rate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7470
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3555
      Width           =   1755
   End
   Begin VB.TextBox txtVAT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6000
      TabIndex        =   16
      Top             =   3585
      Width           =   1380
   End
   Begin VB.ComboBox cboProductType 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1470
      TabIndex        =   7
      Top             =   4500
      Width           =   2565
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8295
      Picture         =   "frmProductNB.frx":03B0
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6900
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7350
      Picture         =   "frmProductNB.frx":093A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6915
      Width           =   930
   End
   Begin VB.TextBox txtErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   1020
      Left            =   3735
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   6435
      Width           =   2955
   End
   Begin VB.TextBox txtEdition 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1485
      TabIndex        =   6
      Top             =   4050
      Width           =   2520
   End
   Begin VB.TextBox txtPublisher 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1485
      TabIndex        =   5
      Top             =   3615
      Width           =   2520
   End
   Begin VB.TextBox txtSubtitle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1485
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2880
      Width           =   3225
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   1485
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2250
      Width           =   3225
   End
   Begin VB.Label lblSupplier 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   6015
      TabIndex        =   41
      Top             =   4050
      Width           =   2880
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   5130
      TabIndex        =   40
      Top             =   4110
      Width           =   810
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   5205
      TabIndex        =   31
      Top             =   3000
      Width           =   750
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "S.P."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   5205
      TabIndex        =   30
      Top             =   2640
      Width           =   750
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "R.R.P."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   5205
      TabIndex        =   29
      Top             =   2265
      Width           =   750
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "V.A.T. Rate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   4830
      TabIndex        =   28
      Top             =   3630
      Width           =   1080
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Product type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   330
      TabIndex        =   27
      Top             =   4545
      Width           =   1080
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   765
      TabIndex        =   25
      Top             =   4080
      Width           =   645
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   150
      TabIndex        =   24
      Top             =   3660
      Width           =   1260
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   765
      TabIndex        =   23
      Top             =   2895
      Width           =   645
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   195
      TabIndex        =   22
      Top             =   2295
      Width           =   1215
   End
End
Attribute VB_Name = "frmProductNB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents oProd As a_Product
Attribute oProd.VB_VarHelpID = -1
Dim flgLoading As Boolean
Dim mCancel As Boolean
Dim XA As XArrayDB
Dim frmPrevious As Form

Sub Component(pProduct As a_Product, Optional pPrevForm As Form)
    Set frmPrevious = pPrevForm
    Set oProd = pProduct
    oProd.BeginEdit
    oProd.SetGeneralProduct
    oProd.GetStatus
End Sub


'Private Sub cboSection_Click()
'    If flgLoading Then Exit Sub
'    oProd.SetSection cboSection
'    txtSection = oProd.Section
'End Sub
Private Sub cboProductType_Click()
    If flgLoading Then Exit Sub
    oProd.SetProductTypeID oPC.Configuration.ProductTypes.Key(cboProductType)
End Sub


Private Sub chkExSales_Click()
    oProd.ExcludeFromSales = IIf(Me.chkExSales = 1, True, False)
End Sub

Private Sub chkNonStock_Click()
    If chkNonStock Then
        oProd.SetMagsEtc
    Else
        oProd.SetGeneralProduct
    End If
End Sub

'Private Sub chkNonStock_Click()
'    oProd.NonStock = chkNonStock
'End Sub

Private Sub chkObsolete_Click()
    oProd.Obsolete = chkObsolete
End Sub



Private Sub cmdDelete_Click()

    If MsgBox("Confirm deletion of product: " & oProd.Title & vbCrLf & "This is permanent!", vbOKCancel + vbExclamation, "Confirm") = vbOK Then
        oProd.Delete
        oProd.ApplyEdit
        Unload Me
    End If


End Sub

'Private Sub cmdAddSection_Click()
'    oProd.SetSection cboSection
'    txtSection = oProd.Section
'End Sub

'Private Sub cmdGenerateEAN_Click()
'Dim oProdCode As New z_ProdCode
'    oProdCode.SetCodesForBook txtCode
'    oProd.SetEAN oProdCode.EAN
'    txtEAN = oProd.EAN
'End Sub
'

Private Sub cmdChangeType_Click()
    If MsgBox("You want to change this product to be a book?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    Else
        oProd.SetProductType "B"
        oProd.ApplyEdit
        Unload Me
    End If
End Sub

Private Sub cmdSetDefault_Click()
    Me.txtVAT = oPC.Configuration.vatRate
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If oProd.IsEditing Then oProd.CancelEdit
End Sub


Private Sub oProd_Valid(strMsg As String)
    Me.txtErrors = strMsg
    Me.cmdOK.Enabled = (strMsg = "")
End Sub
Private Sub cmdCancel_Click()
    oProd.CancelEdit
    Unload Me
End Sub

Private Sub cmdNewCode_Click()
    Me.txtCode = "#"
    oProd.SetCode "#"
End Sub

Private Sub cmdOK_Click()
Dim lngResult As Long
Dim strMsg As String
Dim frmPreview As frmProductNBPrev

    WaitMsg "Saving product . . .", True, Me
    oProd.SetBFDistributorCode "XXX"
    oProd.ApplyEdit lngResult, strMsg
    If lngResult = 99 Then
        WaitMsg "", False, Me
        If strMsg = "DUPLICATE" Then
            MsgBox "Invalid values - check that the code is has not been already used", vbInformation, "Save failed"
        ElseIf strMsg = "TIMEOUT" Then
            MsgBox "The operation has timed out. The record is probably locked by another user." & vbCrLf & "Cancel your update or try again. ", vbInformation, "Save failed"
        End If
    Else
        If frmPrevious Is Nothing Then
            Set frmPreview = New frmProductNBPrev
        Else
            Set frmPreview = frmPrevious
        End If
        frmPreview.Component oProd
        frmPreview.RefreshForm
        frmPreview.Show
        WaitMsg "", False, Me
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    left = 10
    top = 10
    Width = 10000
    Height = 8000
    LoadControls
   ' Me.txtEAN.Enabled = left(txtEAN, 1) <> "2" 'only enable standard
End Sub
Private Sub LoadControls()
    flgLoading = True
    txtCode = oProd.Code
    Me.txtEAN = oProd.EAN
    txtTitle = oProd.Title
    txtSubtitle = oProd.Subtitle
    txtEdition = oProd.Edition
    txtPublisher = oProd.Publisher
    txtRRP = oProd.RRPF
    txtSP = oProd.SPF
    txtCost = oProd.CostF
    txtSection = oProd.Section
    Me.txtVAT = oProd.vatratef
    LoadCombo cboSection, oPC.Configuration.Sections
    LoadCombo cboProductType, oPC.Configuration.ProductTypes
    cboProductType = oPC.Configuration.ProductTypes.Item(oProd.ProductTypeID)
    Me.chkNonStock = IIf(oProd.isNonStock, 1, 0)
    Me.chkObsolete = IIf(oProd.Obsolete, 1, 0)
    Me.chkExSales = IIf(oProd.ExcludeFromSales, 1, 0)
    
    Select Case oProd.Status
    Case "O"
        optOOP.Value = True
    Case "R"
        optRP.Value = True
    Case Else
        optIP.Value = True
    End Select
    flgLoading = False
End Sub

Private Sub optIP_Click()
    oProd.SetStatus enInPrint
End Sub

Private Sub optOOP_Click()
    oProd.SetStatus enOutOfPrint
End Sub

Private Sub optRP_Click()
    oProd.SetStatus enAwaitingReprint
End Sub

Private Sub txtCode_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetCode(txtCode)
    If Err Then
      Beep
      intPos = txtCode.SelStart
      txtCode = oProd.Code
      txtCode.SelStart = intPos - 1
    End If
End Sub
Private Sub txtCode_Validate(Cancel As Boolean)
    Cancel = Not oProd.SetCode(txtCode)
End Sub

Private Sub txtEAN_GotFocus()
    AutoSelect txtEAN
End Sub
Private Sub txtEAN_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetEAN(txtEAN)
    If Err Then
      Beep
      intPos = txtEAN.SelStart
      txtEAN = oProd.EAN
      txtEAN.SelStart = intPos - 1
    End If
End Sub

Private Sub txtEAN_Validate(Cancel As Boolean)
    Cancel = Not oProd.SetEAN(txtEAN)
End Sub


Private Sub txtRRP_GotFocus()
    txtRRP = oProd.RRP
    AutoSelect txtRRP
End Sub
Private Sub txtRRP_Validate(Cancel As Boolean)
    If flgLoading Then Exit Sub
    If Not oProd.SetRRP(txtRRP) Then
        Cancel = True
    End If
    txtRRP = oProd.RRPF
End Sub

'Private Sub txtSection_Validate(Cancel As Boolean)
'    oProd.SetSectionAll txtSection
'    txtSection = oProd.Section
'End Sub

Private Sub txtSP_GotFocus()
    txtSP = oProd.SP
    AutoSelect txtSP
End Sub
Private Sub txtSP_Validate(Cancel As Boolean)
    If flgLoading Then Exit Sub
    If Not oProd.SetSP(txtSP) Then
        Cancel = True
    End If
    txtSP = oProd.SPF
End Sub
Private Sub txtCost_GotFocus()
    txtCost = oProd.Cost
    AutoSelect txtCost
End Sub
Private Sub txtCost_Validate(Cancel As Boolean)
    If flgLoading Then Exit Sub
    If Not oProd.SetCost(txtCost) Then
        Cancel = True
    End If
    txtCost = oProd.CostF
End Sub
'Private Sub txtSpecialPrice_GotFocus()
'    txtSpecialPrice = oProd.SpecialPrice
'    AutoSelect txtSpecialPrice
'End Sub
'Private Sub txtSpecialPrice_Validate(Cancel As Boolean)
'    If flgLoading Then Exit Sub
'    If Not oProd.setspecialPrice(txtSpecialPrice) Then
'        Cancel = True
'    End If
'    txtSpecialPrice = oProd.SpecialPriceF
'End Sub

Private Sub txtSubtitle_LostFocus()
    If flgLoading Then Exit Sub
    txtSubtitle = oProd.Subtitle
End Sub
Private Sub txtSubtitle_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtSubtitle_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetSubTitle(txtSubtitle)
    If Err Then
      Beep
      intPos = txtSubtitle.SelStart
      txtSubtitle = oProd.Subtitle
      txtSubtitle.SelStart = intPos - 1
    End If
End Sub


Private Sub txtTitle_LostFocus()
    If flgLoading Then Exit Sub
    txtTitle = oProd.Title
End Sub
Private Sub txtTitle_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtTitle_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetTitle(txtTitle)
    If Err Then
      Beep
      intPos = txtTitle.SelStart
      txtTitle = oProd.Title
      txtTitle.SelStart = intPos - 1
    End If
End Sub
Private Sub txtPublisher_LostFocus()
    If flgLoading Then Exit Sub
    txtPublisher = oProd.Publisher
End Sub
Private Sub txtPublisher_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtPublisher_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetPublisher(txtPublisher)
    If Err Then
      Beep
      intPos = txtPublisher.SelStart
      txtPublisher = oProd.Publisher
      txtPublisher.SelStart = intPos - 1
    End If
End Sub
Private Sub txtEdition_LostFocus()
    If flgLoading Then Exit Sub
    txtEdition = oProd.Edition
End Sub
Private Sub txtEdition_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtEdition_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetEdition(txtEdition)
    If Err Then
      Beep
      intPos = txtEdition.SelStart
      txtEdition = oProd.Edition
      txtEdition.SelStart = intPos - 1
    End If
End Sub

