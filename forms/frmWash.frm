VERSION 5.00
Begin VB.Form frmWash 
   BackColor       =   &H00D3D3CB&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import and Export"
   ClientHeight    =   6465
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      ForeColor       =   &H8000000D&
      Height          =   4815
      Left            =   270
      TabIndex        =   2
      Top             =   450
      Width           =   2745
      Begin VB.CheckBox optBookstatus 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Book status"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   4050
         Width           =   1635
      End
      Begin VB.CheckBox optBIC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "B.I.C."
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Top             =   3780
         Width           =   1635
      End
      Begin VB.CheckBox optRRP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "R.R.P."
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   3480
         Width           =   1635
      End
      Begin VB.CheckBox optUKPrice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "U.K. price"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Top             =   3195
         Width           =   1635
      End
      Begin VB.CheckBox optPublicationDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Publication date"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   2910
         Width           =   1635
      End
      Begin VB.CheckBox optSeriesTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Series title"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   2625
         Width           =   1635
      End
      Begin VB.CheckBox optPublishername 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Puiblisher name"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   2340
         Width           =   1635
      End
      Begin VB.CheckBox optSupplierCode 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Supplier code"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   2055
         Width           =   1635
      End
      Begin VB.CheckBox optEdition 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Edition"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   1770
         Width           =   1635
      End
      Begin VB.CheckBox optBinding 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Binding"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   1485
         Width           =   1635
      End
      Begin VB.CheckBox optAvailability 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Availability"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   1185
         Width           =   1635
      End
      Begin VB.CheckBox optSubtitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Subtitle"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   900
         Width           =   1635
      End
      Begin VB.CheckBox optTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Title"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   615
         Width           =   1635
      End
      Begin VB.CheckBox optAuthor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Author"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   330
         Width           =   1635
      End
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Start"
      CausesValidation=   0   'False
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1140
      Picture         =   "frmWash.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5550
      Width           =   1000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Select which fields to update"
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
      Height          =   300
      Left            =   150
      TabIndex        =   1
      Top             =   180
      Width           =   2775
   End
End
Attribute VB_Name = "frmWash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim bCancelled As Boolean
Public Property Get Author()
    Author = Me.optAuthor
End Property

Public Property Get Title()
    Title = Me.optTitle
End Property

Public Property Get Subtitle()
    Subtitle = Me.optSubtitle
End Property

Public Property Get Availability()
    optAvailability = Me.optAvailability
End Property

Public Property Get Binding()
    Binding = Me.optBinding
End Property

Public Property Get Edition()
    Edition = Me.optEdition
End Property

Public Property Get SupplierCode()
    SupplierCode = Me.optSupplierCode
End Property

Public Property Get Publishername()
    Publishername = Me.optPublishername
End Property
Public Property Get SeriesTitle()
    SeriesTitle = Me.optSeriesTitle
End Property
Public Property Get PublicationDate()
    PublicationDate = Me.optPublicationDate
End Property
Public Property Get UKPrice()
    UKPrice = Me.optUKPrice
End Property

Public Property Get RRP()
    RRP = Me.optRRP
End Property
Public Property Get BIC()
    BIC = Me.optBIC
End Property
Public Property Get BookStatus()
    BookStatus = Me.optBookstatus
End Property

Private Sub Form_Load()
    bCancelled = True
End Sub

Private Sub OKButton_Click()
    bCancelled = False
    Me.Hide

End Sub

Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property

