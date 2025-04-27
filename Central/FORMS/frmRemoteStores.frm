VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmRemoteStores 
   Caption         =   "Stores"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   480
      Left            =   1695
      TabIndex        =   4
      Top             =   90
      Width           =   4425
      Begin VB.OptionButton optCurrencies 
         Caption         =   "Currencies"
         Height          =   210
         Left            =   2730
         TabIndex        =   7
         Top             =   165
         Width           =   1350
      End
      Begin VB.OptionButton optCountries 
         Caption         =   "Countries"
         Height          =   210
         Left            =   1575
         TabIndex        =   6
         Top             =   165
         Width           =   1350
      End
      Begin VB.OptionButton optStoreCodes 
         Caption         =   "Store codes"
         Height          =   210
         Left            =   210
         TabIndex        =   5
         Top             =   165
         Width           =   1350
      End
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   450
      Left            =   7275
      TabIndex        =   3
      Top             =   120
      Width           =   780
   End
   Begin VB.TextBox txtStoreCode 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   165
      Width           =   660
   End
   Begin TrueOleDBGrid60.TDBGrid Grid 
      Height          =   3825
      Left            =   195
      OleObjectBlob   =   "frmRemoteStores.frx":0000
      TabIndex        =   0
      Top             =   990
      Width           =   7890
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Store"
      Height          =   300
      Left            =   75
      TabIndex        =   2
      Top             =   225
      Width           =   690
   End
End
Attribute VB_Name = "frmRemoteStores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
