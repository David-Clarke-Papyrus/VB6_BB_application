VERSION 5.00
Object = "{CCA2C66D-33FD-11D5-8D72-005004532BDF}#1.3#0"; "CCubeX.ocx"
Begin VB.Form frmGLAccounts 
   Caption         =   "Accounts"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   9030
   Begin CCubeX.ContourCubeX CC 
      Height          =   2700
      Left            =   885
      TabIndex        =   0
      Top             =   1380
      Width           =   4500
      BackColor       =   14215660
      Enabled         =   -1  'True
      MainAxis        =   0
      DataSourceType  =   0
      ConnectionString=   ""
      SQL             =   ""
      PreGrouping     =   -1  'True
      Active          =   0   'False
      HDrillDownLevel =   -1
      VDrillDownLevel =   1
      Transposed      =   0   'False
      SuppressZeroRows=   0   'False
      SuppressZeroCols=   0   'False
      ViewFlags       =   0
      BorderStyle     =   1
      AllowInactiveDimArea=   -1  'True
      AllowFilter     =   -1  'True
      AllowExpand     =   -1  'True
      AllowPivot      =   -1  'True
      AllowTitle      =   -1  'True
      ShowAsPercent   =   0
      TotalsString    =   ""
      CubeTitle       =   ""
      TitleAlign      =   0
      TitleBkColor    =   14215660
      DimBkColor      =   14215660
      DimTitleBkColor =   14898176
      DimTitleInactiveBkColor=   10070188
      DimFilterBkColor=   14215660
      InactiveDimAreaBkColor=   14215660
      HeadingBkColor  =   14215660
      DataGridColor   =   14215660
      DataBkColor     =   16777215
      TotalBkColor    =   14679807
      GrandTotalBkColor=   14679807
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DimFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DimTitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DimFilterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DataFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TotalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GrandTotalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
      Object.Visible         =   -1  'True
      MousePointer    =   0
      TitleForeColor  =   0
      DimForeColor    =   0
      DimTitleForeColor=   16777215
      DimFilterForeColor=   0
      HeadingForeColor=   0
      DataForeColor   =   0
      TotalForeColor  =   -2147483640
      GrandTotalForeColor=   -2147483640
      UnusedDataAreaColor=   -2147483643
      MainAxisDim     =   ""
      DimTitleDragBkColor=   32768
      FactsCaption    =   "Facts"
      ShowFactsBitmap =   -1  'True
      ADOCursorLocation=   2
      AutoRefreshView =   0   'False
      FPErrString     =   "FPErr"
      NULLValueString =   ""
      NonExistentValueString=   ""
      DefaultFactFormat=   "###,###,###,###,###,##0.00"
      AllowFactFilter =   -1  'True
      VERSION_NO      =   2
      FIELDS_SETTINGS =   $"frmGLAccounts.frx":0000
   End
End
Attribute VB_Name = "frmGLAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

