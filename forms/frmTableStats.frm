VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmTableSTats 
   Caption         =   "Table statistics"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   5820
      Left            =   180
      OleObjectBlob   =   "frmTableStats.frx":0000
      TabIndex        =   0
      Top             =   420
      Width           =   9615
   End
End
Attribute VB_Name = "frmTableSTats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Component(rs As ADODB.Recordset)
  '  Set d1.Recordset = rs
  rs.Sort = "DataSpaceUsed"
    G1.DataSource = rs
    
End Sub


