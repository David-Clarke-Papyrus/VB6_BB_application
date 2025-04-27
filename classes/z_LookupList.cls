VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_CommonLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim tlSUppliers As New z_TextList
Dim tlCategories As New z_TextList
Dim tlCatalogues As New z_TextList
Public Property Get Supplierlist() As z_TextList
    Set Supplierlist = tlSUppliers
End Property

Public Function Categorylist() As z_TextList
    Set Categorylist = tlCategories
End Function

Public Function Cataloguelist() As z_TextList
    Set Cataloguelist = tlCatalogues
End Function

Private Sub Class_Initialize()
    tlSUppliers.Load ("getsupplier")
    tlCategories.Load ("getcategory")
    tlCatalogues.Load ("getcatalogue")
End Sub

Private Sub Class_Terminate()
    Set tlSUppliers = Nothing
    Set tlCategories = Nothing
    Set tlCatalogues = Nothing
End Sub
