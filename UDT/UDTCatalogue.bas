Attribute VB_Name = "UDTCatalogue"
Public Type CatalogueProps
  ID  As Long
  Serial As Integer
  DateFirstPrinted As Date
  Description As String * 50
  IsNew As Boolean
  IsDirty As Boolean
  IsDeleted As Boolean
  dbactionStatus As Integer
End Type
Public Type CatalogueData
    buffer As String * 62
End Type

Public Type CatHeadProps
  ID  As Long
  Parent As Long
  SortTag As String * 50
  Description As String * 150
  ParentDescription As String * 150
  IsNew As Boolean
  IsDirty As Boolean
  IsDeleted As Boolean
End Type
Public Type CatHeadData
    buffer As String * 358
End Type


Sub tesCatalugueProps()
Dim x As CatHeadProps
    MsgBox LenB(x) / 2
End Sub
