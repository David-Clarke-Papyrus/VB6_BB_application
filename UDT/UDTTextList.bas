Attribute VB_Name = "UDTTextList"
Option Explicit

Public Type TextListProps
    Key As String * 30
    Item As String * 255
    f3 As String * 50
    f4 As String * 50
    Active As Boolean
End Type

Public Type TextListData
    buffer As String * 400
End Type

Public Type TextListSProps
    Item As String * 400
End Type

Public Type TextListSData
    buffer As String * 400
End Type

