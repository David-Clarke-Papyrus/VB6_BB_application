Attribute VB_Name = "UDTTextList"
Option Explicit

Public Type TextListProps
    Key As String * 30
    Item As String * 255
    F3 As String * 10
    F4 As String * 10
End Type

Public Type TextListData
    buffer As String * 305
End Type

Public Type TextListSProps
    Item As String * 255
End Type

Public Type TextListSData
    buffer As String * 255
End Type

