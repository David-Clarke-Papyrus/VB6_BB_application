VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmErrorList 
   Caption         =   "Errors"
   ClientHeight    =   4440
   ClientLeft      =   525
   ClientTop       =   3090
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   10950
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   285
      TabIndex        =   1
      Top             =   3300
      Width           =   1485
   End
   Begin MSComCtlLib.ListView lvwItems 
      Height          =   2985
      Left            =   270
      TabIndex        =   0
      Top             =   300
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   5265
      SortKey         =   6
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Number"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Where"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Where"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "User name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "sortcolumn"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmErrorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private c_ERR As c_Error
Private lngID As Long
Private Sub cmdRefresh_Click()
    Set c_ERR = Nothing
    Set c_ERR = New c_Error
    c_ERR.Load
    FillList c_ERR
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set c_ERR = Nothing
End Sub
Private Sub Form_Load()
    Me.Height = 4845
    Me.Width = 11015
    Set c_ERR = New c_Error
    c_ERR.Load
    FillList c_ERR
End Sub
Public Sub FillList(objComponent As c_Error)
Dim objItem As d_Error
Dim itmList As ListItem
Dim lngIndex As Long

    Set c_ERR = objComponent
    lvwItems.ListItems.Clear
    For lngIndex = 1 To c_ERR.Count
        With objItem
            Set objItem = c_ERR.Item(lngIndex)
            Set itmList = lvwItems.ListItems.Add(Key:=Format$(objItem.ID) & " K")
            With itmList
                .Text = Format(objItem.DateOfError, "dd/mm/yy")
                .SubItems(1) = objItem.Number
                .SubItems(2) = objItem.Description
                .SubItems(3) = objItem.FormName
                .SubItems(4) = objItem.ReportName
                .SubItems(5) = objItem.UserName
                .SubItems(6) = CDbl(CDate(objItem.DateOfError))
            End With
        End With
    Next
End Sub

Private Sub cmdCancel_Click()
    lngID = 0
    Hide
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    ' When a ColumnHeader object is clicked, the ListView control is
    ' sorted by the subitems of that column.
    ' Set the SortKey to the Index of the ColumnHeader - 1

    lvwItems.SortKey = ColumnHeader.Index - 1
    ' Set Sorted to True to sort the list.
    lvwItems.Sorted = True
    If lvwItems.SortOrder = lvwAscending Then
        lvwItems.SortOrder = lvwDescending
    Else
        lvwItems.SortOrder = lvwAscending
    End If
End Sub
