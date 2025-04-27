VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmStoreSelection 
   Caption         =   "Select branches"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   3645
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   465
      Left            =   2400
      TabIndex        =   1
      Top             =   3930
      Width           =   1095
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   3165
      Left            =   60
      OleObjectBlob   =   "frmStoreSelection.frx":0000
      TabIndex        =   0
      Top             =   645
      Width           =   3450
   End
   Begin VB.Label lblHeading 
      ForeColor       =   &H8000000D&
      Height          =   555
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   3420
   End
End
Attribute VB_Name = "frmStoreSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XA As New XArrayDB
Dim rs As New ADODB.Recordset
Dim xMLDoc As New ujXML
Dim oSQL As New z_SQL
Dim actionType As String

Public Sub Component(sHeading As String, pActionType As String)
    Me.lblHeading.Caption = sHeading
    actionType = pActionType
End Sub
Private Sub cmdGo_Click()
Dim s As String
Dim i As Integer
    g1.Update
    
    If actionType = "FromCustBrowse" Then
        Dim oSQL As New z_SQL
        Dim oSM As New z_StockManager
        For i = 1 To XA.UpperBound(1)
            If XA(i, 2) = True Then
                oSQL.SendLoyaltySelectedBranch XA(i, 3)
            End If
        Next i
        oSM.SendCustomerChanges
    Else
    If actionType = "BS" Then
        oSQL.SendInvocation "CustomerStats", "", CreateXMLList, "", ""
    Else
    If actionType = "SOHB" Then
        oSQL.SendInvocation "SOHBulkUpdate", "", CreateXMLList, "", ""
    End If
    End If
    End If
    Me.Hide
End Sub

Private Sub Form_Load()
Dim i As Integer
    Me.Width = 3765
    Me.Height = 4965
    Me.Top = 2000
    Me.Left = 1500
    For i = 1 To oPC.Configuration.Stores.Count
        XA.ReDim 1, i, 1, 3
        XA(i, 1) = oPC.Configuration.Stores.Item(i).Description
        XA(i, 2) = False
        XA(i, 3) = oPC.Configuration.Stores.Item(i).code
    Next
    Set g1.Array = XA
    g1.ReBind
End Sub

Private Sub G1_AfterColUpdate(ByVal ColIndex As Integer)
    XA(g1.Bookmark, ColIndex + 1) = Not (XA(g1.Bookmark, ColIndex + 1))
End Sub


Function CreateXMLList() As String
Dim i As Integer

    Set xMLDoc = New ujXML
    
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "BranchSelection"
            .chCreate "MessageType"
                .elText = "BranchSelection"

            For i = 1 To XA.UpperBound(1)
                    .elCreateSibling "DetailLine", True
                    .chCreate "CODE"
                        .elText = XA(i, 3)
                    .elCreateSibling "SELECTED", True
                        .elText = XA(i, 2)
                    .navUP
            Next i
    End With
    CreateXMLList = xMLDoc.docXML
End Function
