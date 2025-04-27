VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmStoreSelectionForCashupResend 
   Caption         =   "Select branches for cashup resend"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   3645
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   330
      Left            =   510
      TabIndex        =   3
      Top             =   855
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   582
      _Version        =   393216
      Format          =   16449537
      CurrentDate     =   40258
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   465
      Left            =   2370
      TabIndex        =   1
      Top             =   3900
      Width           =   1095
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   2565
      Left            =   45
      OleObjectBlob   =   "frmStoreSelectionForCashupResend.frx":0000
      TabIndex        =   0
      Top             =   1230
      Width           =   3450
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   330
      Left            =   2190
      TabIndex        =   4
      Top             =   855
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   582
      _Version        =   393216
      Format          =   16449537
      CurrentDate     =   40258
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "to"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   1830
      TabIndex        =   6
      Top             =   915
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "From"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   30
      TabIndex        =   5
      Top             =   915
      Width           =   420
   End
   Begin VB.Label lblHeading 
      ForeColor       =   &H8000000D&
      Height          =   480
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   3420
   End
End
Attribute VB_Name = "frmStoreSelectionForCashupResend"
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
    If actionType = "Cashup" Then
        Me.Caption = "Request cashup data"
    Else
        Me.Caption = "Request sales order data"
    End If
End Sub
Private Sub cmdGo_Click()
Dim s As String
Dim i As Integer
    g1.Update
    If actionType = "Cashup" Then
        oSQL.SendInvocation "CASHUP_RESEND", "", CreateXMLList, "", ""
    ElseIf actionType = "COLS" Then
        oSQL.SendInvocation "COLS_SEND", "", CreateXMLList, "", ""
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
                .elText = "SpecsForFetch"
            .elCreateSibling "DATEFROM", True
                .elText = ReverseDate(CStr(dtpFrom))
            .elCreateSibling "DATETO", True
                .elText = ReverseDate(CStr(dtpTo))
            For i = 1 To XA.UpperBound(1)
                If XA(i, 2) = "-1" Then
                    .elCreateSibling "DetailLine", True
                    .chCreate "CODE"
                        .elText = XA(i, 3)
                    .navUP
                End If
            Next i
    End With
    CreateXMLList = xMLDoc.docXML
End Function
