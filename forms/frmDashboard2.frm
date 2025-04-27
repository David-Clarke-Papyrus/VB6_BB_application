VERSION 5.00
Begin VB.Form frmDashboard 
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   11205
   Begin VB.PictureBox frleft 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      Height          =   4575
      Left            =   0
      ScaleHeight     =   4515
      ScaleWidth      =   11145
      TabIndex        =   0
      Top             =   0
      Width           =   11205
      Begin VB.CommandButton cmdRefresh 
         Height          =   405
         Left            =   75
         Picture         =   "frmDashboard2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton cmdSeeDeliveries 
         BackColor       =   &H00D9C6CB&
         Caption         =   "See Deliveries"
         Height          =   375
         Left            =   8910
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3105
         Width           =   1920
      End
      Begin VB.CommandButton cmdSeePurchaseOrders 
         BackColor       =   &H00D9C6CB&
         Caption         =   "See Purchase Orders"
         Height          =   375
         Left            =   8925
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2115
         Width           =   1920
      End
      Begin VB.CommandButton cmdSeeCustomerOrders 
         BackColor       =   &H00D9C6CB&
         Caption         =   "See Customer Orders"
         Height          =   375
         Left            =   8925
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1020
         Width           =   1920
      End
      Begin VB.CommandButton cmdSeeCustomerRequests 
         BackColor       =   &H00D9C6CB&
         Caption         =   "See requests"
         Height          =   375
         Left            =   8925
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   495
         Width           =   1920
      End
      Begin VB.Label lblDEL2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "All customer orders are actioned"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   315
         TabIndex        =   13
         Top             =   3465
         Width           =   8625
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPO2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "All customer orders are actioned"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   330
         TabIndex        =   11
         Top             =   2505
         Width           =   8625
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCO2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "All customer orders are actioned"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   330
         TabIndex        =   10
         Top             =   1410
         Width           =   8655
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDEL 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "All customer orders are actioned"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   330
         TabIndex        =   6
         Top             =   3105
         Width           =   8655
      End
      Begin VB.Label lblPO 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "All customer orders are actioned"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   345
         TabIndex        =   5
         Top             =   2145
         Width           =   8655
      End
      Begin VB.Label lblCO1 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "All customer orders are actioned"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   345
         TabIndex        =   4
         Top             =   1050
         Width           =   8655
      End
      Begin VB.Label lblCustomerRequestStatus 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "All customer requests are actioned"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   345
         TabIndex        =   3
         Top             =   525
         Width           =   8655
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "All customer orders are actioned"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRefresh_Click()
    FetchData
End Sub

Private Sub cmdSeeCustomerOrders_Click()
    Forms(0).BrowseOrders
End Sub

Private Sub cmdSeeCustomerRequests_Click()
    Forms(0).BrowseOrderRequests
End Sub

Private Sub cmdSeeDeliveries_Click()
    Forms(0).BrowseDELS
End Sub

Private Sub cmdSeePurchaseOrders_Click()
    Forms(0).BrowsePOs
End Sub

Private Sub Form_Load()
'MsgBox "Dashboard 1"
    Me.Height = 4905
    Me.Width = 11325
'MsgBox "Dashboard 2"
    FetchData
End Sub
Private Sub FetchData()
    On Error GoTo errHandler
Dim QtyCO_Unissued As Long
Dim QtyCO_UnfulfilledOrders As Long
Dim QtyCO_UnfulfilledOrderLines As Long
Dim QtyCO_UnfulfilledOrderItems As Long
Dim QtyPO_UnIssued As Long
Dim QtyPO_UnfulfilledOrders As Long
Dim QtyPO_UnfulfilledOrderLines As Long
Dim QtyPO_UnfulfilledOrderItems As Long
Dim QtyOrderRequests As Long
Dim QtyDEL_Unissued  As Long
Dim QtyCOUnActioned As Long
Dim MostRecentCustomerOrderDate As Date
Dim MostRecentPurchaseOrder As Date
Dim MostRecentDelivery As Date

    Dim zSQL As New z_SQL
    Dim Res As Long
'MsgBox "Dashboard 2a"
    Res = zSQL.GetDashboardData(QtyCO_Unissued, QtyCO_UnfulfilledOrders, QtyCO_UnfulfilledOrderLines, QtyCO_UnfulfilledOrderItems, QtyPO_UnIssued, QtyPO_UnfulfilledOrders, _
            QtyPO_UnfulfilledOrderLines, QtyPO_UnfulfilledOrderItems, QtyOrderRequests, QtyDEL_Unissued, QtyCOUnActioned, MostRecentCustomerOrderDate, MostRecentPurchaseOrder, MostRecentDelivery)
    
    If QtyOrderRequests = 0 Then
        Me.lblCustomerRequestStatus.Caption = "All customer requests are actioned"
        Me.lblCustomerRequestStatus.ForeColor = &H8000000D
    Else
        Me.lblCustomerRequestStatus.Caption = "Customer requests to be actioned: " & CStr(QtyOrderRequests)
        Me.lblCustomerRequestStatus.ForeColor = vbRed
    End If
  ' MsgBox "Dashboard 2b"
    If QtyCO_Unissued = 0 Then
        Me.lblCO1.Caption = "All customer orders are issued" & "  (" & CStr(QtyCO_UnfulfilledOrders) & " are unfulfilled)"
        Me.lblCustomerRequestStatus.ForeColor = &H8000000D
    Else
        Me.lblCO1.Caption = "Customer orders to be issued: " & CStr(QtyCO_Unissued) & "  (" & CStr(QtyCO_UnfulfilledOrders) & " are unfulfilled)"
        Me.lblCO1.ForeColor = vbRed
    End If
    Me.lblCO2.Caption = "Unactioned titles: " & CStr(QtyCOUnActioned) & vbCrLf & "Most recent issued customer order: " & Format(MostRecentCustomerOrderDate, "dd-mm-yyyy")
    
    If QtyPO_UnIssued = 0 Then
        Me.lblPO.Caption = "All purchase orders are issued" & "  (" & CStr(QtyPO_UnfulfilledOrders) & " are unfulfilled.)"
        Me.lblCustomerRequestStatus.ForeColor = &H8000000D
    Else
        Me.lblPO.Caption = "Purchase orders to be issued: " & CStr(QtyPO_UnIssued) & "  (" & CStr(QtyPO_UnfulfilledOrders) & " are unfulfilled.)"
        Me.lblPO.ForeColor = vbRed
    End If
    Me.lblPO2.Caption = "Most recent issued purchase order: " & Format(MostRecentPurchaseOrder, "dd-mm-yyyy")
'MsgBox "Dashboard 2c"
    If QtyDEL_Unissued = 0 Then
        Me.lblDEL.Caption = "All deliveries are issued"
        Me.lblCustomerRequestStatus.ForeColor = &H8000000D
    Else
        Me.lblDEL.Caption = "Deliveries to be issued: " & CStr(QtyPO_UnIssued)
        Me.lblDEL.ForeColor = vbRed
    End If
    Me.lblDEL2.Caption = "Most recent issued delivery: " & Format(MostRecentDelivery, "dd-mm-yyyy")
    
    Exit Sub
errHandler:
'MsgBox "Dashboard in error fetchdata"

    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDashboard.FetchData"
End Sub
Private Sub Form_Resize()
    frleft.Height = NonNegative_Lng(Me.Height - 145)
End Sub
