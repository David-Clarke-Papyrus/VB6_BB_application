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
         Picture         =   "frmDashboard.frx":0000
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
            Size            =   14.25
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
         Top             =   3120
         Width           =   8655
      End
      Begin VB.Label lblPO 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "All customer orders are actioned"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
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
            Size            =   14.25
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
            Size            =   14.25
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
         Size            =   14.25
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
63760     FetchData
End Sub

Private Sub cmdSeeCustomerOrders_Click()
63770     Forms(0).BrowseOrders
End Sub

Private Sub cmdSeeCustomerRequests_Click()
63780     Forms(0).BrowseOrderRequests
End Sub

Private Sub cmdSeeDeliveries_Click()
63790     Forms(0).BrowseDELS
End Sub

Private Sub cmdSeePurchaseOrders_Click()
63800     Forms(0).BrowsePOs
End Sub

Private Sub Form_Load()
      'MsgBox "Dashboard 1"
63810     Me.Height = 4905
63820     Me.Width = 11325
      'MsgBox "Dashboard 2"
63830     FetchData
End Sub
Private Sub FetchData()
63840     On Error GoTo errHandler
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
63850     Res = zSQL.GetDashboardData(QtyCO_Unissued, QtyCO_UnfulfilledOrders, QtyCO_UnfulfilledOrderLines, QtyCO_UnfulfilledOrderItems, QtyPO_UnIssued, QtyPO_UnfulfilledOrders, _
                  QtyPO_UnfulfilledOrderLines, QtyPO_UnfulfilledOrderItems, QtyOrderRequests, QtyDEL_Unissued, QtyCOUnActioned, MostRecentCustomerOrderDate, MostRecentPurchaseOrder, MostRecentDelivery)
          
63860     If QtyOrderRequests = 0 Then
63870         Me.lblCustomerRequestStatus.Caption = "All customer requests are actioned"
63880         Me.lblCustomerRequestStatus.ForeColor = &H8000000D
63890     Else
63900         Me.lblCustomerRequestStatus.Caption = "Customer requests to be actioned: " & CStr(QtyOrderRequests)
63910         Me.lblCustomerRequestStatus.ForeColor = vbRed
63920     End If
        ' MsgBox "Dashboard 2b"
63930     If QtyCO_Unissued = 0 Then
63940         Me.lblCO1.Caption = "All customer orders are issued" & "  (" & CStr(QtyCO_UnfulfilledOrders) & " are unfulfilled)"
63950         Me.lblCustomerRequestStatus.ForeColor = &H8000000D
63960     Else
63970         Me.lblCO1.Caption = "Customer orders to be issued: " & CStr(QtyCO_Unissued) & "  (" & CStr(QtyCO_UnfulfilledOrders) & " are unfulfilled)"
63980         Me.lblCO1.ForeColor = vbRed
63990     End If
64000     Me.lblCO2.Caption = "Unactioned titles: " & CStr(QtyCOUnActioned) & vbCrLf & "Most recent issued customer order: " & Format(MostRecentCustomerOrderDate, "dd-mm-yyyy")
          
64010     If QtyPO_UnIssued = 0 Then
64020         Me.lblPO.Caption = "All purchase orders are issued" & "  (" & CStr(QtyPO_UnfulfilledOrders) & " are unfulfilled.)"
64030         Me.lblCustomerRequestStatus.ForeColor = &H8000000D
64040     Else
64050         Me.lblPO.Caption = "Purchase orders to be issued: " & CStr(QtyPO_UnIssued) & "  (" & CStr(QtyPO_UnfulfilledOrders) & " are unfulfilled.)"
64060         Me.lblPO.ForeColor = vbRed
64070     End If
64080     Me.lblPO2.Caption = "Most recent issued purchase order: " & Format(MostRecentPurchaseOrder, "dd-mm-yyyy")
      'MsgBox "Dashboard 2c"
64090     If QtyDEL_Unissued = 0 Then
64100         Me.lblDEL.Caption = "All deliveries are issued"
64110         Me.lblCustomerRequestStatus.ForeColor = &H8000000D
64120     Else
64130         Me.lblDEL.Caption = "Deliveries to be issued: " & CStr(QtyPO_UnIssued)
64140         Me.lblDEL.ForeColor = vbRed
64150     End If
64160     Me.lblDEL2.Caption = "Most recent issued delivery: " & Format(MostRecentDelivery, "dd-mm-yyyy")
          
64170     Exit Sub
errHandler:
      'MsgBox "Dashboard in error fetchdata"

64180     If ErrMustStop Then Debug.Assert False: Resume
64190     ErrorIn "frmDashboard.FetchData"
End Sub
Private Sub Form_Resize()
64200     frleft.Height = NonNegative_Lng(Me.Height - 145)
End Sub
