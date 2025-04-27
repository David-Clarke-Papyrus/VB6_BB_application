VERSION 5.00
Object = "{CCB90040-B81E-11D2-AB74-0040054C3719}#1.0#0"; "OPOSCashDrawer.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Posiflex Cash Drawer OPOS Control Demo"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCD 
      Height          =   360
      Left            =   1575
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   15
      Width           =   1245
   End
   Begin VB.CommandButton Command7 
      Caption         =   "release"
      Height          =   495
      Left            =   5520
      TabIndex        =   12
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "claim"
      Height          =   495
      Left            =   5520
      TabIndex        =   11
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "OpenEpson"
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Frame Frame5 
      Caption         =   "Device Statistics"
      Height          =   2040
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   6600
      Begin VB.TextBox txtRetrieveSt 
         Height          =   765
         Left            =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   600
         Width           =   6375
      End
      Begin VB.CommandButton cmdRetrieveSt 
         Caption         =   "Retrieve Statistics"
         Height          =   435
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "RetrieveStatistics parameter"
         Height          =   255
         Left            =   100
         TabIndex        =   8
         Top             =   300
         Width           =   2775
      End
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   225
      TabIndex        =   4
      Top             =   570
      Width           =   4695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Check Health"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin OposCashDrawer_1_8_LibCtl.OPOSCashDrawer Posiflex_CR1 
      Left            =   5040
      OleObjectBlob   =   "Form1.frx":0000
      Top             =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ob As Object
Dim OposSuccess
Dim OposChExternal

Private Sub Command1_Click()
Set ob = Posiflex_CR1
 
    Cls
    List1.AddItem "ob.Open()= " + Str(ob.Open(txtCD))
'    List1.AddItem "ob.Claim()= " + Str(ob.Claim(1000))
    ob.DeviceEnabled = True
    List1.AddItem "ob.DeviceEnabled= " + Str(ob.DeviceEnabled)
    List1.AddItem "ob.DeviceName= " + ob.DeviceName
    List1.AddItem "CapMultiDrawer: " + Str(ob.CapStatusMultiDrawerDetect)
    List1.AddItem ""
    List1.ListIndex = List1.ListCount - 1
    DoEvents
 
    ' the demonstration device name "CR Demo"
    ' can be changed from the registery by using
    ' Posiflex OPOS Manager
End Sub

Private Sub Command2_Click()
Set ob = Posiflex_CR1
Dim buf As String

    List1.AddItem "CapStatus= " + Str(ob.CapStatus)
    DoEvents
    List1.AddItem "DrawerOpened= " + Str(ob.DrawerOpened)
    DoEvents
    List1.AddItem "OpenDrawer= " + Str(ob.OpenDrawer())
    DoEvents
    List1.AddItem "DrawerOpened= " + Str(ob.DrawerOpened)
    DoEvents
    List1.AddItem "CheckHealth= " + Str(ob.CheckHealth(OposChExternal))
    DoEvents
    List1.AddItem "DrawerOpened= " + Str(ob.DrawerOpened)
    DoEvents
    'List1.AddItem "WaitForClosed= " + Str(ob.WaitForDrawerClose(1000, 1000, 1000, 1000))

' ---------------------------------------------
'    List1.AddItem "Release: " + Str(ob.Release)
'
'    List1.AddItem "ob.DeviceEnabled= " + Str(ob.DeviceEnabled)
'    List1.AddItem "ob.Claimed= " + Str(ob.Claimed)
'    List1.AddItem "ob.DeviceName= " + ob.DeviceName
'    List1.AddItem "OpenDrawer: " + Str(ob.OpenDrawer)
'    List1.AddItem ""
'    List1.ListIndex = List1.ListCount - 1
'    DoEvents
'
'    ob.DeviceEnabled = False
'    List1.AddItem "ob.DeviceEnabled= " + Str(ob.DeviceEnabled)
'    List1.AddItem "ob.Claimed= " + Str(ob.Claimed)
'    List1.AddItem "ob.DeviceName= " + ob.DeviceName
'    List1.AddItem "OpenDrawer: " + Str(ob.OpenDrawer)
'    List1.AddItem ""
'    List1.ListIndex = List1.ListCount - 1
'    DoEvents
'
'    ob.DeviceEnabled = True
'    DoEvents
'    List1.AddItem "ob.DeviceEnabled= " + Str(ob.DeviceEnabled)
'    List1.AddItem "ob.Claimed= " + Str(ob.Claimed)
'    List1.AddItem "ob.DeviceName= " + ob.DeviceName
'    List1.AddItem "OpenDrawer: " + Str(ob.OpenDrawer)
'    List1.AddItem ""
'    List1.ListIndex = List1.ListCount - 1
'    DoEvents
'
'    List1.AddItem "Claim: " + Str(ob.Claim(1000))
'
'    List1.AddItem "ob.DeviceEnabled= " + Str(ob.DeviceEnabled)
'    List1.AddItem "ob.Claimed= " + Str(ob.Claimed)
'    List1.AddItem "ob.DeviceName= " + ob.DeviceName
'    List1.AddItem "OpenDrawer: " + Str(ob.OpenDrawer)
'    List1.AddItem ""
'    List1.ListIndex = List1.ListCount - 1
'    DoEvents
'
'    ob.DeviceEnabled = False
'    List1.AddItem "ob.DeviceEnabled= " + Str(ob.DeviceEnabled)
'    List1.AddItem "ob.Claimed= " + Str(ob.Claimed)
'    List1.AddItem "ob.DeviceName= " + ob.DeviceName
'    List1.AddItem "OpenDrawer: " + Str(ob.OpenDrawer)
'    List1.AddItem ""
'    List1.ListIndex = List1.ListCount - 1
'    DoEvents
'
'    ob.DeviceEnabled = True
'    DoEvents
'    List1.AddItem "ob.DeviceEnabled= " + Str(ob.DeviceEnabled)
'    List1.AddItem "ob.Claimed= " + Str(ob.Claimed)
'    List1.AddItem "ob.DeviceName= " + ob.DeviceName
'    List1.AddItem "OpenDrawer: " + Str(ob.OpenDrawer)
'    List1.ListIndex = List1.ListCount - 1
'    DoEvents

' ---------------------------------------------
'    buf = """""=2"
'    buf = "U_=3"
'    buf = "CommunicationErrorCount=4,DrawerGoodOpenCount=5"
'    List1.AddItem "UpdateStt= " + Str(ob.UpdateStatistics(buf))
'
'    buf = ""
'    buf = "U_"
'    buf = "CommunicationErrorCount,DrawerFailedOpenCount"
'    List1.AddItem "ResetStt= " + Str(ob.ResetStatistics(buf))
'
'    buf = ""
'    buf = "U_"
'    List1.AddItem "RetrieveStt= " + Str(ob.RetrieveStatistics(buf))
'
'    List1.AddItem buf

    List1.ListIndex = List1.ListCount - 1
End Sub

Private Sub Command3_Click()
Set ob = Posiflex_CR1

    List1.AddItem "ob.Close()= " + Str(ob.Close())
    List1.ListIndex = List1.ListCount - 1
End Sub

Private Sub Command4_Click()
    Command3_Click
    DoEvents
    End
End Sub

Private Sub cmdRetrieveSt_Click()

    Dim strParam As String
    Dim lLen As Long
    Dim strErrMsg As String
    Dim strXMLPath As String
    Dim strFindXMLPath As String

    strParam = txtRetrieveSt.Text
    strErrMsg = ""
    strFindXMLPath = ""

    With Posiflex_CR1
        .RetrieveStatistics strParam
        If (.ResultCode <> OPOS_SUCCESS) Then
            strErrMsg = "RetrieveStatistics method error." + vbCrLf + vbCrLf
            strErrMsg = strErrMsg + "ResultCode = " + CStr(.ResultCode) + vbCrLf
            strErrMsg = strErrMsg + "ResultCodeExtended = " + CStr(.ResultCodeExtended)
            MsgBox strErrMsg, vbOKOnly + vbExclamation, "CashDrawer"
            Exit Sub
        End If
    End With
    '*** Retrieve Statistics
    Text1.Text = strParam
    
    strXMLPath = App.Path + "\demo.xml"
    'Delete XML file.
    strFindXMLPath = Dir(strXMLPath)
    If strFindXMLPath <> "" Then
        Kill (strXMLPath)
    End If
    'Create XML file.
    Open strXMLPath For Binary Access Write As #1
        Put #1, , strParam
    Close #1
    
    'Opens another window and indicates the information of the XML file.
    RetrieveStBrowser.Show

End Sub

Private Sub Command5_Click()
Set ob = Posiflex_CR1
 
    Cls
    List1.AddItem "ob.Open()= " + Str(ob.Open("StandardP"))
    'List1.AddItem "ob.Claim()= " + Str(ob.Claim(1000))
    ob.DeviceEnabled = True
    List1.AddItem "ob.DeviceEnabled= " + Str(ob.DeviceEnabled)
    List1.AddItem "ob.DeviceName= " + ob.DeviceName
    List1.AddItem "CapMultiDrawer: " + Str(ob.CapStatusMultiDrawerDetect)
    List1.AddItem ""
    List1.ListIndex = List1.ListCount - 1
End Sub

Private Sub Command6_Click()
    List1.AddItem "claim: " + Str(Posiflex_CR1.ClaimDevice(1000))
End Sub

Private Sub Command7_Click()
    List1.AddItem "release: " + Str(Posiflex_CR1.ReleaseDevice())
    DoEvents
    List1.AddItem "State: " + Str(Posiflex_CR1.State)
End Sub

Private Sub Command8_Click()
    List1.AddItem "State: " + Str(Posiflex_CR1.State)
End Sub

Private Sub Form_Load()
    txtRetrieveSt.Text = "UnifiedPOSVersion,DeviceCategory,ManufacturerName,ModelName,SerialNumber,ManufacturerDate,MechanicalRevision,FirmwareRevision,Interface,InstallationDate,HoursPoweredCount,CommunicationErrorCount,DrawerGoodOpenCount,DrawerFailedOpenCount"
    OposSuccess = 0
    OposChExternal = 2
End Sub

Private Sub Posiflex_CR1_DirectIOEvent( _
    ByVal EventNumber As Long, pData As Long, pString As String)

    List1.AddItem "DIO: " + Str(EventNumber)
    List1.AddItem ""
End Sub

Private Sub Posiflex_CR1_StatusUpdateEvent(ByVal Data As Long)
    List1.AddItem "SUE: " + Str(Data) _
        + ", de=" + Str(Posiflex_CR1.DeviceEnabled) _
        + ", cl=" + Str(Posiflex_CR1.Claimed) _
        + ", do=" + Str(Posiflex_CR1.DrawerOpened)
    DoEvents
    List1.AddItem "wfdc: " + Str(Posiflex_CR1.WaitForDrawerClose(1000, 1000, 1000, 1000))
    List1.AddItem ""
End Sub
