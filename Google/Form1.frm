VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   615
      Left            =   8850
      TabIndex        =   3
      Top             =   2175
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   2670
      Left            =   495
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   1470
      Width           =   8235
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   7530
      TabIndex        =   1
      Top             =   270
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Height          =   765
      Left            =   465
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   255
      Width           =   6705
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   8910
      Top             =   1425
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mXML As ujXML
Dim X As New XArrayDB

Private Sub Command1_Click()
If Inet1.StillExecuting Then Exit Sub
    Inet1.url = Text1
    Screen.MousePointer = vbHourglass
    
    Me.Text2 = Inet1.OpenURL
    SaveSetting "PBKS", "GoogleSearch", "SearchString", Inet1.url
End Sub

Private Sub Command1_LostFocus()
If Inet1.StillExecuting Then Command1.SetFocus
End Sub
Private Function GetISBN(p As String) As String
Dim a() As String
    GetISBN = ""
    a = Split(p, ":")
    If UBound(a) > 0 Then
        If a(0) = "ISBN" Then
            GetISBN = a(1)
        End If
    End If
End Function
Private Sub Command2_Click()
Dim i As Integer
Dim res
Dim vwr As ujXML
Dim f As ujXML
Dim bISBNFOUND As Boolean
Dim T As String

    Set mXML = New ujXML
    mXML.docLoadXML Me.Text2
    res = mXML.navLocate("entry")
    i = 0
    Do
        Set vwr = mXML.docCreateViewer(True)
        i = i + 1
        X.ReDim 1, i, 1, 10
        res = vwr.chLocate("dc:date")
        If res Then
            X(i, 1) = vwr.Element.Text
            vwr.navUP
        Else
            X(i, 1) = ""
        End If
        res = vwr.chLocate("dc:format")
        If res Then
            X(i, 2) = vwr.Element.Text
            vwr.navUP
        Else
            X(i, 2) = ""
        End If
            bISBNFOUND = False
            res = vwr.chLocate("dc:identifier")
            If res Then
                If GetISBN(vwr.Element.Text) > "" Then
                    X(i, 3) = GetISBN(vwr.Element.Text)
                    bISBNFOUND = True
                End If
            Else
                X(i, 3) = ""
            End If
            vwr.navNext
        Do
            If res Then
                If GetISBN(vwr.Element.Text) > "" Then
                    X(i, 3) = GetISBN(vwr.Element.Text)
                    bISBNFOUND = True
                End If
            Else
                X(i, 3) = ""
            End If
            vwr.navNext
        Loop While vwr.Element.nodeName = "dc:identifier" And bISBNFOUND = False
       ' If bISBNFOUND Then vwr.navUP
        vwr.navUP
        
        res = vwr.chLocate("dc:publisher")
        If res Then
            X(i, 4) = vwr.Element.Text
            vwr.navUP
        Else
            X(i, 4) = ""
        End If
        res = vwr.chLocate("dc:title")
        If res Then
            X(i, 5) = vwr.Element.Text
            vwr.navUP
        Else
            X(i, 5) = ""
        End If
        
    Loop While mXML.navNext
    T = ""
    For i = 1 To X.UpperBound(1)
        T = T & CStr(i) & ":" & X(i, 1) & "," & X(i, 2) & "," & X(i, 3) & "," & X(i, 4) & "," & X(i, 5) & vbCrLf
    Next
    MsgBox "Done:" & vbCrLf & T
    Set mXML = Nothing
End Sub

Private Sub Form_Load()
    'SaveSetting "PBKS", "GoogleSearch", "SearchString", Inet1.url
    Me.Text1 = GetSetting("PBKS", "GoogleSearch", "SearchString", "http://books.google.com/books/feeds/volumes?q=football+-soccer")
End Sub

Private Sub Text2_Change()
    Screen.MousePointer = vbDefault

End Sub
