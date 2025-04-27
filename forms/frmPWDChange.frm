VERSION 5.00
Begin VB.Form frmPWDChange 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Change your password"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5580
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtShortname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   3825
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   105
      Width           =   1530
   End
   Begin VB.TextBox txtNew2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3825
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1695
      Width           =   1530
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3330
      Picture         =   "frmPWDChange.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2595
      Width           =   1000
   End
   Begin VB.TextBox txtNew1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   3825
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1230
      Width           =   1530
   End
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4350
      Picture         =   "frmPWDChange.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2595
      Width           =   1000
   End
   Begin VB.TextBox txtOld 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   3825
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   1530
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      Caption         =   "Passwords are not case sensitive."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1350
      TabIndex        =   11
      Top             =   2100
      Width           =   4035
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   1050
      TabIndex        =   10
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      Caption         =   "1. Short name (3 characters max.)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   -45
      TabIndex        =   9
      Top             =   165
      Width           =   3630
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      Caption         =   "4. New password again"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1035
      TabIndex        =   8
      Top             =   1710
      Width           =   2580
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      Caption         =   "3. New password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1020
      TabIndex        =   7
      Top             =   1245
      Width           =   2580
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      Caption         =   "2. Existing password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1035
      TabIndex        =   6
      Top             =   555
      Width           =   2580
   End
End
Attribute VB_Name = "frmPWDChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCancel As Boolean

Public Property Get IsCancelled() As Boolean
    On Error GoTo errHandler
    IsCancelled = bCancel
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPWDChange.IsCancelled"
End Property

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    bCancel = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPWDChange.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSelect_Click()
    On Error GoTo errHandler
Dim strfullname As String
Dim i As Integer

    If Len(txtNew1) < 2 Then
        MsgBox "Your password must be at least 3 alpha-numeric characters long.", vbInformation, "Warning"
        Exit Sub
    End If

    For i = 1 To Len(txtNew1)
        If Asc(Mid(txtNew1, i, 1)) < 48 Or Asc(Mid(txtNew1, i, 1)) > 122 Then
            MsgBox "Your password must use only 0-9 and A-Z.", vbInformation, "Warning"
            txtNew1 = ""
            Exit Sub
        End If
    Next i
    
    
    If oPC.Configuration.ChangePassword(txtShortname, txtOld, txtNew1, strfullname) Then
        MsgBox "Password for " & strfullname & " has been changed!", vbInformation, "Success"
        bCancel = False
        Me.Hide
    Else
        MsgBox "Your password has not been changed! " & vbCrLf & "Either box 1 or box 2 is invalid,or you have typed your present password incorrectly.", vbInformation, "Success"
    End If
    
    If UCase(txtNew1) <> UCase(txtNew2) Then
        MsgBox "Your new password is typed differently in boxes 3 and 4. Please reenter.", vbInformation, "Can't change password'"
        txtNew1 = ""
        txtNew2 = ""
        Exit Sub
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPWDChange.cmdSelect_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Label4_Click()
    On Error GoTo errHandler
Dim str As String
    str = "In box one, enter your shortname." & vbCrLf _
            & "If you normally enter IANCAR when you sign a document" & vbCrLf _
            & "or access a secure part of the system then" & vbCrLf _
            & "enter IAN in box 1 and CAR in box 2" & vbCrLf _
            & "In box 3 enter your new password," & vbCrLf _
            & "In box 4 re-enter the same new password." & vbCrLf
    MsgBox str, vbInformation, "Help"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPWDChange.Label4_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub txtNew1_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim i As Integer
'    For i = 1 To Len(txtNew1)
'        If Asc(Mid(txtNew1, i, 1)) < 33 Or Asc(Mid(txtNew1, i, 1)) > 120 Then
'            MsgBox "Your password must use only 0-9 and A-Z.", vbInformation, "Warning"
'            txtNew1 = ""
'            Cancel = True
'            Exit For
'        End If
'    Next i
'    If Len(txtNew1) < 3 Then
'        MsgBox "Your password must be at least 3 alpha-numeric characters long.", vbInformation, "Warning"
'        Cancel = True
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPWDChange.txtNew1_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtOld_Change()
    On Error GoTo errHandler
    oPC.CurrentSecurityCode = txtOld
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPWDChange.txtOld_Change", , EA_NORERAISE
    HandleError
End Sub
