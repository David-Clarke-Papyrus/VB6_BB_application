VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Host"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type HostProps
  TP_ID As Long
  HostName As String * 50
  SoundexName As String * 20
  HeadAddress As String * 100
  HeadPCode As String * 10
  PostAddress As String * 50
  PostPCode As String * 10
  Country As String * 50
  Phone As String * 20
  Fax As String * 20
  Email As String * 40
  LogoName  As String * 50
  VATRate As Double
End Type


Dim udtProps As HostProps


Private Sub Class_Initialize()
  Load
End Sub

Private Sub Load()
  Dim rs As ADODB.Recordset
  'Dim cmd As New ADODB.Command
  'Dim strSQL As String
  Dim strTemp As String
  Dim i As Long
  
  Set rs = New ADODB.Recordset
  rs.Open "GetHostInfo", oPC.CO
  If Not rs.EOF Then
      With udtProps
         .TP_ID = FNN(rs!Trading_Partner_ID)
         .HostName = FNS(rs!TP_Name)
         .SoundexName = FNS(rs!TP_SoundexName)
         
         If Len(FNS(rs!Add_Line_1)) > 0 Then
            strTemp = Trim(FNS(rs!Add_Line_1))
         End If
         If Len(FNS(rs!Add_Line_2)) > 0 Then
            strTemp = strTemp & vbNewLine & Trim(FNS(rs!Add_Line_2))
         End If
         If Len(FNS(rs!Add_Line_3)) > 0 Then
            strTemp = strTemp & vbNewLine & Trim(FNS(rs!Add_Line_3))
         End If
         If Len(FNS(rs!Add_Line_4)) > 0 Then
            strTemp = strTemp & vbNewLine & Trim(FNS(rs!Add_Line_4))
         End If
         If Len(FNS(rs!Add_Line_5)) > 0 Then
            strTemp = strTemp & vbNewLine & Trim(FNS(rs!Add_Line_5))
         End If
         .HeadAddress = strTemp
         .HeadPCode = FNS(rs!ADD_PCode)
         .Country = FNS(rs!CTR_Name)
         .Phone = FNS(rs!ADD_Tel)
         .Fax = FNS(rs!ADD_Fax)
         .Email = FNS(rs!ADD_Email)
         .LogoName = FNS(rs!CF_LogoFileName)
     '    .VatNumber = FNS(rs!CF_VatNumber)
         .VATRate = FNN(rs!CF_VatRate)
         
         rs.MoveNext
         If Not rs.EOF Then
            strTemp = ""
            If Len(FNS(rs!Add_Line_1)) > 0 Then
               strTemp = Trim(FNS(rs!Add_Line_1))
            End If
            If Len(FNS(rs!Add_Line_2)) > 0 Then
               strTemp = strTemp & vbNewLine & Trim(FNS(rs!Add_Line_2))
            End If
            If Len(FNS(rs!Add_Line_3)) > 0 Then
               strTemp = strTemp & vbNewLine & Trim(FNS(rs!Add_Line_3))
            End If
            If Len(FNS(rs!Add_Line_4)) > 0 Then
               strTemp = strTemp & vbNewLine & Trim(FNS(rs!Add_Line_4))
            End If
            If Len(FNS(rs!Add_Line_5)) > 0 Then
               strTemp = strTemp & vbNewLine & Trim(FNS(rs!Add_Line_5))
            End If
            If Len(FNS(rs!ADD_PCode)) > 0 Then
               strTemp = strTemp & vbNewLine & Trim(FNS(rs!ADD_PCode))
            End If
            .PostAddress = strTemp
            .PostPCode = FNS(rs!ADD_PCode)
         End If
    End With
  Else
    MsgBox "Failed to load Host details!"
  End If
  rs.Close
  Set rs = Nothing
End Sub

Public Property Get HostName() As String
  HostName = Trim$(udtProps.HostName)
End Property

Public Property Get HeadAddress() As String
  HeadAddress = Trim$(udtProps.HeadAddress)
End Property

Public Property Get HeadPostCode() As String
  HeadPostCode = Trim$(udtProps.HeadPCode)
End Property

Public Property Get PostAddress() As String
  PostAddress = Trim$(udtProps.PostAddress)
End Property

Public Property Get PostPCode() As String
  PostPCode = Trim$(udtProps.PostPCode)
End Property

Public Property Get Country() As String
  Country = Trim$(udtProps.Country)
End Property

Public Property Get Phone() As String
  Phone = Trim$(udtProps.Phone)
End Property

Public Property Get Fax() As String
  Fax = Trim$(udtProps.Fax)
End Property

Public Property Get Email() As String
  Email = Trim$(udtProps.Email)
End Property

Public Property Get LogoFileName() As String
  LogoFileName = Trim$(udtProps.LogoName)
End Property
