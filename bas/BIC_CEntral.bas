Attribute VB_Name = "BIC"
Option Explicit

Sub DropAndCreateBICTable()
    On Error Resume Next
Dim OpenResult As Integer
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COSHORT.Execute "DROP TABLE tBIC"
    oPC.COSHORT.Execute "CREATE TABLE tBIC(BIC_ID INT IDENTITY(1,1) NOT NULL,BIC_CODE VARCHAR(12) NOT NULL,BIC_DESCRIPTION VARCHAR(80) NOT NULL,BIC_LEVEL INT NOT NULL)"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

Sub ImportBIC(pName As String)
Dim strLine As String
Dim Src As z_TextFile
Dim ar() As String
Dim OpenResult As Integer
Dim strSQL As String
Dim i As Integer
Dim strCode As String
Dim strDescription As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set Src = New z_TextFile
    Src.OpenTextFileToRead pName
    Do While Not Src.IsEOF
        strLine = Src.ReadLinefromTextFile
        If Not IsNumeric(left(strLine, 1)) Then
            i = InStr(1, strLine, ",")
            strCode = left(strLine, i - 1)
            strDescription = Mid(strLine, i + 1, 999)

            strSQL = "PBKSC.dbo.tBIC"
            oPC.COSHORT.Execute "INSERT INTO tBIC (BIC_Code,BIC_Description,BIC_Level) VALUES ('" & strCode & "','" & strDescription & "'," & Len(strCode) & ")"

        End If
    Loop
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Src.CloseTextFile
End Sub

