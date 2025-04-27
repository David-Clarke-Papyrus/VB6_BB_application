Attribute VB_Name = "Formats"
Option Explicit
Public Function PBKSPercentF(pIn As Double) As String
    On Error GoTo errHandler
    If CLng(pIn) = pIn Then
        PBKSPercentF = Format(pIn, "#0\%")
    Else
        PBKSPercentF = Format(pIn, "##.#0\%")
    End If

    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then   'overflow
        PBKSPercentF = "overflow"
        Err.Clear
        Exit Function
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Formats.PBKSPercentF(pIn)", pIn
End Function
Public Function PercentF(pIn As Double, Decimals As Integer) As String
    If CLng(pIn) = pIn Then
        PercentF = Format(pIn, "#0\%")
    Else
        If Decimals = 2 Then
            PercentF = Format(pIn, "##.#0\%")
        Else
            If Decimals = 3 Then
                PercentF = Format(pIn, "##.##0\%")
            Else
                If Decimals = 4 Then
                    PercentF = Format(pIn, "##.###0\%")
                End If
            End If
        End If
    End If

End Function

Public Function FormatCode(strCode As String, Optional pForEXport As Boolean) As String
'
'This CODE IS NOT THE SAME AS IN PRODCODE OBJECT
'
Dim iGroupLength As Integer
Dim iPublisherLength As Integer
Dim itest As Long
Dim strGroup As String
Dim strPublisher As String
Dim strRem As String
Dim strChk As String
Dim strSeqNum As String
Dim fForExport As Boolean
    If pForEXport = True Then
        fForExport = True
    Else
        fForExport = False
    End If
    If IsNull(strCode) Then
        FormatCode = ""
        GoTo EXIT_FormatCode
    End If
    If Left(strCode, 1) = "#" Then
        If fForExport = False Then
            FormatCode = strCode
        Else
            FormatCode = ""
        End If
        GoTo EXIT_FormatCode
    End If
    If Len(strCode) <> 10 Then
        FormatCode = ""
        GoTo EXIT_FormatCode
    End If
    'get the group code
    itest = val(Left(strCode, 1))
    If itest >= 0 And itest <= 7 Then
      iGroupLength = 1
    Else
      itest = val(Left(strCode, 2))
      If itest >= 80 And itest <= 94 Then
        iGroupLength = 2
      Else
        itest = val(Left(strCode, 3))
        If itest >= 950 And itest <= 995 Then
          iGroupLength = 3
        Else
          itest = val(Left(strCode, 4))
          If itest >= 9960 And itest <= 9989 Then
            iGroupLength = 4
          Else
            itest = val(Left(strCode, 5))
            If itest >= 99900 And itest <= 99999 Then
              iGroupLength = 5
            Else
              FormatCode = "OR"
              Exit Function
            End If
          End If
        End If
      End If
    End If
    strGroup = Left(strCode, iGroupLength)
    strRem = Right(strCode, 10 - iGroupLength)
    strChk = Right(strRem, 1)
    strRem = Left(strRem, Len(strRem) - 1)
    
    'get the publisher code
    itest = val(Left(strRem, 2))
    If itest >= 0 And itest <= 19 Then
      iPublisherLength = 2
    Else
      itest = val(Left(strRem, 3))
      If itest >= 200 And itest <= 699 Then
        iPublisherLength = 3
      Else
        itest = val(Left(strRem, 4))
        If itest >= 7000 And itest <= 8499 Then
          iPublisherLength = 4
        Else
          itest = val(Left(strRem, 5))
          If itest >= 85000 And itest <= 89999 Then
            iPublisherLength = 5
          Else
            itest = val(Left(strRem, 6))
            If itest >= 900000 And itest <= 949999 Then
              iPublisherLength = 6
            Else
              itest = val(Left(strRem, 7))
              If itest >= 9500000 And itest <= 9999999 Then
                iPublisherLength = 7
              Else
                FormatCode = "OR"
                Exit Function
              End If
            End If
          End If
        End If
      End If
    End If
    strPublisher = Left(strRem, iPublisherLength)
    strSeqNum = Right(strRem, Len(strRem) - iPublisherLength)
    FormatCode = strGroup & "-" & strPublisher & "-" & strSeqNum & "-" & strChk
EXIT_FormatCode:
End Function


