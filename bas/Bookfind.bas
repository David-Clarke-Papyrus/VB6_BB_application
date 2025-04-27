Attribute VB_Name = "Bookfind"
Option Explicit


  Global NL As String * 2
  Global MSG As String
  Global i As Integer
  Global cStop As String * 1
  Global ReadTags(100) As String * 3
  Global SearchTags(100) As String * 3
  Global strRecord As String * 2048
  Global iread As Long
  Global iSearch As Long
  Global SearchTotal As Long
  Global GetRecNum As Long
  Global EnterFrmInput As String
  Global iRecsFound As Long
  Global Returnval
  Global flgBookFindOK As Boolean
  
'''  Global rec_no_temp As Integer  ' Temporary store used for generating four byte number
'''  Global rec_byte1 As Integer    ' First byte for record number
'''  Global rec_byte2 As Integer    ' Second byte for record number
'''  Global rec_byte3 As Integer    ' Third byte for record number
'''  Global rec_byte4 As Integer    ' Fourth byte for record number
'''
'''
'''' Note Visual Basic is case sensetive when calling a DLL, hence (RUNENGINE)
''''Public Declare Sub RUNENGINE Lib "c:\Bookfind\Endev32.dll" (ByRef bufRequest As restruct, ByRef bufResult As restruct)
''''Declare Sub ShiftLongToBytes Lib "C:\bookfind\BOOKDATA.DLL" Alias "LONGTOBYTES" (ByVal cResBuf As String, ByVal i As Integer, ByVal Num As Long)
'''Public Declare Sub RUNENGINE Lib "Endev32.dll" (ByRef bufRequest As restruct, ByRef bufResult As restruct)
'''Declare Sub ShiftLongToBytes Lib "BOOKDATA.DLL" Alias "LONGTOBYTES" (ByVal cResBuf As String, ByVal i As Integer, ByVal Num As Long)
'''
'''
'''Public bufResult As restruct      ' result buffer is a byte array
'''Public bufRequest As restruct     ' request buffer is a byte array

 Type tProdRec
     code As String * 10
     EAN As String * 13
     MainAuthor As String * 255
     Availability As String * 5
     Title As String * 255
     SubTitle As String * 255
     Description As String * 300
     Edition As String * 20
     PublisherName As String * 50
     UKPrice As String * 10
     USPrice As String * 10
     LocalPrice As String * 10
     SeriesTitle As String * 20
     PublicationDate As String * 50
     BindingCode As String * 5
     MainSupplierName As String * 30
     BFClassification As String * 50
     BICDescription As String * 150
     SACostPrice As String * 15
     Note As String * 40
     Weight As String * 20
     k1 As String * 3
     k6 As String * 30
     MaxResults As Long
  End Type
  Global BFRec As tProdRec

  Type tPubRec
     PublisherName As String * 100
     CodePrefix As String * 10
     ImprintName As String * 100
     strShortname As String * 100
     bfCode As String * 20
  End Type
  Global PubRec As tPubRec
  Type tDistrRec
     DistributorName As String * 100
     DistributorAddress As String * 10
     DistributorTel As String * 20
     DistributorFax As String * 20
     DistributorEMail As String * 40
     DistributorCode As String * 10
  End Type
  Global DistributorRec As tDistrRec

'Function bfPubSearch(fld As String, arg As String) As Long
'On Error GoTo err_bfPubSearch
''
'Dim strSearchSpec As String
'Dim iRecsFound As Long
'
'  ChDrive Left(oPC.BookFindRoot, 1)
'  ChDir oPC.BookFindRoot
'
'    strSearchSpec = "FIND" & Chr$(9) & fld & Chr$(9) & arg
'    If i <> 0 Then
'        MsgBox "Bookfind or in PUBSEARCH"
'    End If
'    bfPubSearch = iRecsFound
'
'EXIT_bfPubSearch:
'    Exit Function
'
'err_bfPubSearch:
'    Select Case Err
''
'        Case Else
'            MsgBox "Bookfind or in PUBSEARCH"
'            Resume EXIT_bfPubSearch
'    End Select
'End Function

'Function bfSearch(pCODE As String, Optional sField As String) As Long
'On Error GoTo err_bfSearch
'
'Dim strSearchSpec As String
'Dim strF As String
'Dim iRecsFound As Long
'  ChDrive Left(oPC.BookFindRoot, 1)
'  ChDir oPC.BookFindRoot
'
'    If Len(sField) > 0 Then
'        strF = sField
'    Else
'        If Len(pCODE) = 13 Then
'            strF = "EA"
'        Else
'            strF = "BN"
'        End If
'    End If
'
'    strSearchSpec = "FIND" & Chr$(9) & strF & Chr$(9) & pCODE
'    If (i <> 0) And (i <> 249) Then
'        MsgBox "Bookfind or in SEARCH"
'
'    End If
'    If i = 249 Then
'        iRecsFound = 0
'        i = 0
'    End If
'    bfSearch = iRecsFound
'
'EXIT_bfSearch:
'    Exit Function
'
'err_bfSearch:
'    Select Case i
'    Case 249
'        iRecsFound = 0
'        Resume Next
'    Case Else
'        MsgBox "Bookfind or in SEARCH"
'        Resume EXIT_bfSearch
'    End Select
'End Function
'
'
'
'
'
'Function Geterr()
'    Geterr = Error
'End Function
'
'Function Geti()
'    Geti = i
'End Function
'
'Sub LoadPubRec()
'On Error GoTo ErrH
'Dim c As Integer
'Dim fMoreTags As Integer
'Dim strTag As String
'Dim strValue As String
'Dim i As Long
'    i = 4
'    c = Asc(Mid(strRecord, i, 1))
'    If c <> 27 Then fMoreTags = True
'    PubRec.PublisherName = ""
'    PubRec.CodePrefix = ""
'    PubRec.ImprintName = ""
'    PubRec.strShortname = ""
'    Do While fMoreTags = True And i <= 1024                 'Handle a tag and text
'        strTag = Mid$(strRecord, i, 1)
'        i = i + 1
'        strTag = strTag + Mid$(strRecord, i, 1)
'        i = i + 2                             'get past the single space
'        c = Asc(Mid$(strRecord, i, 1))
'        strValue = ""
'        Do While c <> 0 And c <> 26
'            strValue = strValue + Mid$(strRecord, i, 1)
'            i = i + 1
'            c = Asc(Mid$(strRecord, i, 1))
'        Loop
'        If c = 26 Then fMoreTags = False
'        i = i + 1
'        Select Case strTag
'        Case "PN"
'            PubRec.PublisherName = strValue
'        Case "IB"
'            PubRec.CodePrefix = strValue
'        Case "IF"
'            PubRec.ImprintName = strValue
'        Case "PU"
'            PubRec.strShortname = strValue
'        End Select
'    Loop
'
'
'    Exit Sub
'
'ErrH:
'    MsgBox Error & "   " & i
'    Exit Sub
'End Sub
'
'Function RemoveBookfindMarkers(pIn As String)
'On Error GoTo err_RemoveBookfindMarkers
'
'Dim i As Integer
'
'    For i = 1 To Len(pIn)
'         If Mid(pIn, i, 1) = "^" Then
'             pIn = Left(pIn, i - 1) & Chr$(10) & Chr$(13) & Right(pIn, Len(pIn) - i - 3)
'             i = i + 2
'         End If
'    Next i
'
'EXIT_RemoveBookfindMarkers:
'    Exit Function
'
'err_RemoveBookfindMarkers:
'    MsgBox "Bookfind or in RemoveBookfindMarkers"
'    Resume EXIT_RemoveBookfindMarkers
'
'End Function


Function ParseXmlDocument(s As String, mudtProps As ProductProps) As Integer
Dim doc As New MSXML2.DOMDocument60
Dim Success As Boolean
Dim nodeList As MSXML2.IXMLDOMNodeList
Dim r As String
Dim r2 As String

   Success = doc.LoadXML(s)
   If Success = False Then
        ParseXmlDocument = 99
   Else
        Set nodeList = doc.SelectNodes("SinglesourceVIColumnDs")
        ParseXmlDocument = 99
        If Not nodeList Is Nothing Then
           Dim Node As MSXML2.IXMLDOMNode
           Dim Node2 As MSXML2.IXMLDOMNode
           Dim Name As String
           Dim Value As String
        
           For Each Node In nodeList
              For Each Node2 In Node.ChildNodes
                    r = UCase(Node2.nodeName)
                    r2 = Node2.text
                    Select Case r
                        Case "VIID"
                            mudtProps.ID = r2
                        Case "EAN"
                            mudtProps.EAN = r2
                            mudtProps.code = ""
                            ParseXmlDocument = 0
                        Case "CODE"
                            mudtProps.code = r2
                            
                        Case "ORIGINATOR"
                            mudtProps.Author = r2
                        Case "NAME1"
                            mudtProps.Title = r2
                        Case "NAME2"
                            mudtProps.SubTitle = r2
                        Case "EDITION"
                            mudtProps.Edition = r2
                        Case "BRANDING"
                            mudtProps.Publisher = r2
                        Case "NOTE"
                            mudtProps.Note = r2
                        Case "SERIESTITLE"
                            mudtProps.SeriesTitle = r2
                        Case "AVAILABILITY"
                            mudtProps.Availability = r2
                        Case "UKPRICE"
                            mudtProps.UKPrice = r2
                        Case "USPRICE"
                            mudtProps.USPrice = r2
                        Case "EUPRICE"
                            mudtProps.EUPrice = r2
                        Case "BASEPRICE"
                            mudtProps.RRP = r2
                            mudtProps.SP = r2
                        Case "LISTPRICEINC"
                            mudtProps.RRP = r2
                            mudtProps.SP = r2
                        Case "PUBLICATIONDATE"
                            mudtProps.PublicationDate = r2
                        Case "PROMOTIONALINFORMATION"
                            mudtProps.Description = r2
                        Case "MAINSUPPLIERNAME"
                            mudtProps.MainSupplierName = r2
                        Case "FORMATCODE"
                            mudtProps.BindingCode = r2
                        Case "BICCODE1"
                            mudtProps.BFClassification = r2
                        Case "WEIGHT"
                            mudtProps.Weight = r2
                        Case "NIELSENSUPPLIERID"
                            mudtProps.BFDistributorCode = r2
        
                    End Select
                    Next Node2
              Next Node
            mudtProps.ProductType = "B"
        End If
   End If
End Function

