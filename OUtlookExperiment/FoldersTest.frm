VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1005
      Left            =   1365
      TabIndex        =   0
      Top             =   855
      Width           =   1830
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim ol As Object
Dim olns As Object
Dim oMI As Object
Dim mfol As Object
Dim fol As Outlook.MAPIFolder
Dim res As Boolean
Dim fold As Outlook.Folders
Dim pAttachmentfilename As String
Dim strDestinationEmail As String
Dim strWholeMessage As String
Dim strReference As String
Dim tmp As String
Dim PapyrusDraftsFolder As String
Dim OutlookParentFolder As String


    Set ol = CreateObject("Outlook.Application")
    Set olns = ol.GetNamespace("MAPI")
    OutlookParentFolder = GetIniKeyValue("c:\pbks\PBKSWS.INI", "NETWORK", "OUTLOOKFOLDERMAIN", "")
    PapyrusDraftsFolder = GetIniKeyValue("c:\pbks\PBKSWS.INI", "NETWORK", "OUTLOOKFOLDERSUB", "")
    
    
        Set fol = olns.GetDefaultFolder(olFolderDrafts)
        MsgBox fol.Name
        
        Set fol = fol.Parent
        MsgBox fol.Name
'says Mailbox - Corina van der Spoel
        
        
        Set fol = olns.Folders(OutlookParentFolder)
        MsgBox fol.Name
        
        Set fold = fol.Folders
        MsgBox fold.Count
        
        Set mfol = fold(PapyrusDraftsFolder)
        MsgBox mfol.Name
       
        
        If fol.Parent <> "Mapi" Then
        Set fol = fol.Parent
        MsgBox fol.Name
        End If
        
     '  fold.Add PapyrusDraftsFolder
     '   Set mfol = fold(PapyrusDraftsFolder)

End Sub
