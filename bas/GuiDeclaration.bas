Attribute VB_Name = "GuiDeclaration"
Option Explicit
Public oError As a_Error


'General Variables
Global retval As Long
'Global oAccess As Access.Application
Global Const ERR_APP_NOTRUNNING As Long = 429

Public Enum PreviewPrint
    PrintReport = 1
    PreviewReport = 2
End Enum

Global iPrintPreview As PreviewPrint



