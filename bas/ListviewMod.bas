Attribute VB_Name = "ListviewMod"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2002 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you can not publish
'               or reproduce this code on any web site,
'               on any online service, or distribute on
'               any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const HDS_BUTTONS As Long = &H2
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_GETHEADER As Long = (LVM_FIRST + 31)
Public Const GWL_STYLE As Long = (-16)
Public Const SWP_DRAWFRAME As Long = &H20
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOZORDER As Long = &H4
Public Const SWP_FLAGS As Long = SWP_NOZORDER Or _
                                 SWP_NOSIZE Or _
                                 SWP_NOMOVE Or _
                                 SWP_DRAWFRAME
  
'Public Declare Function SendMessage Lib "user32" _
'    Alias "SendMessageA" _
'   (ByVal hwnd As Long, _
'    ByVal Msg As Long, _
'    ByVal wParam As Long, _
'    lParam As Any) As Long

'--end block--'


