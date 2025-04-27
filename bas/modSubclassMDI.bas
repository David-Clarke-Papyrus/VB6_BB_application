Attribute VB_Name = "modSubclassMDI"
' ********************************************************************************
'
'  Project: MDIClient texturization
'  Author:  G. D. Sever    aka "The Hand" (thehand@elitevb.com)
'
' ********************************************************************************
'
'  Description:
'
'    This module demonstrates how to add a background texture and a logo to the
'    MDIClient area of an MDIForm.
'
'  Terms of use:
'
'    You are free to use this code however you wish in your compiled projects.
'    If pieces of the source code are published either on their own or in part,
'    give credit where it is due, and we'll all get along just fine.
'
' ********************************************************************************
'      Visit http://www.elitevb.com for more high-powered solutions!!
' ********************************************************************************
Option Explicit

'Used in our safe subclassing hack
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
'Used to store the original procedure addresses by individual window handle
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'Messages we look for - these will all require us to redraw the MDIClient area
Private Const WM_SIZE       As Long = &H5
Private Const WM_PAINT      As Long = &HF
Private Const WM_ERASEBKGND As Long = &H14
'Used to redirect our default window process for the MDI client area
Private Const GWL_WNDPROC   As Long = (-4)
'Used to redirect the messages and invoke the default process for the MDIClient area
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'We use this to get a handle on the MDIClient area
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long


Public Sub subclassMDIClientArea(aMDIForm As MDIForm)
On Error GoTo Errh

    Dim hMDIClient  As Long         ' Handle for our MDIClient area
    Dim oldProc     As Long         ' Original process address for the MDIClient area
    
    ' Find the handle for our MDIForm's MDIClient area
    hMDIClient = FindWindowEx(aMDIForm.hWnd, ByVal 0&, "MDIClient", vbNullString)
    ' Redirect all messages for the MDIClient area to our own procedure
    oldProc = SetWindowLong(hMDIClient, GWL_WNDPROC, AddressOf MainMDIClientProc)
    ' Save a few variables to use later... most importantly the default
    '  process address. We also get a pointer to our MDIForm so we can use
    '  our "safe" subclassing method.
    SetProp hMDIClient, "MAINOldProc", oldProc
    SetProp hMDIClient, "MAINPtr", ObjPtr(aMDIForm)
    ' While we're at it, we'll store the MDIClient's hwnd against the MDIForm's hwnd.
    '  We'll need this when the user selects/deselects stuff so we can redraw it...
    '  and rather than using FindWindowEx again, we'll just use this :-D
    SetProp aMDIForm.hWnd, "MAINhMDIClient", hMDIClient
Exit Sub
Errh:
    MsgBox Error
End Sub
Public Sub unsubclassMDIClientArea(aMDIForm As MDIForm)
    ' ***************************************
    '  Time to turn all the neato stuff off.
    ' ***************************************
    
    Dim hMDIClient  As Long         ' Handle for our MDIClient area
    Dim oldProc     As Long         ' Original process address for the MDIClient area
    
    ' Retrieve the handle and original process for our MDIClient area for the
    '  specified MDI form
    hMDIClient = GetProp(aMDIForm.hWnd, "MAINhMDIClient")
    oldProc = GetProp(hMDIClient, "MAINOldProc")
    ' Start sending the messages to the original procedure address
    SetWindowLong hMDIClient, GWL_WNDPROC, oldProc
    ' Clean up all temporary stuff we stored against our handles
    RemoveProp hMDIClient, "MAINOldProc"
    RemoveProp hMDIClient, "MAINPtr"
    RemoveProp aMDIForm.hWnd, "MAINhMDIClient"

End Sub
Private Function MainMDIClientProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim oldProc As Long     ' Original process address for the MDIClient area
    Dim aPtr    As Long     ' Pointer to the MDIForm which has this MDIClient area
    Dim aFrm    As frmMain  ' Main MDI Form. We're using this as a dummy variable.
                            '  we'll copy the pointer into this variable so we can
                            '  invoke the "DrawLogo" method, then immediately remove
                            '  the pointer address to prevent subclassing crashes.
    
    ' Get the original process address from the interal Windows database
    oldProc = GetProp(hWnd, "MAINOldProc")
    
    If uMsg = WM_PAINT Or uMsg = WM_SIZE Then
        ' Invoke the default process. Trust me, you don't want to forget to do this.
        MainMDIClientProc = CallWindowProc(oldProc, hWnd, uMsg, wParam, lParam)
        ' Retrieve the MDIForm's pointer out of the interal Windows database
        aPtr = GetProp(hWnd, "MAINPtr")
        '  Slimy, slimy hack! Invokes the DrawLogo method on the MDIForm to which
        '   this MDIClient area belongs.
        If aPtr > 0 Then
            CopyMemory aFrm, aPtr, 4
            aFrm.DrawLogo hWnd
            CopyMemory aFrm, 0&, 4
        End If
    Else
        ' Invoke the default process. Trust me, you don't want to forget to do this.
        MainMDIClientProc = CallWindowProc(oldProc, hWnd, uMsg, wParam, lParam)
    End If

End Function


