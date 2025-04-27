Attribute VB_Name = "Services"
Option Explicit

Sub SetupServicePOS()
Dim k As clsKey
Dim X As String, Y As String
Dim fs As New Scripting.FileSystemObject
Dim lngResult As Long
Dim strCommand As String
'DisplayName= 'POS_Server'

    strCommand = "SC DELETE POS_Server"
    F_7_AB_1_ShellAndWaitSimple strCommand, vbHide
    
    strCommand = "SC CREATE POS_Server DisplayName= POS_Server type= own start= auto binpath= C:\\PBKS\\Services\\SRVANY.EXE depend= MSSQL$PBKSINSTANCE2/MSMQ"
    F_7_AB_1_ShellAndWaitSimple strCommand, vbHide

    Set k = New clsKey
    k.hKey = HKEY_LOCAL_MACHINE
    k.Path = "SYSTEM\CurrentControlSet\Services\POS_Server\Parameters"
    If Not k.Exists Then
        k.Path = "SYSTEM\CurrentControlSet\Services\POS_Server"
        k.SubKeys.Add "Parameters"
        k.Path = "SYSTEM\CurrentControlSet\Services\POS_Server\Parameters"
        Call k.Values.Add("Application", "C:\PBKS\Executables\PBKS_POSSvr.exe", 1)
        Call k.Values.Add("AppDirectory", "C:\PBKS\Executables", 1)
        k.Path = "SYSTEM\CurrentControlSet\Services\POS_Server"
        k.Values.Delete ("Type")
        k.Values.Add "Type", 272, REG_DWORD
    End If

End Sub
Sub SetupServiceDispatch()
Dim k As clsKey
Dim X As String, Y As String
Dim fs As New Scripting.FileSystemObject
Dim lngResult As Long
Dim strCommand As String
    
    strCommand = "SC DELETE PapyrusDispatcher"
    F_7_AB_1_ShellAndWaitSimple strCommand, vbHide
    strCommand = "SC CREATE PapyrusDispatcher DisplayName= PapyrusDispatcher type= own start= auto binpath= C:\\PBKS\\Services\\SRVANY.EXE depend= MSSQL$PBKSINSTANCE2"
    F_7_AB_1_ShellAndWaitSimple strCommand, vbHide
    Set k = New clsKey
    k.hKey = HKEY_LOCAL_MACHINE
    k.Path = "SYSTEM\CurrentControlSet\Services\PapyrusDispatcher\Parameters"
    If Not k.Exists Then
        k.Path = "SYSTEM\CurrentControlSet\Services\PapyrusDispatcher"
        k.SubKeys.Add "Parameters"
        k.Path = "SYSTEM\CurrentControlSet\Services\PapyrusDispatcher\Parameters"
        Call k.Values.Add("Application", "C:\PBKS\Executables\PBKS_Dispatch.exe", 1)
        Call k.Values.Add("AppDirectory", "C:\PBKS\Executables", 1)
        k.Path = "SYSTEM\CurrentControlSet\Services\PapyrusDispatcher"
        k.Values.Delete ("Type")
        k.Values.Add "Type", 272, REG_DWORD
    End If

End Sub

