Attribute VB_Name = "RegSettings"
Option Explicit

Public Const RS1_FAILED = 2777 'Error for failing to change regional settings

Public Sub CheckRegionalSettings()
Dim RegS As z_RegionalSettings

    On Error GoTo ERR_Handler
    
    Set RegS = New z_RegionalSettings
    
    With RegS
'        If .CurrencySymbol <> "R" Then
'            MyMsg "In the Control Panel regional options" & vbLf & _
'                   "the symbol for currency is set to '" & .CurrencySymbol & "'" & vbLf & _
'                   "This could cause problems while runing this propgram!" & vbLf & _
'                   "Therefore the currency symbol will be set to 'R'.", vbOKOnly + vbExclamation, "Changing regional setting!"
'            .SetCurrencySymbol = "R"
'        End If
        If .ShortDate <> "dd/MM/yyyy" Or .DateSepSymbol <> "/" Then
            MsgBox "In the Control Panel regional options" & vbLf & _
                        "the setting for Short Date is '" & .ShortDate & "'" & vbLf & _
                        "This could cause problems while runing this propgram!" & vbLf & _
                        "Therefore the Short Date will be set to 'dd/MM/yyyy'.", vbOKOnly + vbExclamation, "Changing regional setting!"
            If .DateSepSymbol <> "/" Then .SetDateSepSymbol = "/"
            .SetShortDate = "dd/MM/yyyy"
        End If
        If .TimeSepSymbol <> ":" Then .SetTimeSepSymbol = ":"
        If .TimeMask <> "HH:mm:ss" Then .SetTimeMask = "HH:mm:ss"
    End With
    
EXIT_Handler:
    Set RegS = Nothing
    Exit Sub
ERR_Handler:
    If Err = RS1_FAILED Then
        MsgBox "Failed to change settings in regonal options!" & vbLf & _
                    "Please change the following settings by manually:" & vbLf & _
                    "Open Start/Settings/Control Panel/Regional Options" & vbLf & _
                    "Select CURRENCY tab and set CURRENCY SYMBOL to 'R'" & vbLf & _
                    "Select TIME tab and set TIME FORMAT to 'hh:mm:ss'" & vbLf & _
                    "Select DATE tab and set SHORT DATE FORMAT to 'dd/MM/yyyy' and DATE SEPARATOR to '/'", _
                    vbOKOnly + vbCritical, "Failed to change Regional Options!"
    End If
    GoTo EXIT_Handler
End Sub

