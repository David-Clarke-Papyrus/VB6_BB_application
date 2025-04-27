Attribute VB_Name = "GuiUtility"

Option Explicit

Public Sub LoadListboxColl(lb As ListBox, List As Collection)
    On Error GoTo errHandler
Dim vntItem As Variant

    With lb
        .Clear
        For Each vntItem In List
            .AddItem vntItem(0)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.LoadListboxColl(lb,List)", Array(lb, List)
End Sub
Public Sub LoadCombocol(Combo As ComboBox, List As Collection)
    On Error GoTo errHandler
Dim vntItem As Variant

    With Combo
        .Clear
        For Each vntItem In List
            .AddItem vntItem
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.LoadCombocol(Combo,List)", Array(Combo, List)
End Sub

Public Function ParseMultiSelect(pCtrl As ListBox) As Collection
    On Error GoTo errHandler
Dim Col As Collection
Dim i As Integer
    Set Col = New Collection
   For i = 0 To pCtrl.ListCount - 1
      If pCtrl.Selected(i) Then
         Col.Add pCtrl.List(i)
      End If
   Next i
    Set ParseMultiSelect = Col
    Set Col = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.ParseMultiSelect(pCtrl)", pCtrl
End Function

Public Function ISOpenForm(pType As Form, j As Integer) As Boolean
    On Error GoTo errHandler
Dim bFound As Boolean
Dim i As Integer
    bFound = False
    For i = 0 To Forms.Count
        If Forms(i).Name = pType.Name Then
            j = i
            bFound = True
            Exit For
        End If
    Next i
    ISOpenForm = bFound
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.ISOpenForm(pType,j)", Array(pType, j)
End Function

Public Sub HandleErrorQuiet(pCLose As Boolean)
    pCLose = False
    On Error GoTo errHandler
'    If InException Then
'        oTF.WriteToTextFile ErrDescription
'    '    MsgBox ErrDescription, vbOKOnly, "Exception"
'    Else
        ErrSaveToFile
        pCLose = True
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.HandleError"
End Sub
'//--------------------------------------------------------------------
'// PURPOSE:
'// Pause or delay a procedure for a specified number of seconds
'//
'// ARGUMENTS:
'// Number of seconds. May use fractions in a decimal format (#.##)
'//
'// COMMENTS:
'// Timer() returns a Single value rounded to the nearest 1/100 of a
'// second like a stopwatch. Also, Timer() has a "bug" - it resets
'// itself at midnight. Therefore we need to adjust for this, using
'// some sort of counter. The simplest way is to concatenate the day
'// in front of it with Day(Date) but then the days get reset when the
'// month changes, and of course we need to adjust when the months are
'// reset by the changing year. Fortunately that's as far as we have
'// to go. To avoid an extremely large number by concatenating one in
'// front of the other, we add the different parts of the Date together
'// and then concatenate with the sum.
'//--------------------------------------------------------------------
Public Sub EventPause(sngSeconds As Single)

    '// A Single will convert to scientific notation when concatenating a
    '//  number resulting in 8-digits or more. This can introduce inaccuracies
    '//  as a result of the number being rounded when converted. Therefore we
    '//  must declare doubles when working with the date counter to avoid
    '//  converting to scientific notation.
    Dim dblTotal As Double, dblDateCounter As Double, sngStart As Single
    Dim dblReset As Double, sngTotalSecs As Single, intTemp As Integer
        '// For our purposes, it's better to concatenate five zeros onto the
        '//  end of our date counter, then ADD any Timer values to it.
        dblDateCounter = ((Year(Date) + Month(Date) + Day(Date)) _
          & 0 & 0 & 0 & 0 & 0)
        '// Initialize start time.
        sngStart = Timer
        '// We also need to adjust for the possible resetting of Timer()
        '//  (such as if the Time happens to be just before midnight) when
        '//  adding the Pause time onto the Start time. The folowing formula
        '//  takes ANY value of the total seconds, whether it's above or below
        '//  the 86400 limit, and converts it to a format compatible to the
        '//  date counter.
        sngTotalSecs = (sngStart + sngSeconds)
        intTemp = (sngTotalSecs \ 86400)   '// Return the integer portion only
        dblReset = (intTemp * 100000) + (sngTotalSecs - (intTemp * 86400))
        '// Now we can initialize our total time.
        dblTotal = dblDateCounter + dblReset
    
    '// Timer loop
    Do
        DoEvents        '// Make sure any other tasks get some attention
    '// For this to work properly, we cannot create a variable with the
    '//  concatenated expression and plug it in unless we reset the variable
    '//  during the loop. Much better to do it like this:
    Loop While (dblDateCounter + Timer) < dblTotal
    
End Sub

