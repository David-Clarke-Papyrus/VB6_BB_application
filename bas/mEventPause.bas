Attribute VB_Name = "mEventPause"
Option Explicit

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
        dblDateCounter = ((Year(DATE) + Month(DATE) + Day(DATE)) _
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



