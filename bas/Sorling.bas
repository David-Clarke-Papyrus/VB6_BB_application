Attribute VB_Name = "Sorling"
Option Explicit

Public Function CollectionSort(ByRef oCollection As Collection, pProperty As enSortField, Optional bSortAscending As Boolean = True) As Long
    On Error GoTo errHandler
    Dim lSort1 As Long, lSort2 As Long
    Dim vTempItem1 As Variant, vTempItem2 As Variant, bSwap As Boolean
    
    For lSort1 = 1 To oCollection.Count - 1
        For lSort2 = lSort1 + 1 To oCollection.Count
            If bSortAscending Then
                If pProperty = 1 Then   'Numeric not alpha
                    If CLng(oCollection(lSort1).Properties(pProperty)) > CLng(oCollection(lSort2).Properties(pProperty)) Then
                        bSwap = True
                    Else
                        bSwap = False
                    End If
                Else
                    If UCase(oCollection(lSort1).Properties(pProperty)) > UCase(oCollection(lSort2).Properties(pProperty)) Then
                        bSwap = True
                    Else
                        bSwap = False
                    End If
                End If
            Else
                If pProperty = 1 Then   'Numeric not alpha
                    If CLng(oCollection(lSort1).Properties(pProperty)) < CLng(oCollection(lSort2).Properties(pProperty)) Then
                        bSwap = True
                    Else
                        bSwap = False
                    End If
                Else
                    If UCase(oCollection(lSort1).Properties(pProperty)) < UCase(oCollection(lSort2).Properties(pProperty)) Then
                        bSwap = True
                    Else
                        bSwap = False
                    End If
                End If
            End If
            If bSwap Then
                'Store the items
                If VarType(oCollection(lSort1)) = vbObject Then
                    Set vTempItem1 = oCollection(lSort1)
                Else
                    vTempItem1 = oCollection(lSort1)
                End If
                
                If VarType(oCollection(lSort2)) = vbObject Then
                    Set vTempItem2 = oCollection(lSort2)
                Else
                    vTempItem2 = oCollection(lSort2)
                End If
                
                
                oCollection.Remove lSort2
                oCollection.Add vTempItem2, vTempItem2.Key, lSort1
                
            End If
        Next
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Sorling.CollectionSort(oCollection,pProperty,bSortAscending)", Array(oCollection, _
         pProperty, bSortAscending)
End Function


