Attribute VB_Name = "UpdateSupplier"
Option Explicit
Dim oBatch As z_Batch
Dim lngRecordsReturned As Long

Function UpdateLastSupplierUsed(pPublisher As String, pLastUsedSupplier As Long, pNewSupplier As Long)

On Error GoTo ERR_Handler

Dim rs As adodb.Recordset
    Set oBatch = New z_Batch
    retval = oBatch.DropTable("TEMP_ListOfPublishers", "Erasing table . . . ")
    lngRecordsReturned = oBatch.RunProc("q_ListOfPublishers", Array(), "Processing report . . .")
    Set rs = New adodb.Recordset
    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_ListOfPublishers", Array(), " Writing report . . .", rs)
    
EXIT_Handler:
    Exit Function

ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler

End Function
