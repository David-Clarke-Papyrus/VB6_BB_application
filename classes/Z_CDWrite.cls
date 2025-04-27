VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Z_CDWrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Dim WithEvents CDWriterPro1 As CDRPROXLib.CDWriterPro
Event Status(pMsg As String, bErr As Boolean)
Event EraseDiscComplete()
Event EraseDiscStart()
Event Complete()

Private Sub CDWriterPro1_EraseDiscComplete()
    RaiseEvent EraseDiscComplete
End Sub
Private Sub CDWriterPro1_EraseDiscStart()
    RaiseEvent EraseDiscStart
End Sub

Private Sub CDWriterPro1_TrackWriteStatus(ByVal lTrackNumber As Long, ByVal lBlocksWritten As Long, ByVal lBlocksToWrite As Long)
    On Error GoTo errHandler
    Dim intPercentTrackWritten As Integer

    'Calc Percent of Current track done
    If lBlocksToWrite = 0 Then Exit Sub
    RaiseEvent Status("Written: " & CStr((lBlocksWritten * 100) \ lBlocksToWrite) & "%", False)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.CDWriterPro1_TrackWriteStatus(lTrackNumber,lBlocksWritten,lBlocksToWrite)", _
         Array(lTrackNumber, lBlocksWritten, lBlocksToWrite)
End Sub


Private Sub CDWriterPro1_ISOItemAdded(ByVal sParentDestPath As String, ByVal sItemDestPath As String, ByVal sItemName As String, ByVal SourceFilePath As String, ByVal bIsDirectory As Boolean, ByVal dteFileDate As Date, ByVal lFileSize As Long)
Dim strErrmsg As String

    On Error GoTo errHandler
    RaiseEvent Status("Adding " & sItemName & " size:" & CStr(lFileSize), False)
    If WriteDisc(strErrmsg) = False Then
        RaiseEvent Status(strErrmsg, True)
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.CDWriterPro1_ISOItemAdded(sParentDestPath,sItemDestPath,sItemName," & _
        "SourceFilePath,bIsDirectory,dteFileDate,lFileSize)", Array(sParentDestPath, sItemDestPath, sItemName, _
         SourceFilePath, bIsDirectory, dteFileDate, lFileSize)
End Sub

Private Sub CDWriterPro1_ClosingSession()
    On Error GoTo errHandler
    RaiseEvent Status("Closing Session...", False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.CDWriterPro1_ClosingSession"
End Sub
Private Sub CDWriterPro1_WriteCancelled()
    On Error GoTo errHandler
    'Inform user of cancelled write
    RaiseEvent Status("Writing Cancelled...", False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.CDWriterPro1_WriteCancelled"
End Sub

Private Sub CDWriterPro1_WriteComplete()
    On Error GoTo errHandler
    RaiseEvent Status("", False)
    RaiseEvent Complete
    'If not in test mode..eject
    If CDWriterPro1.TestWrite = False Then
        'Eject disc
        Call CDWriterPro1.EjectLoad(False)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.CDWriterPro1_WriteComplete"
End Sub

Private Sub CDWriterPro1_WriteError(ByVal WriteError As CDRPROXLibCtl.eWriteErrorType, ByVal DriveError As CDRPROXLibCtl.eCDError, ByVal sErrorInfo As String, ByVal sSenseInfo As String)
    On Error GoTo errHandler
    Dim strError As String
    
    'Get the error type and
    strError = "Writing Error: (" & CStr(WriteError) & ")   " & sErrorInfo & vbCrLf
    
    'If it is a drive error, add the drive error information
    'to the displayed message
    If WriteError = errDriveError Then
        strError = strError & GetDriveErrorMessage(DriveError) & vbCrLf & " Error Sense Data: " & sSenseInfo
        MsgBox "strError = " & strError
    End If
    
    RaiseEvent Status(strError, True)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.CDWriterPro1_WriteError(WriteError,DriveError,sErrorInfo,sSenseInfo)", _
         Array(WriteError, DriveError, sErrorInfo, sSenseInfo)
End Sub
Private Sub CDWriterPro1_ReadingTrackFileError(ByVal TrackFileError As CDRPROXLibCtl.eTrackFileError, ByVal sFileName As String, ByVal lTrackNumber As Long)
    On Error GoTo errHandler
    Dim strErrorMsg As String
    strErrorMsg = GetTrackFileErrorMessage(TrackFileError, sFileName)
    RaiseEvent Status(strErrorMsg, True)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.CDWriterPro1_ReadingTrackFileError(TrackFileError,sFileName,lTrackNumber)", _
         Array(TrackFileError, sFileName, lTrackNumber)
End Sub

Private Sub CDWriterPro1_ReplaceImportedISOFile(ByVal sDestPath As String, ByVal sNewSourcePath As String, ByVal sFileName As String, bReplaceFile As Boolean)
    On Error GoTo errHandler
    bReplaceFile = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.CDWriterPro1_ReplaceImportedISOFile(sDestPath,sNewSourcePath,sFileName," & _
        "bReplaceFile)", Array(sDestPath, sNewSourcePath, sFileName, bReplaceFile)
End Sub

Private Sub CDWriterPro1_PreparingToWrite()
    On Error GoTo errHandler
    RaiseEvent Status("Preparing to Write....", False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.CDWriterPro1_PreparingToWrite"
End Sub


Public Function SetDefaultDrive(pDriveLetter As String) As Boolean
    On Error GoTo errHandler
Dim bFound As Boolean
Dim iDrive As Integer

    Set CDWriterPro1 = New CDRPROXLib.CDWriterPro
    CDWriterPro1.LicenseCode = "Q969Z4D114F142D72224S"
    If CDWriterPro1.SetLogging(oPC.SharedFolderRoot & "\CDError.txt", True) = False Then
        RaiseEvent Status("Error enabling logging", False)
    End If
    If CDWriterPro1.InitDrives(False) = False Then
        MsgBox "Drives Cannot be initialized...Contact support!"
    End If
    'Get the ONLY recordable drives
    bFound = False
    For iDrive = 0 To CDWriterPro1.GetDriveCount() - 1
        'Is recorder - all drives are reported not just writers
        'so we need to save the index so we know which drive to open
        If CDWriterPro1.IsDriveWriter(iDrive) = True Then
            If pDriveLetter > "" Then
                If CDWriterPro1.GetDriveLetter(iDrive) = pDriveLetter Then
                    If CDWriterPro1.OpenDrive(iDrive) = True Then
                        bFound = True
                    End If
                End If
            Else
                If CDWriterPro1.OpenDrive(iDrive) = True Then
                    bFound = True
                    Exit For
                End If
            End If
        End If
    Next
    SetDefaultDrive = bFound
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.SetDefaultDrive(pDriveLetter)", pDriveLetter
End Function
Public Function ISRewritable() As Boolean
    ISRewritable = (CDWriterPro1.GetMediaType = mtCDRW)
End Function
Public Function EraseCD()
    On Error GoTo errHandler
    Call CDWriterPro1.EraseDisc(True)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.EraseCD"
End Function

Public Function AddFile(pPath As String)
    On Error GoTo errHandler
    Call CDWriterPro1.InsertISOItem(GetLastPathItem(pPath), pPath)  'GetLastPathItem(pPath))
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.AddFile(pPath)", pPath
End Function
Private Function GetLastPathItem(strPath As String) As String
    On Error GoTo errHandler
    Dim strTemp As String
    Dim intPos As Integer
    
    'Find first path seperator in reverse
    intPos = InStrRev(strPath, "\")
    strTemp = Right(strPath, Len(strPath) - intPos)
    GetLastPathItem = strTemp
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.GetLastPathItem(strPath)", strPath
End Function

Public Function WriteDisc(pErrMsg As String)
    On Error GoTo errHandler
    WriteDisc = True
    'Check for media loaded
    If CDWriterPro1.GetMediaType() = mtNotLoaded Then
        MsgBox "No CD loaded"
        WriteDisc = False
        Exit Function
    End If

    'Set the properties of the write
    With CDWriterPro1
        .CloseDisc = False 'True '(chkFinalizeDisc.Value = vbChecked) 'Finalize
        
        'Always Close session
        .CloseSession = True
        .SetWriteSpeed (704)
        'Also write Joliet Directory structures in addition to ISO structures
        .VolumeType = vtyISO9660_ONLY ' = vtyISO9660_JOLIET
        .VolumeIdentifier = "Backup" & Format(Date, "yyyymmdd") 'txtVolIdentifier.Text 'Were setting only the Volume Identifier..You could set all the volume descriptors however
                
        'Use this setting if creating an image from network files or when
        'creating an image with a substantial amount of small files
        'Only valid for Data images (ISO/Joliet not Audio discs)
        'Maximum cache is currently limited to 4.2GB by FAT32
        .CacheImage = False ' (chkCacheImage.Value = vbChecked)
        .UseFileDates = True
        'Use Burn Proof/JustLink on this write?
        If (CDWriterPro1.GetDriveCapabilityFlag(SupportsBurnProof) = True) Then
            .SetBufferProtection True
            EventPause 2
       '     MsgBox "Bufferprotect"
        End If
         'wtpDataMode2_XA has been added for backward compatibility
         'wtpDataMode1 is the standard for ISO9660 data disc.
         .WriteType = wtpDataMode1  'IIf((chkUseMode2XA.Value = vbChecked), wtpDataMode2_XA, wtpDataMode1)
         'NOT VALID For DVD formats CD ONLY
         .TestWrite = False ' (chkTestWrite.Value = vbChecked) 'Only write in test mode

    End With
    
    'Start the disc writing process..this should always return True

    'Finally - Write the disc....
    If CDWriterPro1.WriteDisc() = False Then
        pErrMsg = "Wrong media"
        WriteDisc = False
        Exit Function
    End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.WriteDisc(pErrMsg)", pErrMsg
End Function

Public Function GetDriveErrorMessage(ByVal DriveError As CDRPROXLibCtl.eCDError)
    On Error GoTo errHandler
    
    Dim strMsg As String
    
    'Get a String which explains the error code as explained in the help files
    Select Case DriveError
    Case cdNoAdditionalErrorData '1000
        strMsg = "No additional error data was reported"
    Case cdIOTerminated '1001
        strMsg = "Abnormal I/O Termination"
    Case cdLogicalUnitNotReady '1002
        strMsg = "The drive is not ready"
    Case cdLogicalUnitCommFailed '1003
        strMsg = "Communication with drive unit failed"
    Case cdDeviceTrackingError '1004
        strMsg = "The drive could not track properly"
    Case cdWriteGenericError '1005
        strMsg = "Writing error of unknown origin"
    Case cdWriteRecoveryNeeded '1006
        strMsg = "Writing occurred, but recovery is needed"
    Case cdWriteRecoveryFailed '1007
        strMsg = "Recovery attempt failed"
    Case cdWriteLossOfStreaming '1008
        strMsg = "A buffer under-run has occurred"
    Case cdReadUnrecovered '1009
        strMsg = "The disc could not be read"
    Case cdReadRetriesExhausted '1010
        strMsg = "The drive's attempts at reading retries failed"
    Case cdReadErrorTooLong '1011
        strMsg = "The read timed out"
    Case cdReadLECUncorrectable '1012
        strMsg = "While reading, the LEC was not recovered"
    Case cdReadCIRCUnrecovered '1013
        strMsg = "The CIRC could not be validated"
    Case cdReadUPCEANFailed '1014
        strMsg = "Reading of the UPC failed"
    Case cdReadISRCFailed '1015
        strMsg = "Reading of the ISRC failed"
    Case cdReadLossOfStreaming '1016
        strMsg = "Streaming while reading was interrupted"
    Case cdPositioningError '1017
        strMsg = "Drive could not position media"
    Case cdParameterListLengthError '1018
        strMsg = "An incompatible parameter length was sent to the drive"
    Case cdSynchronousTransferError '1019
        strMsg = "A transfer error occurred to the drive"
    Case cdInvalidCommandCode '1020
        strMsg = "An invalid command was sent to the drive"
    Case cdLBAOutOfRange '1021
        strMsg = "Error trying to write past the end of the media"
    Case cdInvalidCDBField '1022
        strMsg = "Invalid command field"
    Case cdInvalidParamterListField '1023
        strMsg = "An incompatible parameter field was sent to the drive"
    Case cdParameterNotSupported '1024
        strMsg = "A command parameter is not supported"
    Case cdParamterValueInvalid '1025
        strMsg = "A command parameter had an invalid value"
    Case cdBusOrDeviceReset '1026
        strMsg = "The SCSI/ATAPI bus was reset and caused a write failure"
    Case cdParametersChanged '1027
        strMsg = "A command parameter changed while in progress"
    Case cdIncompatibleMedium '1028
        strMsg = "The disc is not compatible with the drive mode"
    Case cdReadUnknownMediumFormat '1029
        strMsg = "The drive does not recognize the format of the disc"
    Case cdReadIncompatibleMediumFormat '1030
        strMsg = "The disc format is not compatible with the drive"
    Case cdWriteUnknownMediumFormat '1031
        strMsg = "The disc is of an unknown format"
    Case cdIncompatibleWriteFormat '1032
        strMsg = "The drive cannot write because of an incompatible format"
    Case cdMediaNotPresent '1033
        strMsg = "A disc is not present"
    Case cdLogicalUnitFailure '1034
        strMsg = "The drive had an unknown failure"
    Case cdLogicalUnitTimedOut '1035
        strMsg = "The drive has timed out while completing a command"
    Case cdEraseFailed '1036
        strMsg = "The disc could not be erased"
    Case cdUnableToRecoverTOC '1037
        strMsg = "The Table of Contents is unrecoverable"
    Case cdEndOfUserAreaOnTrack '1038
        strMsg = "Error trying to write past the user area of the media"
    Case cdPacketDoesNotFit '1039
        strMsg = "Packet recording is not configured correctly"
    Case cdIllegalTrackMode '1040
        strMsg = "The current track mode is incompatible with the disc format"
    Case cdInvalidPacketSize '1041
        strMsg = "Packet recording has incorrect size"
    Case cdSessionFixationError '1042
        strMsg = "A generic session closing error occurred"
    Case cdSessionFixationErrorLeadIn '1043
        strMsg = "Error closing Lead-in area"
    Case cdSessionFixationErrorLeadOut '1044
        strMsg = "Error closing Lead-out area"
    Case cdSessionFixationIncompleteTrack '1045
        strMsg = "While closing, the track was never completed"
    Case cdEmptyPartialReservedTrack '1046
        strMsg = "Error attempting to write to a reserved track"
    Case cdPowerCalibrationFull '1047
        strMsg = "Power calibration area is full"
    Case cdPowerCalibrationAreaError '1048
        strMsg = "A flaw exists in the Power calibration area"
    Case cdPMAUpdateFailure '1049
        strMsg = "The disc's PMA could not be updated"
    Case cdPMAFull '1050
        strMsg = "The disc's PMA is full"
    Case cdUnknownError '1051
        strMsg = "Unknown error - use extended data for more information"
    Case cdNoError '1052 - You will never see this most likely
        strMsg = "No Error Reported"
    Case cdNoSeekComplete '1053
        strMsg = "A seek command was interrupted by another command"
    Case cdNTIOError '1054
        strMsg = "A NT disc I/O operation failed"
    Case cdFormatInProgress '1055
        strMsg = "A format is in progress causing operation failure"
    End Select

    'return the error string
    strMsg = strMsg & " (" & CStr(DriveError) & ")."

    GetDriveErrorMessage = strMsg

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.GetDriveErrorMessage(DriveError)", DriveError
End Function
Public Function GetTrackFileErrorMessage(ByVal FileError As CDRPROXLibCtl.eTrackFileError, ByVal FileName As String)
    On Error GoTo errHandler
    Dim strErrorMsg As String

    'This function will handle a track file error
    Select Case FileError
    Case tfeFileOpenError
        strErrorMsg = "Error Opening File: " & FileName
    Case tfeFileCloseError
        strErrorMsg = "Error Closing File: " & FileName
    Case tfeFileReadError
        strErrorMsg = "Error Reading file: " & FileName
    Case Else
        strErrorMsg = "Unknown file error!:" & FileName
    End Select

    'Return error message
    GetTrackFileErrorMessage = strErrorMsg
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.GetTrackFileErrorMessage(FileError,FileName)", Array(FileError, FileName)
End Function

Private Sub Class_Initialize()
    On Error GoTo errHandler
  '  Set CDWriterPro1 = New CDWriterPro
  '  Set CDWriterPro1 = createobject("cdrprox.dll
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set CDWriterPro1 = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_CDWrite.Class_Terminate"
End Sub
