Attribute VB_Name = "BICImport"
'****************************************************************
'Microsoft SQL Server 2000
'Visual Basic file generated for DTS Package
'File Name: C:\PapyRusCodeNew\BAS\ImportBIC.bas
'Package Name: ImportBIC
'Package Description: DTS package description
'Generated Date: 2002/09/20
'Generated Time: 08:43:25
'****************************************************************

Option Explicit
Public goPackageOld As New DTS.Package
Public goPackage As DTS.Package2
Private strSourceFile As String
Public Sub IMPORTBIC(pSource As String)
        strSourceFile = pSource
        Set goPackage = goPackageOld
        
        goPackage.Name = "ImportBIC"
        goPackage.Description = "DTS package description"
        goPackage.WriteCompletionStatusToNTEventLog = False
        goPackage.FailOnError = False
        goPackage.PackagePriorityClass = 2
        goPackage.MaxConcurrentSteps = 4
        goPackage.LineageOptions = 0
        goPackage.UseTransaction = True
        goPackage.TransactionIsolationLevel = 4096
        goPackage.AutoCommitTransaction = True
        goPackage.RepositoryMetadataOptions = 0
        goPackage.UseOLEDBServiceComponents = True
        goPackage.LogToSQLServer = False
        goPackage.LogServerFlags = 0
        goPackage.FailPackageOnLogFailure = False
        goPackage.ExplicitGlobalVariables = False
        goPackage.PackageType = 0
        


'---------------------------------------------------------------------------
' create package connection information
'---------------------------------------------------------------------------

Dim oConnection As DTS.Connection2

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("DTSFlatFile")

        oConnection.ConnectionProperties("Data Source") = strSourceFile
        oConnection.ConnectionProperties("Mode") = 1
        oConnection.ConnectionProperties("Row Delimiter") = vbCrLf
        oConnection.ConnectionProperties("File Format") = 1
        oConnection.ConnectionProperties("Column Delimiter") = "="
        oConnection.ConnectionProperties("File Type") = 1
        oConnection.ConnectionProperties("Skip Rows") = 0
        oConnection.ConnectionProperties("Text Qualifier") = """"
        oConnection.ConnectionProperties("First Row Column Name") = False
        oConnection.ConnectionProperties("Max characters per delimited column") = 8000
        
        oConnection.Name = "Connection 1"
        oConnection.ID = 1
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = "E:\CLASSDOC\BIC.RT"
        oConnection.ConnectionTimeout = 60
        oConnection.UseTrustedConnection = False
        oConnection.UseDSL = False
        
        'If you have a password for this connection, please uncomment and add your password below.
        'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("SQLOLEDB")

        oConnection.ConnectionProperties("Persist Security Info") = True
        oConnection.ConnectionProperties("User ID") = "sa"
        oConnection.ConnectionProperties("Initial Catalog") = "PJ"
        oConnection.ConnectionProperties("Data Source") = "DAVIDCLARKE"
        oConnection.ConnectionProperties("Application Name") = "DTS  Import/Export Wizard"
        
        oConnection.Name = "Connection 2"
        oConnection.ID = 2
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = "DAVIDCLARKE"
        oConnection.UserID = "sa"
        oConnection.ConnectionTimeout = 60
        oConnection.Catalog = "PJ"
        oConnection.UseTrustedConnection = False
        oConnection.UseDSL = False
        
        'If you have a password for this connection, please uncomment and add your password below.
        'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'---------------------------------------------------------------------------
' create package steps information
'---------------------------------------------------------------------------

Dim oStep As DTS.Step2
Dim oPrecConstraint As DTS.PrecedenceConstraint

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "Copy Data from BIC to [PJ].[dbo].[tBIC] Step"
        oStep.Description = "Copy Data from BIC to [PJ].[dbo].[tBIC] Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Copy Data from BIC to [PJ].[dbo].[tBIC] Task"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = False
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'---------------------------------------------------------------------------
' create package tasks information
'---------------------------------------------------------------------------

'------------- call Task_Sub1 for task Copy Data from BIC to [PJ].[dbo].[tBIC] Task (Copy Data from BIC to [PJ].[dbo].[tBIC] Task)
Call Task_Sub1(goPackage)

'---------------------------------------------------------------------------
' Save or execute package
'---------------------------------------------------------------------------

'goPackage.SaveToSQLServer "(local)", "sa", ""
goPackage.Execute
tracePackageError goPackage
goPackage.UnInitialize
'to save a package instead of executing it, comment out the executing package lines above and uncomment the saving package line
Set goPackage = Nothing

Set goPackageOld = Nothing

End Sub


'-----------------------------------------------------------------------------
' error reporting using step.GetExecutionErrorInfo after execution
'-----------------------------------------------------------------------------
Public Sub tracePackageError(oPackage As DTS.Package)
Dim ErrorCode As Long
Dim ErrorSource As String
Dim ErrorDescription As String
Dim ErrorHelpFile As String
Dim ErrorHelpContext As Long
Dim ErrorIDofInterfaceWithError As String
Dim i As Integer

        For i = 1 To oPackage.Steps.Count
                If oPackage.Steps(i).ExecutionResult = DTSStepExecResult_Failure Then
                        oPackage.Steps(i).GetExecutionErrorInfo ErrorCode, ErrorSource, ErrorDescription, _
                                        ErrorHelpFile, ErrorHelpContext, ErrorIDofInterfaceWithError
                        MsgBox oPackage.Steps(i).Name & " failed" & vbCrLf & ErrorSource & vbCrLf & ErrorDescription
                End If
        Next i

End Sub

'------------- define Task_Sub1 for task Copy Data from BIC to [PJ].[dbo].[tBIC] Task (Copy Data from BIC to [PJ].[dbo].[tBIC] Task)
Public Sub Task_Sub1(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask1 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
oTask.Name = "Copy Data from BIC to [PJ].[dbo].[tBIC] Task"
Set oCustomTask1 = oTask.CustomTask

        oCustomTask1.Name = "Copy Data from BIC to [PJ].[dbo].[tBIC] Task"
        oCustomTask1.Description = "Copy Data from BIC to [PJ].[dbo].[tBIC] Task"
        oCustomTask1.SourceConnectionID = 1
        oCustomTask1.SourceObjectName = pSource
        oCustomTask1.DestinationConnectionID = 2
        oCustomTask1.DestinationObjectName = "[PJ].[dbo].[tBIC]"
        oCustomTask1.ProgressRowCount = 1000
        oCustomTask1.MaximumErrorCount = 0
        oCustomTask1.FetchBufferSize = 1
        oCustomTask1.UseFastLoad = True
        oCustomTask1.InsertCommitSize = 0
        oCustomTask1.ExceptionFileColumnDelimiter = "|"
        oCustomTask1.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask1.AllowIdentityInserts = False
        oCustomTask1.FirstRow = 0
        oCustomTask1.LastRow = 0
        oCustomTask1.FastLoadOptions = 2
        oCustomTask1.ExceptionFileOptions = 1
        oCustomTask1.DataPumpOptions = 0
        
Call oCustomTask1_Trans_Sub1(oCustomTask1)
                
                
goPackage.Tasks.Add oTask
Set oCustomTask1 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask1_Trans_Sub1(ByVal oCustomTask1 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask1.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4
                
                Set oColumn = oTransformation.SourceColumns.New("Col001", 1)
                        oColumn.Name = "Col001"
                        oColumn.Ordinal = 1
                        oColumn.FLAGS = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col002", 2)
                        oColumn.Name = "Col002"
                        oColumn.Ordinal = 2
                        oColumn.FLAGS = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("BIC_Code", 2)
                        oColumn.Name = "BIC_Code"
                        oColumn.Ordinal = 2
                        oColumn.FLAGS = 104
                        oColumn.Size = 10
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("BIC_Description", 3)
                        oColumn.Name = "BIC_Description"
                        oColumn.Ordinal = 3
                        oColumn.FLAGS = 104
                        oColumn.Size = 70
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties

                
        Set oTransProps = Nothing

        oCustomTask1.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

