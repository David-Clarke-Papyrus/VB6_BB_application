' **************************************
' This VBScript will Connect the database file of your choice to any SQL/MSDE server using standard commands.
' It will also create a user of your choice and make them an owner of the database you just connected.
' It handles the Case where you forget to Remove the user before you Detach the database and several other cases that I encountered.
' This script uses SQL-DMO. As a result, you must include the SQL-DMO merge module with your project.
' This script should be run at the end of the Install Execute Sequence.
' To use it you must fill in the following properties in your Install Script:
' DBPathDirectoryName = The directory REFERENCE in the database that equates to the path of where the database file is held. (i.e. "Data")
' DBFileName = File Name of the database to be attached.
' DBName = Name of the database you want the file to be attached as
' DBUserName = Name of the user account that you want created (Leave this blank if you don't want to create one.)
' DBUserPassword = Password For the above account
' SAPassword = Password For the SA account.
' Note that you should install the Scripting host first or know that the target platform already has it before you run this script.
' Also note that you should only execute the script if its the full install, not uninstalls etc.
' ****************************************
' Created by James Hancock of Darwin Consulting.
' Freely redistributable and go ahead and modify it. Please leave this header on it though. Just send me a thank you if you like it. If you don't, be nice!
' email me @: jamie@darwinconsulting.com
' Copywrite 2001 James Hancock and Darwin Consulting Inc.
' www.darwinconsulting.com
' *****************************************
Dim sql, Lgn, usr, FilePath, saPassword
Dim fso, f1
Dim dbPath, dbfileName, dbName, dbUserName, dbInstanceName
Dim errPos
Dim WshNetwork 
Dim ComputerName
Dim foldername
Dim RES

On Error resume Next 'Lovely VBScript error handling...
dbFileName = "WR.MDF"
dbName = "WR"
dbUserName = "sa"
dbUserPassword = ""
dbInstanceName = "WRInstance"
saPassword = ""
dbPath = Session.TargetPath("DATABASEDIR")


FilePath = dbPath & dbFileName


	Set fso = CreateObject("Scripting.FileSystemObject")
	FolderName = fso.GetParentFolderName(dbPath) & "\IMPORT"
	fso.createFolder(FolderName )
'msgbox  "Delete this " & dbPath & "*.ldf"
	fso.DeleteFile dbPath & "*.ldf"
	Set fso = Nothing

err.clear

	Set sql = CreateObject("SQLDMO.SQLServer")


   	Set WshNetwork = CreateObject("WScript.Network")

 	Computername = WshNetwork.ComputerName
	Set wshNetwork = Nothing

'MsgBox "Starting  :" & ComputerName & "\WRInstance"
err.clear
	sql.Start true,ComputerName & "\WRInstance","sa","" 'first parameter determines whether a connection is made or not

'	If err.number <> 0 Then
'		MsgBox "Could not Start db.  Setup will terminate with errors: " & err.description & errPos & " " & "MSSQL$" & dbInstanceName
'	else



		sql.detachDB("WR")
'MsgBox "Result of Detach = " & Err.description
err.clear
		Res =  sql.AttachDBWithSingleFile("WR", FilePath) 'If no error attach the user to the database that you just attached.
'MsgBox "RES = " & RES
		If Not (Len(Res) = 0) then
			'You want to leave the error here so that the installer can deal with it.
			MsgBox "The database could not be attached because: " & err.description & errPos
		End If
		SQL.Disconnect
		Set SQL = Nothing
	end if