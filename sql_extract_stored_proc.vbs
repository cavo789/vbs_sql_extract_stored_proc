' ===================================================
'
' Author	: Christophe Avonture
' Date		: May 2018
'
' Connect to a SQL Server database, obtain the list of
' stored procedures (USPs) in that db (process all schemas), get
' the code in these stored procs and save them as text files (.md)
'
' At the end, we'll have as many files as there are stored procs
' in the database. One text file by stored proc.
'
' Files will be saved under the /results folder.
'
' Running this script against a SQL Server DB will take a local
' copy of your USPs : you can then take a backup of them easily.
'
' NOTE : The user should have enough permissions on SQL Server niveau
' 	for retrieving the code of the stored procedure. This is never the
'	case of a "simple" user and requires advanced permissions. So; if
'	generated files are empty, first check user's permissions (or directly
'	use an "admin" user to check if it's better).
'
' Documentation : https://github.com/cavo789/vbs_sql_extract_stored_proc
' ===================================================

Option Explicit

Const cServerName = "" 		' <== Name of your SQL server
Const cDatabaseName = ""	' <== Name of the database
Const cUserName = ""		' <== User name
Const cPassword = ""		' <== User password

Dim sDatabaseName, sServerName, sUserName, sPassword

' ---------------------------------------------------
'
' Show help screen
'
' ---------------------------------------------------
Sub ShowHelp()

	wScript.echo " ==============================="
	wScript.echo " = vbs_sql_extract_stored_proc ="
	wScript.echo " ==============================="
	wScript.echo ""
	wScript.echo " This script requires four parameters : the server, "
	wScript.echo " database name, login and password to use for the "
	wScript.echo "connection."
	wScript.echo ""
	wScript.echo " " & wScript.ScriptName & " 'ServerName', 'dbTest', 'Login', 'Password'"
	wScript.echo ""
	wScript.echo "To get more info, please read https://github.com/cavo789/vbs_sql_extract_stored_proc"
	wScript.echo ""

	wScript.quit

End sub

' ---------------------------------------------------
'
' Return the current, running, folder
'
' ---------------------------------------------------
Public Function getCurrentFolder()

	Dim objFSO, objFile
	Dim sFolder

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile(Wscript.ScriptFullName)

	sFolder = objFSO.GetParentFolderName(objFile) & "\"

	Set objFile = Nothing
	Set objFSO = Nothing

	getCurrentFolder = sFolder

End Function

' ---------------------------------------------------
'
' Create a folder if not yet there
'
' ---------------------------------------------------
Public Function makeFolder(ByVal sFolderName)

	Dim objFSO

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	If Not objFSO.FolderExists(sFolderName) Then
		Call objFSO.CreateFolder(sFolderName)
	End if

	Set objFSO = Nothing

End Function

' ---------------------------------------------------
'
' Remove all files in the specified folder
'
' ---------------------------------------------------
Public Function emptyFolder(ByVal sFolderName)

	Dim objFSO, objFiles, objFile

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	Set objFiles = objFSO.GetFolder(sFolderName).Files

	For Each objFile In objFiles
		objFile.Delete
	Next

	set objFile = Nothing
	set objFSO = Nothing

End function

' ---------------------------------------------------
'
' Create a text file on the disk, UTF-8 with LF
'
' ---------------------------------------------------
Public Sub CreateTextFile(ByVal sFileName, ByVal sContent)

	Dim objStream

	Set objStream = CreateObject("ADODB.Stream")

	With objStream
		.Open
		.CharSet = "x-ansi" ' "UTF-8"
		.LineSeparator = 10
		.Type = 2 ' adTypeText
		.WriteText sContent
		.SaveToFile sFileName, 2
		.Close
	End with

	set objStream = Nothing

End Sub

Dim sDSN, sSQL
Dim objConn, rs
Dim sLine, sPath, sFileName, sProcName, sContent, sMDTable
Dim wProcsCount

	' Get constants
	sServerName = trim(cServerName)
	sDatabaseName = trim(cDatabaseName)
	sUserName = trim(cUserName)
	sPassword = trim(cPassword)

	' If one variable is not set by constants, get from
	' command line arguments
	If (sServerName = "") or (sDatabaseName = "") or _
		(sUserName = "") or (sPassword = "") Then
		If (wScript.Arguments.Count < 4) Then
			Call ShowHelp
			wScript.quit
		Else
			' Read parameters server -> db -> login -> password
			sServerName = Trim(Wscript.Arguments.Item(0))
			sDatabaseName = Trim(Wscript.Arguments.Item(1))
			sUserName = Trim(Wscript.Arguments.Item(2))
			sPassword = Trim(Wscript.Arguments.Item(3))
		End if
	End If

	' Define the results folder : a subfolder of the folder
	' containing this VBS script.
 	sPath = getCurrentFolder() & "results\"
	makeFolder(sPath)

	' Remove files from a previous run
	emptyFolder(sPath)

	wProcsCount = 0

	' Define the connection string
	sDSN = "Driver={SQL Server};Server={" & sServerName & "};" & _
		"Database={" & sDatabaseName & "};" & _
		"User Id={" & sUserName & "};" & _
		"Password={" & sPassword & "};"

	Set objConn = CreateObject("ADODB.Connection")
	Set rs = CreateObject("ADODB.Recordset")

	objConn.ConnectionTimeout = 60
	objConn.CommandTimeout = 60

	objConn.Open sDSN

	' Get the list of tables in the database
	sSQL = "SELECT Specific_Catalog, Specific_Schema, Specific_Name, " & _
		"Routine_Definition, Created, Last_Altered " & _
		"FROM INFORMATION_SCHEMA.Routines " & _
		"WHERE ROUTINE_TYPE = 'PROCEDURE';"

	Set rs = objConn.Execute(sSQL)

	If Not rs.Eof Then

		' Iterate for each table
		Do While Not rs.EOF

			' Derive the filename :
			'	* The database name (f.i. dbAdmin)
			'	* The schema (f.i. dbo)
			'	* The stored proc name (f.i. uspGetData)
			sProcName = rs.Fields("Specific_Catalog").Value & _
				"." & rs.Fields("Specific_Schema").Value & _
				"." & rs.Fields("Specific_Name").Value

			sFileName = sPath & replace(sProcName, ".", "_") & ".md"

			' Get the stored procedure content i.e. the programmation
			sContent = "```SQL" & vbLf  & _
				rs.fields("Routine_Definition").Value & vbLf & _
				"```"

			sContent = "# Stored procedure definition" & vbLf & vbLf & _
				"> Created on " & rs.fields("Created").Value & " | " & vbLf & _
				"> Last updated on " & rs.fields("Last_Altered").Value & vbLf & vbLf & _
				sContent & vbLf

			Call CreateTextFile(sFileName, sContent)

			rs.MoveNext
		Loop

	End if

	rs.Close

	Set rs = Nothing
	Set objConn = Nothing