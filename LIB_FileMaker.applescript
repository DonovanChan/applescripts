---------------------------------------------
--	SCRIPT LIBRARY: FILEMAKER
---------------------------------------------

--Handler: Prompt user for file
on promptForDatabase(dbList)
	tell application "Finder"
		set _db_selected to choose from list dbList with prompt "Select database:"
	end tell
end promptForDatabase

--Handler: Returns list of all values for a specified field
--  Use document class to retrieve values for found set only
on getFieldValues(theDatabase, tableName, fieldName)
	tell application "FileMaker Pro Advanced"
		tell database (theDatabase as text)
			field (fieldName as text) of table (tableName as text)
		end tell
	end tell
end getFieldValues

--Handler: Sets FileMaker field value
on setField(databaseName, tableName, fieldName, theValue)
	tell application "FileMaker Pro Advanced"
		tell database (databaseName as text)
			tell table (tableName as text)
				set field fieldName to theValue
			end tell
		end tell
	end tell
end setField

--Handler: Performs script, sets result to global field, then performs callback script
--    Required handlers: splitFieldName()
on performScript(databaseName, scriptName, resultFieldNameFull, callbackScriptName)
	tell application "FileMaker Pro Advanced"
		tell database databaseName
			set _result to ""
			try
				do script scriptName
			on error _errorText number _errorNumber from _object partial result _partial to _to
				set _result to ("Error " & _errorNumber & ": " & _errorText & return & ¬
					"Object: " & _object & return & ¬
					"Expected Type: " & _to & return & ¬
					"Partial Result: " & _partial) as text
			end try
		end tell
	end tell
	-- Ensure process is complete
	delay 1
	-- Store result in designated global field
	tell application "FileMaker Pro Advanced"
		tell database databaseName
			if resultFieldNameFull is not "" then
				set _result_field_items to my splitFieldName(resultFieldNameFull)
				set _result_table to item 1 of _result_field_items
				set _result_field to item 2 of _result_field_items
				tell table _result_table
					set field _result_field to _result
				end tell
			end if
			if callbackScriptName is not "" then
				do script callbackScriptName
			end if
		end tell
	end tell
end performScript

--Handler: Splits fully qualified field name into array {table,field}
to splitFieldName(fullyQualifiedFieldName)
	set _delim to AppleScript's text item delimiters
	set AppleScript's text item delimiters to "::"
	set _field_items to fullyQualifiedFieldName's text items
	set AppleScript's text item delimiters to _delim
	if (count of _field_items) < 2 then
		display dialog "Error: Please specify the table within this field name: " & fullyQualifiedFieldName
		error
	else
		return _field_items
	end if
end splitFieldName

--Handler: Returns list of open FileMaker databases
on getOpenDatabases()
	tell application "FileMaker Pro Advanced"
		set _file_list to {}
		copy number of databases to _file_count
		if _file_count = 0 then
			display dialog "There are no FileMaker databases open!" buttons "Ok" with icon stop giving up after 5
		else
			repeat with _file_cur from 1 to _file_count
				--get name of each file
				copy name of database _file_cur to end of _file_list
			end repeat
		end if
		_file_list
	end tell
end getOpenDatabases

--Handler: Returns list of script names for specified FileMaker document
on getScriptNames(theDatabase)
	tell application "FileMaker Pro Advanced"
		set myTblNum to 1
		set myDB to theDatabase
		set myScriptList to {}
		set myScriptIdList to {}
		set myScriptCount to the count of every FileMaker script
		if myScriptCount > 1 then
			set myScriptIdList to the ID of every FileMaker script
			repeat with i from 1 to the count of myScriptIdList
				set myScriptItem to the name of FileMaker script ID (item i of myScriptIdList) of database myDB
				set myScriptList to myScriptList & myScriptItem & tab
				set i to i + 1
			end repeat
			myScriptList
		else
			if myScriptCount = 1 then
				set myScriptName to the name of FileMaker script 1 of database myDB
			else
				set myScriptName to ""
			end if
			myScriptName
		end if
	end tell
end getScriptNames

--Handler: Returns list of table names of frontmost FileMaker file
on getTableNames
	tell application "FileMaker Pro Advanced"
		set myTableList to {}
		copy number of tables to mytblcount
		if mytblcount = 0 then
			display dialog "There are no FileMaker 7 databases open!" buttons "Ok" with icon stop giving up after 5
		else
			repeat with tblLoop from 1 to mytblcount
				--get name of each table
				copy name of table tblLoop to end of myTableList
			end repeat
		end if
		myTableList
	end tell
end getTableNames

--Handler: Converts xml text to FileMaker clipboard format
--Parameters: clipText, outputFormat [script|script_step|table|field|custom_function]
--Methodology: Write text to temp file so that it can be converted from file
--Formats:
--	XMSC for script definitions
--	XMSS for script steps
--	XMTB for table definitions
--	XMFD for field definitions
--	XMCF for custom functions
on convertClip(clipText, outputFormat)
	set temp_path to (path to temporary items as Unicode text) & "FMClip.dat"
	set temp_ref to open for access file temp_path with write permission
	set eof temp_ref to 0
	write clipText to temp_ref
	close access temp_ref
	if outputFormat is "script" then
		set clipTextFormatted to read file temp_path as «class XMSC»
	else if outputFormat is "script_step" then
		set clipTextFormatted to read file temp_path as «class XMSS»
	else if outputFormat is "table" then
		set clipTextFormatted to read file temp_path as «class XMTB»
	else if outputFormat is "field" then
		set clipTextFormatted to read file temp_path as «class XMFD»
	else if outputFormat is "custom_function" then
		set clipTextFormatted to read file temp_path as «class XMCF»
	end if
	return clipTextFormatted
end convertClip