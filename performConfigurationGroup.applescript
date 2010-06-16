--NAME:  performConfigurationGroup
--VERSION: 1.0
--PURPOSE: 
--NOTES:


-- Prompt for database
set _db_list to getOpenDatabases()
tell application "Finder"
	set _db_selected to (choose from list _db_list with prompt "Select database:") as text
	if _db_selected is (false as text) then
		display dialog "Selection is empty. Aborting." buttons {"OK"}
		error number -128
	end if
end tell

-- Prompt for configuration(s)
set _config_list to getFieldValues(_db_selected, "Configuration", "Name")
tell application "Finder"
	set _config_selected_list to choose from list _config_list with prompt "Select configuration(s):" with multiple selections allowed
	if _config_selected_list is (false as text) then
		display dialog "Selection is empty. Aborting." buttons {"OK"}
		error number -128
	end if
	
	-- Perform each configuration
	repeat with _config_id from 1 to the count of _config_selected_list
		my setField(_db_selected, "Global", "Utility_AppleScriptParameter_gt", item _config_id of _config_selected_list as text)
		my performScript(_db_selected, "s execute script from applescript", "Global::Utility_AppleScriptResult_gt", "")
		delay 1
		my setField(_db_selected, "Global", "Utility_AppleScriptParameter_gt", "")
	end repeat
end tell

----------------------------------------
-- HANDLERS
----------------------------------------

-- Handler: Prompt user for file
on promptForDatabase(dbList)
	tell application "Finder"
		set _db_selected to choose from list dbList with prompt "Select database:"
	end tell
end promptForDatabase

-- Handler: Returns list of all values for a specified field
--  Use document class to retrieve values for found set only
on getFieldValues(theDatabase, tableName, fieldName)
	tell application "FileMaker Pro Advanced"
		tell database (theDatabase as text)
			field (fieldName as text) of table (tableName as text)
		end tell
	end tell
end getFieldValues

-- Handler: Sets field value
on setField(databaseName, tableName, fieldName, theValue)
	tell application "FileMaker Pro Advanced"
		tell database (databaseName as text)
			tell table (tableName as text)
				set field fieldName to theValue
			end tell
		end tell
	end tell
end setField

-- Handler: Performs script, sets result to global field, then performs callback script
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

-- Handler: Splits fully qualified field name into array {table,field}
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

-- Handler: Returns list of open FileMaker databases
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
