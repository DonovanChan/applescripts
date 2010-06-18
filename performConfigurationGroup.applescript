--NAME:  performConfigurationGroup
--VERSION: 3.0
--PURPOSE: Performs series of Configurations in FileMaker Metrics file
--HISTORY: Created 2010.06.17 by Donovan Chandler
--NOTES:
--TO DO: Add option to close and reopen the file between configurations

-- Set variables
global field_for_parameter
global field_for_result
global script_wrapper
global script_callback
global field_for_configuration

set field_for_parameter to "Global::Utility_AppleScriptParameter_gt"
set field_for_result to "Global::Utility_AppleScriptResult_gt"
set script_wrapper to "s execute script from applescript"
set script_callback to ""
set field_for_configuration to "Configuration::Name"

-- Prompt for database
set _db_list to getOpenDatabases()
set _db_selected to (choose from list _db_list with prompt "Select database:") as text
if _db_selected is (false as text) then
	my displayCancelMessage()
	error number -128
end if


-- Prompt for configuration(s)
set _config_list to getFieldValues(_db_selected, field_for_configuration)
set _config_selected_list to choose from list _config_list with prompt "Select configuration(s):" with multiple selections allowed
if _config_selected_list is (false as text) then
	displayCancelMessage()
	error number -128
end if

-- Prompt for number of times to repeat set of configurations
try
	set _loop_total to promptForInteger("How many times would you like to perform this set of Configurations?", 1)
on error
	displayCancelMessage()
	error number -128
end try

-- Repeat set of configurations
repeat _loop_total times
	
	-- Perform each configuration
	repeat with _config_id from 1 to the count of _config_selected_list
		setField(_db_selected, field_for_parameter, item _config_id of _config_selected_list as text)
		performScript(_db_selected, script_wrapper, field_for_result, script_callback)
		delay 1
		setField(_db_selected, field_for_parameter, "")
	end repeat
	
end repeat

-- Nofify user of completion
display dialog "Configurations Complete" buttons {"OK"} giving up after 5

----------------------------------------
-- HANDLERS
----------------------------------------

-- Handler: Prompt user for file
on promptForDatabase(dbList)
	tell application "Finder"
		set _db_selected to choose from list dbList with prompt "Select database:"
	end tell
end promptForDatabase

-- Handler: Cancel Message
on displayCancelMessage()
	display dialog "Selection is empty. Aborting." buttons {"OK"} giving up after 5
end displayCancelMessage

-- Handler: Prompts for integer
on promptForInteger(message, defaultAnswer)
	repeat
		set _dialog to display dialog message ¬
			default answer defaultAnswer ¬
			with icon 1 ¬
			buttons {"OK"} ¬
			default button "OK"
		try
			set _result_integer to (text returned of _dialog) as integer
			if _result_integer is not 0 then exit repeat
		end try
		display dialog "Invalid entry. Please enter an integer greater than 0." buttons {"Enter Again", "Cancel"} default button 1
	end repeat
	_result_integer
end promptForInteger

-- Handler: Returns list of all values for a specified field
--    Required handlers: splitFieldName()
--    Use document class to retrieve values for found set only
on getFieldValues(theDatabase, fullyQualifiedFieldName)
	tell application "FileMaker Pro Advanced"
		set _field_items to my splitFieldName(fullyQualifiedFieldName)
		set _table_name to item 1 of _field_items
		set _field_name to item 2 of _field_items
		tell database (theDatabase as text)
			field (_field_name as text) of table (_table_name as text)
		end tell
	end tell
end getFieldValues

-- Handler: Sets field value
--    Required handlers: splitFieldName()
on setField(databaseName, fullyQualifiedFieldName, theValue)
	tell application "FileMaker Pro Advanced"
		set _field_items to my splitFieldName(fullyQualifiedFieldName)
		set _table_name to item 1 of _field_items
		set _field_name to item 2 of _field_items
		tell database (databaseName as text)
			tell table (_table_name as text)
				set field _field_name to theValue
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
	delay 2
	-- Store result in designated global field
	tell application "FileMaker Pro Advanced"
		tell database databaseName
			if resultFieldNameFull is not "" then
				set _result_field_items to my splitFieldName(resultFieldNameFull)
				set _result_table to item 1 of _result_field_items
				set _result_field to item 2 of _result_field_items
				set _continue to true
				repeat while _continue is true
					set _continue to false
					try
						tell table _result_table
							set field _result_field to _result
						end tell
					on error _error_text number _error_number
						delay 10
						set _continue to true
					end try
				end repeat
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
