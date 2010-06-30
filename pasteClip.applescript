--NAME: PasteClip
--VERSION: 1.0
--PURPOSE: Pastes FileMaker object into active FileMaker file using object from defined file alias
--HISTORY: Created 2010.06.30 by Donovan Chandler, donovan_c@beezwax.net
--NOTES: 

------------------------------------------------
---- Settings ----
--Placeholder, replaced with calculated date
set strDatePlaceholder to "%DATE%"

--Alias to XML file containting FileMaker clip
set clipAlias to alias "Macintosh HD:Users:Donovan:Documents:Repository:Clip Manager Clips:Scripts:Partial:scs_focusHeader.xmss"

--Format of FileMaker object in clip [script|script_step|table|field|custom_function]
set clipClass to "script_step"

--Paste clip in automatically?
set autoPasteClip to false

------------------------------------------------
---- Format Date ----
set strDateFull to do shell script "date '+%Y-%h-%d %Hh%M'"

---- Edit Clip ----
set clipText to readFile(clipAlias)
set clipTextNew to searchReplaceText(clipText, strDatePlaceholder, strDateFull)

---- Convert clip XMSC ----
set clipTextFormatted to convertClip(clipTextNew, clipClass)

---- Insert Clip ----
set the clipboard to clipTextFormatted
if autoPasteClip is true then paste()

------------------------------------------------
-- HANDLERS
------------------------------------------------

--Handler: Performs paste
on paste()
	tell application "System Events" to keystroke "v" using {command down}
end paste

--Handler: Searches and replaces string within text block
to searchReplaceText(theText, searchString, replaceString)
	set searchString to searchString as list
	set replaceString to replaceString as list
	set theText to theText as text
	
	set oldTID to AppleScript's text item delimiters
	repeat with i from 1 to count searchString
		set AppleScript's text item delimiters to searchString's item i
		set theText to theText's text items
		set AppleScript's text item delimiters to replaceString's item i
		set theText to theText as text
	end repeat
	set AppleScript's text item delimiters to oldTID
	
	return theText
end searchReplaceText

--Handler: Returns text from file.  Prompts for file if no alias specified.
on readFile(fileAlias)
	if fileAlias = "" then
		set theFile to choose file with prompt (localized string "chooseFile")
	else
		set theFile to fileAlias without write permission
	end if
	try
		open for access theFile
		set fileText to (read theFile)
		close access theFile
		return fileText
	on error errMsg number errNum
		try
			close access theFile
		end try
		display dialog errMsg & return & errNum
	end try
end readFile

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