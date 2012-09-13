(*
NAME: PasteClip
VERSION: 1.0
PURPOSE: Pastes FileMaker object into active FileMaker file using object from defined file alias
HISTORY:
	Created 2010.06.30 by Donovan Chandler, donovan_c@beezwax.net
	Modified 2012.04.03 by Donovan Chandler: Added support for XML2 format (in FM 12)
NOTES: 
*)

------------------------------------------------
---- Settings ----
-- Placeholder, replaced with calculated date
set strDatePlaceholder to "%DATE%"

-- Alias to XML file containting FileMaker clip
--	e.g., "Macintosh HD:Users:Me:Desktop:myClip.xml"
--set clipAlias to "" as alias

-- Alternative alias
--	File in ../Clips
--	Disable if clip path specified above
if true then
	set clipName to "scriptHeaderSteps.xml"
	set clipDir to POSIX file (POSIX path of (path to me) & "/..") as text
	set clipAlias to (clipDir & "Clips:" & clipName) as alias
end if

-- Format of FileMaker object in clip [script|script_step|table|field|custom_function]
set clipClass to "script_step"

-- Paste clip in automatically?
--	Not yet supported
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
--  HANDLERS
------------------------------------------------

-- Handler: Performs paste
on paste()
	tell application "System Events" to keystroke "v" using {command down}
end paste

-- Handler: Searches and replaces string within text block
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

-- Handler: Returns text from file.  Prompts for file if no alias specified.
on readFile(fileAlias)
	if fileAlias = "" then
		set theFile to choose file with prompt (localized string "chooseFile")
	else
		set theFile to fileAlias
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

-- HANDLER: Converts xml text to FileMaker clipboard format
-- Parameters: clipText, outputClass [script|script_step|table|field|custom_function]
-- Methodology: Write text to temp file so that it can be converted from file
-- Formats:
--	XMSC for script definitions
--	XMSS for script steps
--	XMTB for table definitions
--	XMFD for field definitions
--	XMCF for custom functions
--	XML2 for layout objects in FileMaker 12
--	XMLO for layout objects in FileMaker 7-11
on convertClip(clipText, outputClass)
	set temp_path to (path to temporary items as Unicode text) & "FMClip.dat"
	set temp_ref to open for access file temp_path with write permission
	set eof temp_ref to 0
	write clipText to temp_ref as «class utf8»
	close access temp_ref
	if outputClass is "XMSC" then
		set clipTextFormatted to read file temp_path as «class XMSC»
	else if outputClass is "XMSS" then
		set clipTextFormatted to read file temp_path as «class XMSS»
	else if outputClass is "XMTB" then
		set clipTextFormatted to read file temp_path as «class XMTB»
	else if outputClass is "XMFD" then
		set clipTextFormatted to read file temp_path as «class XMFD»
	else if outputClass is "XMFN" then
		set clipTextFormatted to read file temp_path as «class XMFN»
	else if outputClass is "XML2" then
		set clipTextFormatted to read file temp_path as «class XML2»
	else if outputClass is "XMLO" then
		set clipTextFormatted to read file temp_path as «class XMLO»
	else
		return "Error: Snippet class not recognized"
	end if
	return clipTextFormatted
end convertClip