---------------------------------------------
--	SCRIPT LIBRARY: FILE  MANAGEMENT
---------------------------------------------

-- HANDLER: Returns text from file.  Prompts for file if no alias specified.
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

-- HANDLER: Saves text to file
on saveText(theText, filePath)
	if filePath = "" then
		set filePath to choose file name with prompt "Choose file to write to"
	end if
	try
		set fileRef to open for access filePath with write permission
		set eof of fileRef to 0
		write theText to fileRef starting at eof
		close access fileRef
	on error
		try
			close access fileRef
		end try
	end try
	return filePath as alias
end saveText

-- HANDLER: Returns AppleScript path of file as text
--	"/Users/joe/Desktop" => "Macintosh HD:Users:joe:Desktop"
to appleScriptPath(filePath)
	return POSIX file (POSIX path of filePath) as text
end appleScriptPath

-- HANDLER: Appends new timestamped line to log file
--	Alternative file name or location must be modified in the code
on updateLog(logEntry)
	do shell script "TEXT=$(echo " & quoted form of logEntry & " ); echo \"`date '+%Y.%m.%d %H:%M:%S'`	$TEXT\" >> ~/Desktop/error_log.txt"
end updateLog

-- HANDLER: Replaces file with timestamped entry
--	Alternative file name or location must be modified in the code
on replaceLog(logEntry, logPath)
	do shell script "TEXT=$(echo " & quoted form of logEntry & " );echo \"`date '+%Y.%m.%d %H:%M:%S'`
$TEXT\" > " & logPath
end replaceLog

-- HANDLER: Returns list of files as return-delimited text string
--	addStartupDrive allows prepending of disk name to each path
--	Required handlers: trimLinesRight
to fileListToText(theList, addStartupDrive)
	set theText to ""
	tell application "Finder"
		set hdName to get name of startup disk
	end tell
	repeat with i in theList
		set myPath to POSIX path of (i as alias)
		if addStartupDrive is true then set myPath to hdName & (myPath as text)
		set theText to theText & myPath & return
	end repeat
	return trimLinesRight(theText)
end fileListToText

-- HANDLER: Removes trailing newlines
to trimLinesRight(theText)
	repeat while theText ends with return
		set theText to theText's text 1 thru -2
	end repeat
	return theText
end trimLinesRight

-- HANDLER: Returns path as text with startup disk removed. Intended for use with posix paths.
to stripStartupDisk(thePath)
	set pathText to thePath as text
	tell application "Finder"
		set hdName to name of startup disk
	end tell
	set hdLen to length of hdName
	if text 1 thru hdLen of pathText is equal to hdName then
		set pathText to text (hdLen + 1) thru -1 of pathText
	else if text 1 thru (hdLen + 1) of pathText is equal to ("/" & hdName) then
		set pathText to text (hdLen + 2) thru -1 of pathText
	end if
	return pathText
end stripStartupDisk

-- HANDLER: Monitors file until it is downloaded (determined by its size being constant)
--	TODO check for version that monitors file.download
to waitForDownload(listOfFileAliases, delayDuration)
	repeat with f in listOfFileAliases
		set oldSize to 0
		set newSize to -1
		-- Wait for size to remain constant
		repeat while newSize is not equal to oldSize
			--set oldSize to size of (info for f)
			set oldSize to size of f
			delay delayDuration
			--set newSize to size of (info for f)
			set newSize to size of f
		end repeat
	end repeat
end waitForDownload

-- HANDLER: Deletes specified files
to deleteFiles(listOfFileAliases)
	--if class of listOfFileAliases is text then set listOfFileAliases to {listOfFileAliases}
	repeat with f in listOfFileAliases
		set the file_path to the quoted form of the POSIX path of f
		do shell script ("rm -fr " & file_path)
	end repeat
end deleteFiles

-- HANDLER: Filters one list of documents by another
to filterDocuments(listToFilter, valuesToOmit)
	set listToFilter to convertDocumentList(listToFilter)
	set valuesToOmit to convertDocumentList(valuesToOmit)
	set myResult to {}
	repeat with i in listToFilter
		if valuesToOmit does not contain i then set end of myResult to i as text
	end repeat
	return myResult
end filterDocuments

-- HANDLER: Returns list of data describing file or folder:
--   the path to its parent directory,
--   its name without its file extension, and
--   its file extension
--   Source: http://www.alecjacobson.com/weblog/?p=229
on fileInfo(this_file)
	set default_delimiters to AppleScript's text item delimiters
	-- if given file is a folder then strip terminal ":" so as to return
	-- folder name as file name and true parent directory
	if last item of (this_file as string) = ":" then
		set AppleScript's text item delimiters to ""
		set this_file to (items 1 through -2 of (this_file as string)) as string
	end if
	set AppleScript's text item delimiters to ":"
	set this_parent_dir to (text items 1 through -2 of (this_file as string)) as string
	set this_name to (text item -1 of (this_file as string)) as string
	-- default or no extension is empty string
	set this_extension to ""
	if this_name contains "." then
		set AppleScript's text item delimiters to "."
		set this_extension to the last text item of this_name
		set this_name to (text items 1 through -2 of this_name) as string
	end if
	set AppleScript's text item delimiters to default_delimiters
	return {this_parent_dir, this_name, this_extension}
end fileInfo

-- HANDLER: Takes list of aliases as text and returns list of file names
--	Note: Requires stripPath() handler
on listAliasesToNames(aliasList)
	set nameList to {}
	repeat with i in aliasList
		set end of nameList to stripPath(i as text)
	end repeat
	return nameList
end listAliasesToNames

-- HANDLER: Strips directories from file path, leaving name only
--	Note: requires lastOffset() handler
on stripPath(thePath)
	set nameStart to (my lastOffset(thePath, ":")) + 1
	return text nameStart thru (length of thePath) of thePath
end stripPath

-- HANDLER: Strips directories from posix-style file path, leaving name only
--	Note: requires lastOffset() handler
on stripPathPOSIX(thePath)
	set nameStart to (my lastOffset(thePath, "/")) + 1
	return text nameStart thru (length of thePath) of thePath
end stripPath

-- HANDLER: Returns alias based on posix path
--	e.g., posixAlias(POSIX file "Macintosh HD/.DS_Store")
on posixAlias(posixPath)
	lstripString(posixPath as text, ":") as alias
end posixAlias

-- HANDLER: Checks if file exists
on fileExists(filePath)
	try
		alias (filePath as text)
		return true
	on error
		return false
	end try
end fileExists