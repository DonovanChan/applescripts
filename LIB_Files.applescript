---------------------------------------------
--	SCRIPT LIBRARY: FILE  MANAGEMENT
---------------------------------------------

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

-- Handler: Saves text to file
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

-- Handler: Returns POSIX path of file as text
to posixPath(filePath)
	return POSIX file (POSIX path of filePath) as text
end posixPath

-- Handler: Returns path as text with startup disk removed
to stripStartupDisk(thePath)
	set pathText to thePath as text
	tell application "Finder"
		set hdName to name of startup disk
	end tell
	set hdLen to length of hdName
	if text 1 thru hdLen of pathText is equal to hdName then
		set pathText to text (hdLen + 1) thru -1 of pathText
	end if
	return pathText
end stripStartupDisk
