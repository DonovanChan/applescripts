
-- Handler: Prompts for workbook and returns its reference
on getWorkbookName()
	set theWorkbookFile to choose file with prompt "Please select an Excel workbook file:"
	set theWorkbookName to name of (info for theWorkbookFile)
	tell application "Microsoft Excel"
		--open theWorkbookFile
		set theWorkbook to workbook theWorkbookName
		return theWorkbook
	end tell
end getWorkbookName