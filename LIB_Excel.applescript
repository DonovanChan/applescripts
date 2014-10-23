
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

on findCellWithValue(findWhat)
	try
		tell application "Microsoft Excel" to return my parseCellAddress((find (range "A1:IV65536") what findWhat look at whole search order by rows))
	on error
		return false
	end try
end findCellWithValue

on findInRange(findWhat, findRange)
	try
		tell application "Microsoft Excel" to return my parseCellAddress((find (range findRange) what findWhat look at whole))
		return rowC of temp
	on error
		return false
	end try
end findInRange

on findInRow(findWhat, rown)
	try
		tell application "Microsoft Excel" to return colC of (my parseCellAddress((find (range ("A" & rown & ":" & my findLastColumn(rown) & rown)) what findWhat look at whole)))
	on error
		return false
	end try
end findInRow

on findInColumn(findWhat, coln)
	try
		tell application "Microsoft Excel" to return rowC of my parseCellAddress((find (range (coln & 1 & ":" & coln & my findLastRow(coln))) what findWhat look at whole))
	on error
		return false
	end try
end findInColumn

on getListFromRow(rown)
	tell application "Microsoft Excel"
		set temp to formula of range ("A" & rown & ":" & my findLastColumn(rown) & rown)
		return item 1 of temp
	end tell
end getListFromRow

on getListFromColumn(coln)
	tell application "Microsoft Excel" to return formula of range (coln & 1 & ":" & coln & my findLastRow(coln))
end getListFromColumn

on findLastRow(coln)
	tell application "Microsoft Excel" to return (rowC of my parseCellAddress((get end range (coln & "65536") direction toward the top)))
end findLastRow

on findLastColumn(rown)
	tell application "Microsoft Excel" to return (colC of my parseCellAddress((get end range ("iv" & rown) direction toward the left)))
end findLastColumn

on parseCellAddress(celladdress)
	tell application "Microsoft Excel" to set tempcell to get address of celladdress
	set AppleScript's text item delimiters to "$"
	set cellInfo to text items of tempcell
	set AppleScript's text item delimiters to ""
	return {colC:item 2 of cellInfo, rowC:item 3 of cellInfo}
end parseCellAddress

on getActiveRow()
	tell application "Microsoft Excel" to return rowC of (my parseCellAddress(active cell))
end getActiveRow

on setCell(rangev, cellval)
	tell application "Microsoft Excel" to set formula of range rangev to cellval
end setCell

on getCell(rangev)
	tell application "Microsoft Excel" to return formula of range (rangev)
end getCell

on getCellValue(rangev)
	tell application "Microsoft Excel" to return value of range (rangev)
end getCellValue

on parseColorCode(colorName)
	set colorList to {{"", "0"}, {"black", "1"}, {"white", "2"}, {"red", "3"}, {"green", "4"}, {"blue", "5"}, {"yellow", "6"}, {"pink", "7"}}
	repeat with i in colorList
		ignoring case
			if colorName = item 1 of i then return (item 2 of i as number)
		end ignoring
	end repeat
	return 0
end parseColorCode

on setCellFontColor(rangev, colornum)
	if (class of colornum is not integer) then set colornum to my parseColorCode(colornum)
	tell application "Microsoft Excel" to set font color index of font object of range rangev of active sheet to colornum
end setCellFontColor

on setCellColor(rangev, colornum)
	if (class of colornum is not integer) then set colornum to my parseColorCode(colornum)
	tell application "Microsoft Excel" to set color index of interior object of range rangev of active sheet to colornum
end setCellColor

on setRowColor(rown, colornum)
	if (class of colornum is not integer) then set colornum to my parseColorCode(colornum)
	tell application "Microsoft Excel" to set color index of interior object of row rown of active sheet to colornum
end setRowColor

on setColumnColor(coln, colornum)
	if (class of colornum is not integer) then set colornum to my parseColorCode(colornum)
	tell application "Microsoft Excel"
		if (class of coln is not integer) then
			set color index of interior object of range (coln & ":" & coln) of active sheet to colornum
		else
			set color index of interior object of column coln of active sheet to colornum
		end if
	end tell
end setColumnColor

on openFile(filepath)
	tell application "Microsoft Excel" to open filepath
end openFile

on closeWithoutSaving()
	tell application "Microsoft Excel" to close active workbook without saving
end closeWithoutSaving

on deleteActiveFile()
	set deskPath to path to desktop as string
	set AppleScript's text item delimiters to ":"
	set deskPath to every text item of deskPath as list
	set AppleScript's text item delimiters to "/"
	set savePath to "/" & (items 2 through 4 of deskPath) & "/" as string
	set AppleScript's text item delimiters to ""
	tell application "Microsoft Excel" to set excelName to name of active workbook
	do shell script ("rm " & savePath & excelName)
end deleteActiveFile

on deleteRow(rownumber)
	tell application "Microsoft Excel" to delete row (rownumber as integer)
end deleteRow

on deleteColumn(col)
	if (class of col is not integer) then set col to my alphaToNumeric(col)
	tell application "Microsoft Excel" to delete column (col as integer)
end deleteColumn

on alphaToNumeric(inchar)
	set alphabet to "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	if (count of inchar) = 1 then return offset of inchar in alphabet

	set letterList to every character of inchar
	set letterValueList to {}
	repeat with i in (items 1 through ((count of inchar) - 1) of letterList as list)
		set letterValueList to letterValueList & {((offset of i in alphabet) * 26)}
	end repeat

	set sum to 0
	repeat with i from 1 to (count of letterValueList) in letterValueList
		set sum to sum + ((item i of letterValueList) * (((count of letterValueList) - i) * 26))
	end repeat

	return sum + (last item of letterValueList) + (offset of (last item of inchar) in alphabet)
end alphaToNumeric

on numericToAlpha(inNumber)
	set alphabet to "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	set temp to ""
	repeat until inNumber is 0
		set modNumber to inNumber mod 26
		set inNumber to inNumber div 26
		if modNumber is 0 then
			set modNumber to 26
			set inNumber to inNumber - 1
		end if
		set temp to (item modNumber of alphabet) & temp
	end repeat
	return temp
end numericToAlpha

-- HANDLER: Save Excel file to CSV
-- Returns AppleScript path to file
-- Uses current file name, overwriting existing csv file if necessary
on saveExcelToCSV(filePath)
	tell application "Microsoft Excel"
		activate
		set _sourceWorkbook to open workbook workbook file name filePath
		set _sourcepath to full name of _sourceWorkbook
		set _newFilename to my stripExtension(my stripPath(_sourcepath)) & ".csv"
		set _sourceWorksheet to worksheet 1 of _sourceWorkbook

		save as _sourceWorksheet ¬
			filename _newFilename ¬
			file format CSV Mac file format with overwrite

		if full name of active workbook is _sourcepath then close active workbook saving yes
		if name of active workbook is _newFilename then
			set _newPath to full name of active workbook
			close active workbook saving yes
		end if
		return _newPath
	end tell
end saveExcelToCSV