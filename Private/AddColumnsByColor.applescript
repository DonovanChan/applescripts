(*
NAME:
	AddColumnsByColor
VERSION:
	1.0
PURPOSE:
	Uses settings to populate new columns based on which background colors exist within each row.
	Performs against active sheet in Microsoft Excel.
HISTORY: 
	Created 2010.06.30 by Donovan Chandler, donovan_c@beezwax.net
NOTES:
	
*)

------------------------------------------------
---- Settings ----
set _headerRowIndex to 1
set _headerUL to "Unit Location"
set _headerN to "NITE"
set _colorRecordList to {¬
	{colorIndex:4, colorName:"dkGreen", fieldName:"METI Case", fieldValue:1}, ¬
	{colorIndex:20, colorName:"ltGreen", fieldName:"CAA Case", fieldValue:1}, ¬
	{colorIndex:12, colorName:"brown", fieldName:"Personal Injury Case", fieldValue:1}, ¬
	{colorIndex:33, colorName:"turquoise", fieldName:"Exponent Report Completed", fieldValue:1}, ¬
	{colorIndex:15, colorName:"gray", fieldName:_headerUL, fieldValue:"Unit Sent to US"}, ¬
	{colorIndex:34, colorName:"skyBlue", fieldName:_headerUL, fieldValue:"Unit in Japan"}, ¬
	{colorIndex:32, colorName:"sage", fieldName:_headerUL, fieldValue:"Unit to Kimura-san"}, ¬
	{colorIndex:38, colorName:"pink", fieldName:_headerUL, fieldValue:"Exception"}, ¬
	{colorIndex:41, colorName:"blue", fieldName:"iPod Blameless", fieldValue:1}, ¬
	{colorIndex:46, colorName:"orange", fieldName:_headerN, fieldValue:"Sep 10, 2008 (14 Cases)"}, ¬
	{colorIndex:7, colorName:"pink", fieldName:_headerN, fieldValue:"June 2009 (2 Cases)"}, ¬
	{colorIndex:13, colorName:"indigo", fieldName:_headerN, fieldValue:"March 2010 (9 Cases)"}, ¬
	{colorIndex:54, colorName:"plum", fieldName:"Model", fieldValue:"1 GB"} ¬
		}

------------------------------------------------

tell application "Microsoft Excel"
	
	-- Gather data about worksheet
	set _worksheet to active sheet
	set _rangeUsed to used range of _worksheet
	
	-- Loop through used rows
	set _rowMaxIndex to first row index of last row of _rangeUsed
	repeat with i from 1 to _rowMaxIndex
		set _rangeCur to my getRowRange(_worksheet, i)
		set _rowCurIndex to first row index of _rangeCur
		set _colColorList to my getCellColorIndex(_rangeCur)
		
		-- Loop through all colors for each column
		repeat with j from 1 to length of _colorRecordList
			set _colorRecord to item j of _colorRecordList
			set _colorCur to colorIndex of _colorRecord
			
			-- If color exists, enter flag in corresponding column
			if _colColorList contains _colorCur then
				set _fieldNameCur to fieldName of _colorRecord
				set _headerList to my getRowValues(_worksheet, _headerRowIndex)
				set _colFlagIndex to my getListPosition(_headerList, _fieldNameCur)
				-- Create column of necessary
				if _colFlagIndex is 0 then
					--Create column if necessary
					set _colEndRange to get end (used range of active sheet) direction toward the right
					set _colNextIndex to get (first column index of _colEndRange) + 1
					set value of range (get address of row _headerRowIndex of column _colNextIndex) to _fieldNameCur
					set _colFlagIndex to _colNextIndex
				end if
				-- Set flag value in corresponding column
				set _flagRange to range (get address of row _rowCurIndex of column _colFlagIndex)
				set _flagValue to fieldValue of _colorRecord
				set value of _flagRange to _flagValue
				
			end if
		end repeat
		
	end repeat
	
end tell

------------------------------------------------
-- HANDLERS
------------------------------------------------
-- Handler: Returns list of values in used area for row
--    Dependencies: getRowRange()
on getRowValues(theWorksheet, rowNumber)
	tell application "Microsoft Excel"
		set _range to my getRowRange(theWorksheet, rowNumber)
		set _rowValueList to value of _range
		--range ("A" & rowNumber & ":" & (get address of row rowNumber of column _colLast))
		item 1 of _rowValueList
	end tell
end getRowValues

-- Handler: Returns specified row of used range (output is class range)
on getRowRange(theWorksheet, rowNumber)
	tell application "Microsoft Excel"
		set _rangeUsed to used range of theWorksheet
		set _colFirst to first column index of first column of _rangeUsed
		set _colLast to first column index of last column of _rangeUsed
		set _rangeRow to range ¬
			((get address of row rowNumber of column _colFirst) & ":" & ¬
				(get address of row rowNumber of column _colLast))
	end tell
end getRowRange

-- Handler: Returns list of color indexes for specified range
on getCellColorIndex(theRange)
	tell application "Microsoft Excel"
		set _result to {}
		if class of selection is range then
			set _range to theRange
		else
			set _range to theRange as range
		end if
		set _cellCount to count of cells of _range
		repeat with i from 1 to _cellCount
			set _colorIndex to color index of interior object of cell i of _range
			set end of _result to _colorIndex
		end repeat
		return _result
	end tell
end getCellColorIndex

-- Handler: Returns offset of item in list
on getListPosition(theList, theItem)
	repeat with i from 1 to the count of theList
		if item i of theList is theItem then return i
	end repeat
	return 0
end getListPosition

------------------------------------------------
-- HANDLERS - UNUSED
------------------------------------------------

-- Handler: Returns range of column following used area, with rows as high as used area.
on getNextColumnRange(theWorksheet)
	tell application "Microsoft Excel"
		set _worksheet to theWorksheet
		set _rangeUsed to used range of _worksheet
		set _rowFirst to first row index of _rangeUsed
		set _rowLast to first row index of last row of _rangeUsed
		set _colLast to first column index of last column of _rangeUsed
		set _colNext to (_colLast + 1)
		
		set _rangeNewStart to (get address of cell _rowFirst of column _colNext)
		set _rangeNewEnd to (get address of cell _rowLast of column _colNext)
		set _rangeNew to range (_rangeNewStart & ":" & _rangeNewEnd)
	end tell
end getNextColumnRange

(*
NAME:
	flagRows(theWorksheet, theRange, colorIndexList, flagList, flagRange)
PURPOSE:
	Iterates over cells in theRange. If current cell has background color found in colorIndexList, corresponding item from flagList will be set to cell in flagRange.
EXAMPLE:
	flagRows( active sheet,"B2:B30",{"33","38"},{"blue","pink"},"W2:W30")
*)
on flagRows(theWorksheet, theRange, colorIndexList, flagList, flagRange)
	tell application "Microsoft Excel"
		set _worksheet to theWorksheet
		if class of selection is range then
			set _range to theRange
		else
			set _range to theRange as range
		end if
		set _colorList to colorIndexList as list
		set _flagList to flagList as list
		set _flagRange to range flagRange
		
		set _cellCount to count of cells of _range
		repeat with i from 1 to _cellCount
			set _colorFoundCur to color index of interior object of cell i of _range
			
			repeat with j from 1 to (length of _colorList)
				set _colorCheckCur to item j of _colorList
				if _colorFoundCur is _colorCheckCur then
					set _cellFlagged to cell i of _flagRange
					set _flagCur to item j of _flagList
					set value of _cellFlagged to _flagCur
				end if
			end repeat
		end repeat
	end tell
end flagRows