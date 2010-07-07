(*
NAME:
	AddColumnsByColor
VERSION:
	3.2
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
set _dataRowStartIndex to _headerRowIndex + 1
_rangeSelectedText to "A1:V562"
set _columnRecordList to {¬
	{columnName:"METI Case", columnIndex:"", colorTests:{¬
		{colorIndex:4, colorName:"dkGreen", fieldValue:1} ¬
			}}, ¬
	{columnName:"CAA Case", columnIndex:"", colorTests:{¬
		{colorIndex:20, colorName:"ltGreen", fieldValue:1} ¬
			}}, ¬
	{columnName:"Personal Injury Case", columnIndex:"", colorTests:{¬
		{colorIndex:12, colorName:"brown", fieldValue:1} ¬
			}}, ¬
	{columnName:"Exponent Report Completed", columnIndex:"", colorTests:{¬
		{colorIndex:33, colorName:"turquoise", fieldValue:1} ¬
			}}, ¬
	{columnName:"Unit Location", columnIndex:"", colorTests:{¬
		{colorIndex:15, colorName:"gray", fieldValue:"Unit Sent to US"}, ¬
		{colorIndex:34, colorName:"skyBlue", fieldValue:"Unit in Japan"}, ¬
		{colorIndex:32, colorName:"sage", fieldValue:"Unit to Kimura-san"}, ¬
		{colorIndex:38, colorName:"pink", fieldValue:"Exception"} ¬
			}}, ¬
	{columnName:"iPod Blameless", columnIndex:"", colorTests:{¬
		{colorIndex:41, colorName:"blue", fieldValue:1} ¬
			}}, ¬
	{columnName:"NITE", columnIndex:"", colorTests:{¬
		{colorIndex:46, colorName:"orange", fieldValue:"Sep 10, 2008 (14 Cases)"}, ¬
		{colorIndex:7, colorName:"hotPink", fieldValue:"June 2009 (2 Cases)"}, ¬
		{colorIndex:13, colorName:"indigo", fieldValue:"March 2010 (9 Cases)"} ¬
			}}, ¬
	{columnName:"Model", columnIndex:"", colorTests:{¬
		{colorIndex:54, colorName:"plum", fieldValue:"1 GB"} ¬
			}} ¬
		}

------------------------------------------------
-- Instantiate variable to store list of column values for each row
set _resultRowList to {}

tell application "Microsoft Excel"
	
	-- Gather data about worksheet
	set _worksheet to active sheet
	--set _rangeSelected to range selection of active window
	set _rangeSelected to range _rangeSelectedText of _worksheet
	
	-- Create new columns
	repeat with col from 1 to length of _columnRecordList
		set _columnRecord to (a reference to item col of _columnRecordList)
		set _columnNameCur to columnName of _columnRecord
		set _headerList to my getRowValues(_worksheet, _headerRowIndex)
		set _columnIndex to my getListPosition(_headerList, _columnNameCur)
		-- Create column if necessary
		if _columnIndex is 0 then
			--Create column if necessary
			set _colEndRange to get end (used range of active sheet) direction toward the right
			set _colNextIndex to get (first column index of _colEndRange) + 1
			set _rangeFlag to range (get address of row _headerRowIndex of column _colNextIndex)
			set value of _rangeFlag to _columnNameCur
			set bold of font object of _rangeFlag to true
			set _columnIndex to _colNextIndex
		end if
		-- Update instructions record with column index of fieldName
		--set contents of _columnRecord to (contents of _columnRecord) & {columnIndex:_columnIndex}
		set columnIndex of contents of _columnRecord to _columnIndex
	end repeat
	set _headerList to my getRowValues(_worksheet, _headerRowIndex)
	set _rangeNew to range ¬
		((get address of row _headerRowIndex of column (columnIndex of first item of _columnRecordList)) & ":" & ¬
			(get address of row _headerRowIndex of column (my getMaxColumnIndex(_columnRecordList))))
	
	-- Loop through used rows
	set _rowMaxIndex to first row index of last row of _rangeSelected
	repeat with i from _dataRowStartIndex to _rowMaxIndex
		set _rangeCur to my getRowRange(_rangeSelected, i)
		set _rowCurIndex to first row index of _rangeCur
		set _colColorList to my getCellColorIndex(_rangeCur)
		
		-- Loop through all new column definitions
		set _testResultList to {}
		repeat with j from 1 to length of _columnRecordList
			set _columnRecord to item j of _columnRecordList
			set _columnTestList to colorTests of _columnRecord
			set _columnIndexCur to columnIndex of _columnRecord
			
			-- Loop through all tests for column
			repeat with _testNum from 1 to length of _columnTestList
				set _testCur to item _testNum of _columnTestList
				set _colorCurIndex to colorIndex of _testCur
				
				-- Append correct value to result list
				if _colColorList contains _colorCurIndex then
					set _resultValueCur to fieldValue of _testCur
				else
					set _resultValueCur to ""
				end if
				
				-- Don't append empty values for multiples tests under same column
				set _columnTestListLength to length of _columnTestList
				if _columnTestListLength > 1 then
					if _resultValueCur is not "" then
						set end of _testResultList to _resultValueCur
					else if _testNum is _columnTestListLength then
						set end of _testResultList to ""
					end if
				else
					set end of _testResultList to _resultValueCur
				end if
			end repeat
			
		end repeat
		-- Append new row values to result list
		set _resultRowList to _resultRowList & {_testResultList}
		
	end repeat
	
	-- Send new values to Excel (looping over every row)
	set _resultIndex to 0
	repeat with i from _dataRowStartIndex to length of _resultRowList
		set _resultIndex to _resultIndex + 1
		set _rowRange to my getRowRange(_rangeNew, i)
		set value of _rowRange to item _resultIndex of _resultRowList
	end repeat
end tell

------------------------------------------------
-- HANDLERS
------------------------------------------------
-- Handler: Returns list of values in used area for row
--    Dependencies: getRowRange()
on getRowValues(theWorksheet, rowNumber)
	tell application "Microsoft Excel"
		set _range to my getRowRange(used range of theWorksheet, rowNumber)
		set _rowValueList to value of _range
		item 1 of _rowValueList
	end tell
end getRowValues

-- Handler: Returns specified row of used range (output is class range)
on getRowRangeUsed(rangeUsed, rowNumber)
	tell application "Microsoft Excel"
		if rangeUsed is "" then
			set _rangeUsed to used range of active sheet
		else
			set _rangeUsed to rangeUsed
		end if
		return my getRowRange(_rangeUsed, rowNumber)
	end tell
end getRowRangeUsed

-- Handler: Returns row of specified range (output is class range)
on getRowRange(theRange, rowNumber)
	tell application "Microsoft Excel"
		set _colFirst to first column index of first column of theRange
		set _colLast to first column index of last column of theRange
		return range ¬
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

-- Handler: Returns item in _columnRecordList with highest columnIndex
on getMaxColumnIndex(recordList)
	set _indexMax to 0
	repeat with i from 1 to length of recordList
		set _indexCur to columnIndex of item i of recordList
		if (_indexCur as integer) is greater than (_indexMax as integer) then
			set _indexMax to _indexCur as integer
		end if
	end repeat
	return _indexMax
end getMaxColumnIndex
