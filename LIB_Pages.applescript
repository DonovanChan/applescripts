---------------------------------------------
--	SCRIPT LIBRARY: Pages
---------------------------------------------

-- HELPFUL REFERENCES
--	http://iworkautomation.com/pages/index.html


-- HANDLER: Replaces text in body of Pages document
--	Using methods of body text object because coercing it to text and using the typical text manipulation methods will remove the rich text formatting 
to substituteInBody(theDocument, searchValue, replaceValue)
	tell application "Pages"
		tell theDocument
			tell body text
				set matchStart to offset of searchValue in (every character as text)
				set matchEnd to matchStart + (length of searchValue)
				delete (characters (matchStart + 1) thru (matchEnd - 1))
				set character matchStart to replaceValue
			end tell
		end tell
	end tell
end substituteInBody