-- HANDLER: Opens url in Safari, using current window if possible
to openInSafari(theURL)
	tell application "Safari"
		activate
		try
			make new tab with properties {URL:theURL as string}
		on error
			make new document with properties {URL:theURL as string}
		end try
	end tell
end openInSafari