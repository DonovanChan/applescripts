-- HANDLER: Opens url in Safari, using current window if possible
to openInSafari(theURL)
	tell application "Safari"
		activate
		-- Ensure we're not working in the Downloads window
		try
			set frontName to name of front window
			if frontName is "Downloads" then 1 / 0
			if frontName begins with "Source" then 1 / 0
		on error
			make new document with properties {URL:theURL as string}
		end try
	end tell
end openInSafari