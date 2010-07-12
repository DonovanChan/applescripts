--NAME: CopyPathDropped
--VERSION: 
--PURPOSE: Copies path to file dropped onto droplet
--HISTORY: Created 2010.06.30 by Donovan Chandler, donovan_c@beezwax.net
--NOTES: 

on open {dropped_item}
	tell application "Finder" to set the clipboard to the ¬
		dropped_item as text
end open