--NAME: ShowFileAlias
--VERSION: 1.0
--PURPOSE: Prompts for file selection then displays alias to selected file
--HISTORY: Created 2010.06.30 by Donovan Chandler, donovan_c@beezwax.net
--NOTES: 

set theFile to choose file with prompt ¬
	"Select your file. Its alias will be copied to the clipboard." default location ¬
	path to home folder

set theFileAlias to theFile as alias
set the clipboard to theFileAlias as text
set theResult to display dialog ¬
	"The following alias has been copied to your clipboard:" default answer theFileAlias as text ¬
	with icon 1 ¬
	buttons {"OK"}