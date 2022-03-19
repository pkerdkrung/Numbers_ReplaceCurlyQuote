#script name: Numbers_ReplaceCurlyQuote
#Purpose: This script is created for a simple job of replacing all stright quotes in a Mac Numbers spreadsheet with curly quotes, in which Replace All functionality cannot be used because it cannot differentiate between opening and closing quotes, as well as the stright quotes for HTML tags. This script will replace all stright quotes with the correct opening and closing curly quotes based on the character that came before or after it, while skipping the ones inside HTML tags. Additionally, if there are HTML tags in a cell, the script will pause and highlight the cell in red, allowing the user to review the cell manually.
#How to use: In Mac Numbers, select a column, or a single-column range cells, that you want to run the script on, then run it from the AppleScript Editor, or you can make this script into a shortcut key and run it from there.

tell application "Numbers"
set theseRange to document 1's sheet's 1 table 1's selection range
set targetRow to 1
set htmlSwitch to 0

repeat count of theseRange's rows times

	set originalText to theseRange's cell targetRow's value
	set charNum to 1

	repeat length of originalText times
		if originalText's item charNum is "<" then
			set htmlSwitch to 1
		else if originalText's item charNum is ">" then
			set htmlSwitch to 0
		else if originalText's item charNum is "\"" and htmlSwitch is equal to 0 then
			if charNum is equal to 1 then
				set originalText to "“" & text 2 thru -1 of originalText
			else if charNum is equal to the length of originalText then
				set originalText to text 1 thru -2 of originalText & "”"
			else if originalText's item (charNum - 1) is in {" ", ">", ":"} then --set more conditions here
				set originalText to text 1 thru (charNum - 1) of originalText & "“" & text (charNum + 1) thru -1 of originalText
			else if originalText's item (charNum + 1) is in {" ", "<", ","} then --set more conditions here
				set originalText to text 1 thru (charNum - 1) of originalText & "”" & text (charNum + 1) thru -1 of originalText
			end if
		end if
		set charNum to (charNum + 1)
	end repeat

	set htmlSwitch to 0
	set theseRange's cell targetRow's value to originalText
	
	if theseRange's cell targetRow's value contains "\"" then
		set selection range of document 1's sheet 1's table 1 to theseRange's cell targetRow
		set theseRange's cell targetRow's background color to {65535, 30000, 30000}
		exit repeat
	else
		set theseRange's cell targetRow's background color to {30000, 65535, 30000}
	end if
	
	set targetRow to targetRow + 1
	
end repeat

end tell
