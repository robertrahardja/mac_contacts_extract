-- Mac Contacts to CSV Exporter (AppleScript)
-- This script exports all Mac contacts to a CSV file with all fields preserved
-- The CSV can then be easily imported to Google Sheets

on run
	-- Get the desktop path for saving the file
	set desktopPath to (path to desktop) as text
	set currentDate to do shell script "date '+%Y-%m-%d_%H-%M-%S'"
	set fileName to "Contacts_Export_" & currentDate & ".csv"
	set filePath to desktopPath & fileName
	
	-- Create the CSV file
	set csvContent to "First Name,Last Name,Middle Name,Nickname,Company,Job Title,Department,Email 1,Email 2,Email 3,Phone 1,Phone 2,Phone 3,Mobile,Home Address,Work Address,Birthday,Notes,Website,Social Media" & return
	
	tell application "Contacts"
		set allPeople to every person
		
		repeat with aPerson in allPeople
			-- Initialize the row
			set personRow to ""
			
			-- Get basic name information
			try
				set personRow to personRow & quoted form of (first name of aPerson as text)
			on error
				set personRow to personRow & "\"\""
			end try
			set personRow to personRow & ","
			
			try
				set personRow to personRow & quoted form of (last name of aPerson as text)
			on error
				set personRow to personRow & "\"\""
			end try
			set personRow to personRow & ","
			
			try
				set personRow to personRow & quoted form of (middle name of aPerson as text)
			on error
				set personRow to personRow & "\"\""
			end try
			set personRow to personRow & ","
			
			try
				set personRow to personRow & quoted form of (nickname of aPerson as text)
			on error
				set personRow to personRow & "\"\""
			end try
			set personRow to personRow & ","
			
			-- Get organization information
			try
				set personRow to personRow & quoted form of (organization of aPerson as text)
			on error
				set personRow to personRow & "\"\""
			end try
			set personRow to personRow & ","
			
			try
				set personRow to personRow & quoted form of (job title of aPerson as text)
			on error
				set personRow to personRow & "\"\""
			end try
			set personRow to personRow & ","
			
			try
				set personRow to personRow & quoted form of (department of aPerson as text)
			on error
				set personRow to personRow & "\"\""
			end try
			set personRow to personRow & ","
			
			-- Get email addresses (up to 3)
			try
				set emailList to emails of aPerson
				set emailCount to 0
				repeat with anEmail in emailList
					if emailCount < 3 then
						set personRow to personRow & quoted form of (value of anEmail as text) & ","
						set emailCount to emailCount + 1
					end if
				end repeat
				-- Fill remaining email columns
				repeat (3 - emailCount) times
					set personRow to personRow & "\"\","
				end repeat
			on error
				set personRow to personRow & "\"\",\"\",\"\","
			end try
			
			-- Get phone numbers
			set phoneData to {"", "", "", ""}
			try
				set phoneList to phones of aPerson
				set phoneIndex to 1
				repeat with aPhone in phoneList
					if phoneIndex ≤ 4 then
						set phoneLabel to label of aPhone as text
						set phoneValue to value of aPhone as text
						if phoneLabel contains "mobile" or phoneLabel contains "cell" or phoneLabel contains "iPhone" then
							set item 4 of phoneData to phoneValue
						else
							set item phoneIndex of phoneData to phoneValue
							set phoneIndex to phoneIndex + 1
						end if
					end if
				end repeat
			end try
			repeat with phoneNum in phoneData
				set personRow to personRow & quoted form of phoneNum & ","
			end repeat
			
			-- Get addresses
			set homeAddr to ""
			set workAddr to ""
			try
				set addressList to addresses of aPerson
				repeat with anAddress in addressList
					set addrLabel to label of anAddress as text
					set addrString to ""
					
					try
						set addrString to street of anAddress as text
					end try
					try
						set addrString to addrString & " " & (city of anAddress as text)
					end try
					try
						set addrString to addrString & " " & (state of anAddress as text)
					end try
					try
						set addrString to addrString & " " & (zip of anAddress as text)
					end try
					try
						set addrString to addrString & " " & (country of anAddress as text)
					end try
					
					if addrLabel contains "home" then
						set homeAddr to addrString
					else if addrLabel contains "work" then
						set workAddr to addrString
					else if homeAddr is "" then
						set homeAddr to addrString
					end if
				end repeat
			end try
			set personRow to personRow & quoted form of homeAddr & ","
			set personRow to personRow & quoted form of workAddr & ","
			
			-- Get birthday
			try
				set bday to birth date of aPerson
				set bdayString to (month of bday as integer) & "/" & (day of bday) & "/" & (year of bday)
				set personRow to personRow & quoted form of bdayString
			on error
				set personRow to personRow & "\"\""
			end try
			set personRow to personRow & ","
			
			-- Get notes
			try
				set noteText to note of aPerson as text
				-- Remove line breaks and commas that might break CSV
				set AppleScript's text item delimiters to {return, linefeed, character id 8232, character id 8233}
				set noteWords to text items of noteText
				set AppleScript's text item delimiters to " "
				set noteText to noteWords as text
				set AppleScript's text item delimiters to ""
				set personRow to personRow & quoted form of noteText
			on error
				set personRow to personRow & "\"\""
			end try
			set personRow to personRow & ","
			
			-- Get URLs
			set urlString to ""
			try
				set urlList to urls of aPerson
				set urlStrings to {}
				repeat with aURL in urlList
					set end of urlStrings to value of aURL as text
				end repeat
				set AppleScript's text item delimiters to "; "
				set urlString to urlStrings as text
				set AppleScript's text item delimiters to ""
			end try
			set personRow to personRow & quoted form of urlString & ","
			
			-- Get social profiles
			set socialString to ""
			try
				set socialList to social profiles of aPerson
				set socialStrings to {}
				repeat with aSocial in socialList
					try
						set socialService to service name of aSocial as text
						set socialUser to user name of aSocial as text
						set end of socialStrings to socialService & ": " & socialUser
					end try
				end repeat
				set AppleScript's text item delimiters to "; "
				set socialString to socialStrings as text
				set AppleScript's text item delimiters to ""
			end try
			set personRow to personRow & quoted form of socialString
			
			-- Add the row to CSV content
			set csvContent to csvContent & personRow & return
		end repeat
	end tell
	
	-- Write to file
	try
		set fileRef to open for access file filePath with write permission
		set eof fileRef to 0
		write csvContent to fileRef as «class utf8»
		close access fileRef
		
		-- Show success dialog with options
		display dialog "Successfully exported " & (count of allPeople) & " contacts to:" & return & return & fileName & return & return & "The file is saved on your Desktop." & return & return & "Would you like to open Google Sheets now to import it?" buttons {"No Thanks", "Open Google Sheets"} default button "Open Google Sheets" with title "Export Complete"
		
		if button returned of result is "Open Google Sheets" then
			open location "https://sheets.google.com"
			
			-- Show import instructions
			display dialog "To import your contacts in Google Sheets:" & return & return & "1. Click 'Blank' to create a new spreadsheet" & return & "2. Click File → Import" & return & "3. Click 'Upload' and select the CSV file from your Desktop" & return & "4. Choose 'Replace current sheet'" & return & "5. Click 'Import data'" buttons {"OK"} default button "OK" with title "Import Instructions"
		end if
		
	on error errMsg
		display dialog "Error writing file: " & errMsg buttons {"OK"} default button "OK" with icon stop
	end try
end run
