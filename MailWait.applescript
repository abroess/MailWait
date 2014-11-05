(*
	Waiting For Mails to OmniFocus Script
	by simplicityisbliss.com, Sven Fechner
	MailTags project and due date compatibility added by Scott Morrison, Indev Software
	
	Based on an Outbox Rule (Mail Act-On) in Mail.app this script adds specific messages
	to the OmniFocus Inbox with the waiting for context
	
	MailTags is required to automatically set the project and the due date.

	Mail Act-On (www.indev.ca) is required to define the Outbox Rule to only
	create tasks for those outgoing emails that are to be tracked in OmniFocus
	
	A sample Outbox rule may be
	if MailTags Tickle Date is After 0 days today  
	Run Apple Script: [ThisAppleScript]
	
	Version 002
*)


--!! EDIT THE PROPERTIES BELOW TO MEET YOUR NEEDS !!--

-- Do you want the actualy mail body to be added to the notes section of the OmniFocus task?
-- Set to 'true' is or 'false' if no
property mailBody : true

-- Text between mail recipient (the person you are waiting for to come back) and the email subject
property MidFix : "Waiting for "

-- Name of your Waiting For context in OmniFocus
property myWFContext : "Waiting for"

-- Default start time
property timeStart : "5:00:00 AM"

-- Default due time
property timeDue : "15:00:00 PM"

-- Default start to due date interval, in days
property dateInterval : "2"

-- !! STOP EDITING HERE IF NOT FAMILAR WITH APPLESCRIPT !! --


-- on perform_mail_action(theData)
using terms from application "Mail"
	on perform mail action with messages theMessages
		--Get going
		tell application "Mail"
			--		set theMessages to |SelectedMessages| of theData --Extract the messages from the rule
			repeat with theMessage in theMessages
				set theSubject to subject of theMessage
				set theRecipient to name of to recipient of theMessage
				set theMessageID to urlencode(the message id of theMessage) of me
				set theStartDate to ""
				set theDueDate to ""
				
				
				try
					using terms from application "MailTagsHelper"
						set theProject to project of theMessage
						set theStartDate to (current date) as date
						set theDueDate to theStartDate + dateInterval * days
						set theDueDate to my setDueDate(theDueDate)
						if (due date of theMessage) is not "" then
							set theStartDate to (due date of theMessage) as date
							set theDueDate to theStartDate
							set theDueDate to my setDueDate(theDueDate)
						end if
					end using terms from
				on error theError
					
				end try
				-- Check if there is one or more recipients
				try
					if (count of theRecipient) > 1 then
						set theRecipientName to (item 1 of theRecipient & (ASCII character 202) & "and" & (ASCII character 202) & ((count of theRecipient) - 1) as string) & (ASCII character 202) & "more"
					else
						set theRecipientName to item 1 of theRecipient
					end if
					
					set theTaskTitle to MidFix & theSubject & " from " & theRecipientName
					set messageURL to "Message://%3C" & (theMessageID) & "%3E"
					set theBody to messageURL
					if mailBody then set theBody to theBody & return & return & the content of theMessage
					
					-- Add waiting for context task to OmniFocus
					if theProject is not missing value then
						
						tell application "OmniFocus"
							tell default document
								set newTaskProps to {name:theTaskTitle}
								set theContext to context myWFContext
								set theProject to (first flattened project where its name = theProject)
								
								if theProject is not missing value then set newTaskProps to newTaskProps & {name:theProject}
								if theContext is not missing value then set newTaskProps to newTaskProps & {context:theContext}
								if theDueDate is not missing value then set newTaskProps to newTaskProps & {due date:theDueDate}
								if theBody is not missing value then set newTaskProps to newTaskProps & {note:theBody}
								
								tell theProject to make new task with properties newTaskProps
								
							end tell
						end tell
						
					else
						
						tell application "OmniFocus"
							tell default document
								set newTaskProps to {name:theTaskTitle}
								set theContext to context myWFContext
								log theContext
								log theDueDate
								log theBody
								if theContext is not missing value then set newTaskProps to newTaskProps & {context:theContext}
								if theDueDate is not missing value then set newTaskProps to newTaskProps & {due date:theDueDate}
								if theBody is not missing value then set newTaskProps to newTaskProps & {note:theBody}
								
								set newTask to make new inbox task with properties newTaskProps
							end tell
						end tell
						
					end if
					
				on error theError
					do shell script "logger -t outboxrule 'Error : " & theError & "' "
				end try
				
			end repeat
		end tell
		
		
	end perform mail action with messages
end using terms from

-- end perform_mail_action

on urlencode(theText)
	set theTextEnc to ""
	repeat with eachChar in characters of theText
		set useChar to eachChar
		set eachCharNum to ASCII number of eachChar
		if eachCharNum = 32 then
			set useChar to "+"
		else if (eachCharNum ­ 42) and (eachCharNum ­ 95) and (eachCharNum < 45 or eachCharNum > 46) and (eachCharNum < 48 or eachCharNum > 57) and (eachCharNum < 65 or eachCharNum > 90) and (eachCharNum < 97 or eachCharNum > 122) then
			set firstDig to round (eachCharNum / 16) rounding down
			set secondDig to eachCharNum mod 16
			if firstDig > 9 then
				set aNum to firstDig + 55
				set firstDig to ASCII character aNum
			end if
			if secondDig > 9 then
				set aNum to secondDig + 55
				set secondDig to ASCII character aNum
			end if
			set numHex to ("%" & (firstDig as string) & (secondDig as string)) as string
			set useChar to numHex
		end if
		set theTextEnc to theTextEnc & useChar as string
	end repeat
	return theTextEnc
end urlencode

on setStartDate(theStartDate)
	set theDate to (date string of theStartDate)
	set newDate to the (date (theDate & " " & timeStart))
	return newDate
end setStartDate

on setDueDate(theDueDate)
	set theDate to (date string of theDueDate)
	set newDate to the (date (theDate & " " & timeDue))
	return newDate
end setDueDate