(*
	Create task from Message
	
	Copyright (c) Microsoft Corporation.  All rights reserved.
*)

tell application "Microsoft Outlook"
	
	-- get the currently selected message or messages
	set selectedMessages to current messages
	
	-- if there are no messages selected, warn the user and then quit
	if selectedMessages is {} then
		display dialog "Please select a message first and then run this script." with icon 1
		return
	end if
	
	repeat with theMessage in selectedMessages
		
		-- get the information from the message, and store it in variables
		set theName to subject of theMessage
		set theCategory to category of theMessage
		set thePriority to priority of theMessage
		set theContent to content of theMessage
		
		-- create a new task with the information from the message
		set newTask to make new task with properties {name:theName, content:theContent, category:theCategory, priority:thePriority}
		
		set todo flag of newTask to not completed
		set categories of newTask to {category "1-1.ASAP"}
		
	end repeat
	
	-- if there was only one message selected, then open that new task
	if (count of selectedMessages) = 1 then open newTask
end tell