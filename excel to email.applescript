
--used to see if you want to end the script
property endScript : false

--header for the message. It has to be this kind of thing because outlook is stupid about inserting text. If you want line breaks, you use <br>
property theContent : "Hello all,<br><br>
blahblahblah, blah blah, blah, blah, blah, blah<br><br>
Thanks!<br><br>
your name<br><br><br><br>"

--we actually change theContent, so this is used as a reset.
property theOriginalContent : "Hello all,<br><br>
blahblahblah, blah blah, blah, blah, blah, blah<br><br>
Thanks!<br><br>
your name<br><br><br><br>"

--main loop
repeat while endScript is false
--pop a list of email addresses. Use " " to put blank spaces in the list. you can choose multiple emails
	try
		set theAddresses to choose from list {"email1@company.com", "email2@company.com", " ", "email1@company2.com", " ", "email1@company3.com"} with title "Contact List" with prompt "chose the contacts for this email" with multiple selections allowed
	on error
		quit
	end try
	
	--if you hit cancel, there's no address, so theAddresses is false. It's weird, just roll with it
	if theAddresses is false then
		quit
	end if
	
	--put the current excel selection on the clipboard
	tell application "Microsoft Excel"
		set thewindow to the active window
		set theSelection to selection of thewindow
		copy range theSelection
	end tell
	
	--paste this into a new bbedit window. It comes in as tab-delimited text
	tell application "BBEdit"
		launch
		set theFirstDoc to make new document
		tell active document of project window 1
			paste
		end tell
		--selects all the text. we need this for the table conversion.
		select text of theFirstDoc
	end tell
	
	--make BBEdit active, needed for the GUI scripting blocks
	tell application "BBEdit" to activate
	
	--gets you the convert to table sheet. BBEdit is so easy to do GUI scripting on because OMG THEY DO THINGS CORRECTLY
	tell application "System Events"
		tell process "BBEdit"
			tell menu bar 1
				tell menu bar item "Markup"
					tell menu "Markup"
						tell menu item "Tables"
							tell menu "Tables"
								click menu item "Convert to Table…"
							end tell
						end tell
					end tell
				end tell
			end tell
		end tell
	end tell
	
	--clicks the convert button on the convert sheet. So far, i've not needed any delays. BBEdit does not suck.
	tell application "System Events"
		tell process "BBEdit"
			tell window 1
				tell sheet 1
					click its button "Convert"
				end tell
			end tell
		end tell
	end tell
	
	--so we do a couple final things here 
	tell application "BBEdit"
		--add a bit of padding into the table, makes the email look nicer
		set theResult to replace "table" using "table cellpadding=\"5\"" searching in text of theFirstDoc options {mode:grep, wrap around:true}
		--take all the text in the document, which is our HTML table, and concatenate that on the end of theContent	
		set theContent to theContent & text of theFirstDoc
		--close the document without saving. On a 2019 MacBook Air, up to this point takes about a second. Maybe 2
		close theFirstDoc saving no
	end tell
	
	tell application "Microsoft Outlook"
		--I use this with an exchange account. You can do this with any account type that can send email
		set theExchangeAccount to the first item of (every exchange account whose name is "nameofexchangeaccount")
		--buld an outgoing message with the body as theContent and a hard-coded subject. 
		set theMessage to make new outgoing message with properties {account:theExchangeAccount, content:theContent, subject:"The Subject is..."}
		--go through the list of email addresses and ad them into the recipient list. Since to/cc/bcc are elements, not properties, this is how we do it...
		repeat with x in theAddresses
			make new to recipient at theMessage with properties {email address:{address:x}}
		end repeat
		--actually opent the message so you can see it. You could automate the send if you really want, I do it manually, just in case.
		open theMessage
		--bring outlook to the front. 
		activate
	end tell
	
	--reset theContent to just the header part
	set theContent to theOriginalContent
	
	--pop a dialog to ask if we want to 
	set theReply to button returned of (display dialog "end the script?" buttons {"No", "OK"})
	if theReply is "OK" then
		set endScript to true
	end if
end repeat

