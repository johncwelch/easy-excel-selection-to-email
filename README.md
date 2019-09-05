# easy-excel-selection-to-email

This is a script I built to take partial table selections in excel, slap them along with a generic header into a set of basic HTML and shove it all into outlook.

it goes through BBEdit for two reasons: 1) Outlook's dictionary hasn't been updated since 2011 in a minor update to Outlook 2011, and honestly, the Outlook management has no interest whatsoever in improving automation. So you can't paste via script. You also can't do it via GUI scripting because their object model in an outgoing message is so fubar'd, even VoiceOver has some issues.

The sad thing is, it's still the best email client on the Mac in terms of scripting. Apple's Mail is FAR worse. Apple supports AppleScript worse than anyone.

I could sort of do it in word, but "send html email" doesn't seem to work in Word, so it's either attachment or nothing, and pasting excel in a functional way in Word is astoudingly painful.

So I use BBEdit, which cares about scripting and works really well. and their tech support even helps you with scripting, something neither MS nor Apple will. Well, MS may if it's VBA. If it's AppleScript, yeah good luck on that.
