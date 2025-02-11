--(2)20-Feb-2021, Add option to add existing files and option to only mail body text
--The HTML body example align left now instead of the default center
--Ron de Bruin, https://macexcel.com/examples/mailpdf/macoutlook/

on CreateMailInOutlook(paramString)
	set {fieldValue1, fieldValue2, fieldValue3, fieldValue4, fieldValue5, fieldValue6, fieldValue7, fieldValue8, fieldValue9, fieldValue10} to SplitString(paramString, ";")
	
	tell application "Microsoft Outlook"
		if fieldValue7 = "pop" then
			set theAccount to the first pop account whose name is fieldValue8
			set NewMail to (make new outgoing message with properties {subject:fieldValue1, content:fieldValue2, account:theAccount})
		else if fieldValue7 = "imap" then
			set theAccount to the first imap account whose name is fieldValue8
			set NewMail to (make new outgoing message with properties {subject:fieldValue1, content:fieldValue2, account:theAccount})
		else if fieldValue7 = "exchange" then
			set theAccount to the first exchange account whose name is fieldValue8
			set NewMail to (make new outgoing message with properties {subject:fieldValue1, content:fieldValue2, account:theAccount})
		else
			set NewMail to (make new outgoing message with properties {subject:fieldValue1, content:fieldValue2})
		end if
		
		tell NewMail
			repeat with toRecipient in my SplitString(fieldValue3, ",")
				make new to recipient at end of to recipients with properties {email address:{address:contents of toRecipient}}
			end repeat
			repeat with toRecipient in my SplitString(fieldValue4, ",")
				make new to recipient at end of cc recipients with properties {email address:{address:contents of toRecipient}}
			end repeat
			repeat with toRecipient in my SplitString(fieldValue5, ",")
				make new to recipient at end of bcc recipients with properties {email address:{address:contents of toRecipient}}
			end repeat
			
			if fieldValue9 is not equal to "" then
				set ExcelAttachment to POSIX file fieldValue9
				make new attachment with properties {file:ExcelAttachment as alias}
				delay 0.5
			end if
			
			repeat with AttachmentPath in my SplitString(fieldValue10, ",")
				set AttachmentPath to POSIX file AttachmentPath
				make new attachment with properties {file:AttachmentPath as alias}
				delay 0.5
			end repeat
			
			if fieldValue6 = "yes" then
				open NewMail
				activate NewMail
			else
				send NewMail
			end if
		end tell
	end tell
end CreateMailInOutlook

on CreateMailInOutlookBody(paramString)
	set {fieldValue1, fieldValue2, fieldValue3, fieldValue4, fieldValue5, fieldValue6, fieldValue7, fieldValue8, fieldValue9} to SplitString(paramString, ";")
	
	tell application "Microsoft Outlook"
		set theFile to fieldValue2
		open for access theFile
		set HTMLMessage to (read theFile)
		close access theFile
		
		if fieldValue7 = "pop" then
			set theAccount to the first pop account whose name is fieldValue8
			set NewMail to (make new outgoing message with properties {subject:fieldValue1, content:HTMLMessage, account:theAccount})
		else if fieldValue7 = "imap" then
			set theAccount to the first imap account whose name is fieldValue8
			set NewMail to (make new outgoing message with properties {subject:fieldValue1, content:HTMLMessage, account:theAccount})
		else if fieldValue7 = "exchange" then
			set theAccount to the first exchange account whose name is fieldValue8
			set NewMail to (make new outgoing message with properties {subject:fieldValue1, content:HTMLMessage, account:theAccount})
		else
			set NewMail to (make new outgoing message with properties {subject:fieldValue1, content:HTMLMessage})
		end if
		
		tell NewMail
			repeat with toRecipient in my SplitString(fieldValue3, ",")
				make new to recipient at end of to recipients with properties {email address:{address:contents of toRecipient}}
			end repeat
			repeat with toRecipient in my SplitString(fieldValue4, ",")
				make new to recipient at end of cc recipients with properties {email address:{address:contents of toRecipient}}
			end repeat
			repeat with toRecipient in my SplitString(fieldValue5, ",")
				make new to recipient at end of bcc recipients with properties {email address:{address:contents of toRecipient}}
			end repeat
			
			repeat with AttachmentPath in my SplitString(fieldValue9, ",")
				set AttachmentPath to POSIX file AttachmentPath
				make new attachment with properties {file:AttachmentPath as alias}
				delay 0.5
			end repeat
			
			if fieldValue6 = "yes" then
				open NewMail
				activate NewMail
			else
				send NewMail
			end if
		end tell
	end tell
end CreateMailInOutlookBody

on SplitString(TheBigString, fieldSeparator)
	tell AppleScript
		set oldTID to text item delimiters
		set text item delimiters to fieldSeparator
		set theItems to text items of TheBigString
		set text item delimiters to oldTID
	end tell
	return theItems
end SplitString
