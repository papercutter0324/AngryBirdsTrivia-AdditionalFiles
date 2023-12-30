on FindConfigDirectory(paramstring)
	set homeDocumentsPath to POSIX path of home folder & "/Documents"
	set configFolder to homeDocumentsPath & "/DYB_AngryBirdsTrivia"
	try
		if not ExistsFolder(configFolder) then
			make new folder at homeDocumentsPath with properties {name:"DYB_AngryBirdsTrivia"}
		end if
	on error
		display dialog "Error creating the folder /DYB_AngryBirdsTrivia in your Documents directory. Please create this folder and reopen the PPT." buttons {"OK"} default button "OK"
	end try
	if not ExistsFolder(configFolder) then
		set configFolder to ""
	end if
	return configFolder
end FindConfigDirectory

on WriteToFile(paramstring)
	set {passedFilePath, passedDataToSave} to SplitString(paramstring, ";")
	if ExistsFile(passedFilePath) then set fileFound to true
	try
		set openedFile to POSIX file passedFilePath
		set openedFile to openedFile as text
		set openedFile to open for access openedFile with write permission
		set dataToWrite to passedDataToSave
		if fileFound then set eof openedFile to 0
		write dataToWrite to openedFile as Çclass utf8È starting at eof
		close access openedFile
		return true
	on error
		try
			close access openedFile
		end try
		return false
	end try
end WriteToFile

on MoveFile(paramstring)
	set {tempConfigFilePath, configFolderPath} to SplitString(paramstring, ";")
	if ExistsFile(tempConfigFilePath) = true then
		set sourceFile to tempConfigFilePath
		set sourceFile to quoted form of POSIX path of sourceFile
		set destinationFolder to quoted form of POSIX path of configFolderPath
		do shell script "mv " & sourceFile & space & destinationFolder
		return true
	else
		return false
	end if
end MoveFile

on LoadConfig(paramstring)
	set fileToOpen to quoted form of POSIX path of paramstring
	if ExistsFile(fileToOpen) = true then
		try
			set fileContent to read file fileToOpen
			set configLines to paragraphs of fileContent
			set inputString to ""
			repeat with aLine in configLines
				set inputString to inputString & line & ";"
			end repeat
			if the last character of inputString is ";" then
				set inputString to text 1 thru -2 of inputString
			end if
		on error
			set inputString to ""
		end try
	end if
	return inputString
end LoadConfig

on SetTempDirectory(paramstring)
	set tempDirectory to path to temporary items
	return tempDirectory
end SetTempDirectory

on ChooseFileToOpen(paramstring)
	set configFolder to paramstring
	set fileToOpen to do shell script "osascript -e 'tell app (path to frontmost application as Unicode text) to set fileToOpen to POSIX path of (choose file with prompt \"Select a Config File\" of type {\"txt\"} default location alias \"" & configFolder & "\")'"
	return fileToOpen
end ChooseFileToOpen

on SplitString(theBigString, fieldSeparator)
	tell AppleScript
		set oldTID to text item delimiters
		set text item delimiters to fieldSeparator
		set theItems to text items of theBigString
		set text item delimiters to oldTID
	end tell
	return theItems
end SplitString

on ExistsFile(filePath)
	tell application "System Events" to return (exists disk item filePath) and class of disk item filePath = file
end ExistsFile

on ExistsFolder(folderPath)
	tell application "System Events" to return (exists disk item folderPath) and class of disk item folderPath = folder
end ExistsFolder

on OpenFolder(folderPath)
	if ExistsFolder(folderPath) = true then
		do shell script "open " & quote & folderPath & quote
	else
		return false
	end if
end OpenFolder

on CopyFile(paramstring)
	set {fieldvalue1, fieldValue2} to SplitString(paramstring, ";")
	if ExistsFile(fieldvalue1) = true and ExistsFolder(fieldValue2) = true then
		set sourceFile to fieldvalue1
		set sourceFile to quoted form of POSIX path of sourceFile
		set destinationFolder to quoted form of POSIX path of (fieldValue2)
		do shell script "cp " & sourceFile & space & destinationFolder
	else
		return false
	end if
end CopyFile

on CopyFolder(paramstring)
	set {fieldvalue1, fieldValue2, fieldValue3} to SplitString(paramstring, ";")
	if ExistsFolder(fieldvalue1) = true and ExistsFolder(fieldValue2) = true and ExistsFolder(fieldValue2 & fieldValue3) = false then
		set Sourcefolder to quoted form of POSIX path of fieldvalue1
		set destinationFolder to quoted form of POSIX path of (fieldValue2 & fieldValue3)
		do shell script "cp -r -n " & Sourcefolder & space & destinationFolder
	else
		return false
	end if
end CopyFolder

on CopyFolderMerge(paramstring)
	set {fieldvalue1, fieldValue2, fieldValue3} to SplitString(paramstring, ";")
	if ExistsFolder(fieldvalue1) = true and ExistsFolder(fieldValue2) = true then
		set Sourcefolder to quoted form of POSIX path of fieldvalue1
		set destinationFolder to quoted form of POSIX path of (fieldValue2 & fieldValue3)
		do shell script "ditto " & Sourcefolder & space & destinationFolder
	else
		return false
	end if
end CopyFolderMerge