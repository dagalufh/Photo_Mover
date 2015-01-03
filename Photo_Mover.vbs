'################################################
'########## Created by Mikael Aspehed (dagalufh) 		##########
'########## https://github.com/dagalufh/Photo_Mover  	##########
'########## Current version: 1.0.0.2 								##########
'################################################

' Define the global objects needed
Set objShell = CreateObject ("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")

' Define some defaults, these can be changed but user will get prompted about it when starting script.
' Set some defaults for the source
SourceFolder_Valid = false
SourceFolder_ErrorMessage = ""
SourceFolder_DefaultSource = "C:\Temp"
SourceFolder = ""

' Set some defaults for the target
TargetFolder_Valid = false
TargetFolder_ErrorMessage = ""
TargetFolder_DefaultSource = "C:\Photos\Year\Month\Day"
TargetFolder = ""

' Request the user to verify paths.
RequestInputSource
RequestInputTarget

' Define the name of the log and create it.
LogeFileName = DatePart("yyyy",Now()) & "" & DatePart("m",Now()) & "" & DatePart("d",Now()) & "-" & DatePart("h",Now()) & "" & DatePart("n",Now()) & "" & DatePart("s",Now())
Set LogFile = fso.CreateTextFile(SourceFolder & "\log-" & LogeFileName & ".txt", True, True)
LogFile.WriteLine Now & " | Script started."  

' Confirm start
RequestUserConfirmation

' Initial call to start the process with the root source.
SourceDirectory SourceFolder


Sub SourceDirectory (path)
	Set objFolder = objShell.Namespace(path)
	For Each strFileName In objFolder.Items
		' Check if it is a folder, if so, call the SourceDirectory again to search that subfolder.
		if fso.FolderExists(path & "\" & strFileName) then
			SourceDirectory  path & "\" & strFileName
		else
		
			' Check if there is anything in the number 12 of extended properties. This is where DateTaken is stored.
			if ( (Len(objFolder.GetDetailsOf(strFileName, 12)) > 0) and (objFolder.GetDetailsOf(strFileName, 2) = "JPG File") ) then
				
				' Remove the time from the field as it's only the date we are interested in.
				DateTaken = Split(Mid(objFolder.GetDetailsOf(strFileName, 12), 1, InStr(objFolder.GetDetailsOf(strFileName, 12), " ")-1), "-")
				
				' Replace the keywords in the target path with the dates from the current photo
				TargetFolder_Temp = replace(TargetFolder, "Year",DateTaken(0),1,-1, 1)
				TargetFolder_Temp = replace(TargetFolder_Temp, "Month",DateTaken(1),1,-1, 1)
				TargetFolder_Temp = replace(TargetFolder_Temp, "Day",DateTaken(2),1,-1, 1)
				
				' Check if the target needs to be created
				CreateTargetFolder TargetFolder_Temp
				
				if (fso.FileExists(TargetFolder_Temp & "\" & strFileName)) then

					LogFile.WriteLine Now & " | Source: [" & path & "\\" & strFileName & "] Target: [" & TargetFolder_Temp & "\" & strFileName & "] Photo already exists in target." 
					
				else
				
					LogFile.WriteLine Now & " | Source: [" & path & "\\" & strFileName & "] Target: [" & TargetFolder_Temp & "\" & strFileName & "] Moving source to target." 
					fso.MoveFile path & "\\" & strFileName, TargetFolder_Temp & "\\"
					
				End if
				
			End if
		End if
	Next
End Sub

' Checks to see if the variable holding the input is empty. If so, we treat it as user cancelled and aborts the scripts.
Sub CheckInput (InputValue)
	
	if IsEmpty(InputValue) or InputValue = "2" then
		WScript.Echo "Aborting and terminating script"
		WScript.Quit
	end if
	
End Sub

' Creates the target folders if needed
Sub CreateTargetFolder (path)
	
	If Not (fso.FolderExists(path)) Then
		CreateTargetFolder fso.GetParentFolderName(path)
		LogFile.WriteLine Now & " | Creating Folder: [" & path & "]"
		fso.CreateFolder(path)
	end if
	
End Sub

Sub RequestInputSource
	
	' Request a source path until we get a valid or user presses cancel
	do until SourceFolder_Valid = true
		SourceFolder = InputBox("Enter source directory." & SourceFolder_ErrorMessage, "Photo Mover - Source Folder", SourceFolder_DefaultSource)
		CheckInput SourceFolder	
		if fso.FolderExists(SourceFolder) then
			SourceFolder_Valid = true
		else
			SourceFolder_ErrorMessage = vbNewLine & vbNewLine & "Error: Source folder does not exist. Please check the spelling."
			SourceFolder_DefaultSource = SourceFolder
		end if
	loop
	
End Sub

Sub RequestInputTarget
		
	' Request a target path until we get a valid or user presses cancel
	do until TargetFolder_Valid = true
		TargetFolder_Temp = replace(TargetFolder_DefaultSource, "Year","2014",1,-1, 1)
		TargetFolder_Temp = replace(TargetFolder_Temp, "Month","10",1,-1, 1)
		TargetFolder_Temp = replace(TargetFolder_Temp, "Day","25",1,-1, 1)
		TargetFolder = InputBox("Enter target directory."& vbNewLine & vbNewLine & "Valid keywords are: year, month, day. They will be replaced by date taken for the photo. Example for photo taken 2014-10-25 will be: " & vbNewLine & TargetFolder_Temp & TargetFolder_ErrorMessage, "Photo Mover - Target Folder", TargetFolder_DefaultSource)
		CheckInput TargetFolder	

		'Check that the drive letter entered by the user is valid
		if fso.FolderExists(mid(TargetFolder,1,InStr(TargetFolder, "\"))) then
			TargetFolder_Valid = true
		else
			TargetFolder_ErrorMessage = vbNewLine & vbNewLine & "Error: The drive (" & mid(TargetFolder,1,3) & ") does not exist in the path you entered. Please select a drive that's available."
			TargetFolder_DefaultSource = TargetFolder
		end if
	loop

End Sub

Sub RequestUserConfirmation
	
	' Ask the user to confirm previously entered information before continuing.
	Confirmation = MsgBox("Please verify that these paths are correct:" & vbNewLine & vbNewLine & "Source Path: " & SourceFolder & vbNewLine & "Target path: " & TargetFolder & vbNewLine & "All the photos from the source including subdirectories will be moved to the target folder" & vbNewLine & "Depending on the amount of photos this can take some time. A message will popup when done.", vbOKCancel, "Photo Mover - Verify Paths")
	CheckInput Confirmation

End Sub

' Notify the user that we reached the end.
LogFile.WriteLine Now & " | Script ended."
WScript.Echo "Complete. See logfile for more information: " & SourceFolder & "\log-" & LogeFileName & ".txt"