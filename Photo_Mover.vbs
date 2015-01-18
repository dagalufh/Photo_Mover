'################################################
'########## Created by Mikael Aspehed (dagalufh) 		##########
'########## https://github.com/dagalufh/Photo_Mover  	##########
'########## Current version: 1.0.0.8 								##########
'################################################

' Define the global objects needed
Set objShell = CreateObject ("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")

' Define some defaults, these can be changed but user will get prompted about it when starting script.
Log_Successful_Move = array("")
Log_Successful_CreateFolder = array("")
Log_Failed_Move = array("")
Log_Failed_CreateFolder = array("")
Log_Ignored_Files = array("")

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
LogeFileName = DatePart("yyyy",Now()) & "" & Right(String(1,"0") & DatePart("m",Now()), 2) & "" & Right(String(1,"0") & DatePart("d",Now()),2) & "-" & Right(String(1,"0") & DatePart("h",Now()),2) & "" & Right(String(1,"0") & DatePart("n",Now()),2) & "" & Right(String(1,"0") & DatePart("s",Now()),2)
Set LogFile = fso.CreateTextFile(SourceFolder & "\log-" & LogeFileName & ".txt", True, True)
LogFile.WriteLine "Script was started at " & Now  

' Confirm start
RequestUserConfirmation

' Initial call to start the process with the root source.
SourceDirectory SourceFolder


Sub SourceDirectory (path)
	Set objFolder = objShell.Namespace(path)
	For Each strFileName In objFolder.Items
		' Defaults
		PreviouslyFailedCreating = False
		
		' Check if it is a folder, if so, call the SourceDirectory again to search that subfolder.
		if fso.FolderExists(path & "\" & strFileName) then
			SourceDirectory  path & "\" & strFileName
		else
			
			' Check if there is anything in the number 12 of extended properties. This is where DateTaken is stored.
			if ( (Len(objFolder.GetDetailsOf(strFileName, 12)) > 0) and ( (InStr(objFolder.GetDetailsOf(strFileName, 2), "JPEG") > 0) or (InStr(objFolder.GetDetailsOf(strFileName, 2), "JPG") > 0) ) ) then
				
				' Remove the time from the field as it's only the date we are interested in.
				DateTaken = Split(Mid(objFolder.GetDetailsOf(strFileName, 12), 1, InStr(objFolder.GetDetailsOf(strFileName, 12), " ")-1), "-")
				
				' Replace the keywords in the target path with the dates from the current photo
				TargetFolder_Temp = replace(TargetFolder, "Year",Right(DateTaken(0),Len(DateTaken(0))-1),1,-1, 1)
				TargetFolder_Temp = replace(TargetFolder_Temp, "Month",Right(DateTaken(1),Len(DateTaken(1))-1),1,-1, 1)
				TargetFolder_Temp = replace(TargetFolder_Temp, "Day",Right(DateTaken(2),Len(DateTaken(2))-1),1,-1, 1)
				
				
				' Check if current path has failed previously				
				for each Failed_CreateFolder in Log_Failed_CreateFolder
					
					' if the current path to be created or used has failed to be created earlier, no need to try again.
					if (InStr(Failed_CreateFolder, TargetFolder_Temp)>0) then
						PreviouslyFailedCreating = True
					end if
					
				next	
				
				' If the folder had not previously been failed, we can continue.
				if (PreviouslyFailedCreating = False) then
					' Attempt to create the target folder.
					CreateTargetFolder TargetFolder_Temp
					
					if (fso.FolderExists(TargetFolder_Temp)) then
						if (fso.FileExists(TargetFolder_Temp & "\" & strFileName)) then

							ReDim Preserve Log_Failed_Move(UBound(Log_Failed_Move)+1)
							Log_Failed_Move(UBound(Log_Failed_Move)) = Now & " | Source: [" & path & "\\" & strFileName & "] Target: [" & TargetFolder_Temp & "\" & strFileName & "] Photo already exists in target." 
						else
							On Error Resume Next
							' Attempt to move the file.
							fso.MoveFile path & "\\" & strFileName, TargetFolder_Temp & "\\"	
							
							' If there was an error number reported, assume something went wrong.
							if (Err.Number <> 0) then
								ReDim Preserve Log_Failed_Move(UBound(Log_Failed_Move)+1)
								Log_Failed_Move(UBound(Log_Failed_Move)) = Now & " | Source: [" & path & "\\" & strFileName & "] Target: [" & TargetFolder_Temp & "\" & strFileName & "] Move failed with error number: " & err.number & " and description: " & err.description 
								Err.Clear
							else
								ReDim Preserve Log_Successful_Move(UBound(Log_Successful_Move)+1)
								Log_Successful_Move(UBound(Log_Successful_Move)) = Now & " | Source: [" & path & "\\" & strFileName & "] Target: [" & TargetFolder_Temp & "\" & strFileName & "]" 
							end if
							
						end if
					else
						ReDim Preserve Log_Ignored_Files(UBound(Log_Ignored_Files)+1)
						Log_Ignored_Files(UBound(Log_Ignored_Files)) = Now & " | Source: [" & path & "\\" & strFileName & "] Target: [" & TargetFolder_Temp & "\" & strFileName & "] Target folder failed to be created."
					end if
				else
					ReDim Preserve Log_Ignored_Files(UBound(Log_Ignored_Files)+1)
					Log_Ignored_Files(UBound(Log_Ignored_Files)) = Now & " | Source: [" & path & "\\" & strFileName & "] Target: [" & TargetFolder_Temp & "\" & strFileName & "] Target folder failed to be created previously."	
				end if
				
			else 
				ReDim Preserve Log_Ignored_Files(UBound(Log_Ignored_Files)+1)
				Log_Ignored_Files(UBound(Log_Ignored_Files)) = Now & " | Ignored file: [" & path & "\" & strFileName & "] Filetype: [" & objFolder.GetDetailsOf(strFileName, 2) & "] and date taken value of: [" & objFolder.GetDetailsOf(strFileName, 12) & "]"
			End if
		End if
	Next
End Sub

' Checks to see if the variable holding the input is empty. If so, we treat it as user cancelled and aborts the scripts.
Sub CheckInput (InputValue)
	
	if IsEmpty(InputValue) or InputValue = "2" then
		if not IsEmpty(LogFile) then
			Output_To_LogFile
		end if
		
		WScript.Echo "Aborting and terminating script"
		WScript.Quit
	end if
	
End Sub

' Creates the target folders if needed
Sub CreateTargetFolder (path)
	
	If Not (fso.FolderExists(path)) Then
		CreateTargetFolder fso.GetParentFolderName(path)
		
		On Error Resume Next
		' Attempt to create the folder
		fso.CreateFolder(path)
		
		' If there was an error number reported, assume something went wrong.
		if (Err.Number <> 0) then
			ReDim Preserve Log_Failed_CreateFolder(UBound(Log_Failed_CreateFolder)+1)
			Log_Failed_CreateFolder(UBound(Log_Failed_CreateFolder)) = Now & " | Failed creating folder: [" & path & "] with error number: " & err.number & " and description: " & err.description 
			Err.Clear
		else
			ReDim Preserve Log_Successful_CreateFolder(UBound(Log_Successful_CreateFolder)+1)
			Log_Successful_CreateFolder(UBound(Log_Successful_CreateFolder)) = Now & " | Created folder: [" & path & "]"
		end if
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

		'Check that the drive letter entered by the user is valid and exists
		if fso.FolderExists(mid(TargetFolder,1,InStr(TargetFolder, "\"))) then
			TargetFolder_Valid = true
		else
			TargetFolder_ErrorMessage = vbNewLine & vbNewLine & "Error: The drive (" & mid(TargetFolder,1,3) & ") does not exist in the path you entered. Please select a drive that's available."
			TargetFolder_DefaultSource = TargetFolder
			TargetFolder_Valid = false
		end if
		
		'Check that the target path does not contain the source path. As this will cause a loop.
		if (InStr(1,TargetFolder,SourceFolder & "\",1) > 0) then
			TargetFolder_ErrorMessage = vbNewLine & vbNewLine & "Error: The target (" & TargetFolder & ") can't contain the same path as the source ( " & SourceFolder & ")."
			TargetFolder_DefaultSource = TargetFolder
			TargetFolder_Valid = false
		else
			TargetFolder_Valid = true
		end if
		
	loop

End Sub

Sub RequestUserConfirmation
	
	' Ask the user to confirm previously entered information before continuing.
	Confirmation = MsgBox("Please verify that these paths are correct:" & vbNewLine & vbNewLine & "Source Path: " & SourceFolder & vbNewLine & "Target path: " & TargetFolder & vbNewLine & "All the photos from the source including subdirectories will be moved to the target folder" & vbNewLine & "Depending on the amount of photos this can take some time. A message will popup when done.", vbOKCancel, "Photo Mover - Verify Paths")
	CheckInput Confirmation

End Sub

Sub Output_To_LogFile
	
	' First, output some statistics
	LogFile.WriteLine "Statistics:"
	LogFile.WriteLine "Successful Moves: " & UBound(Log_Successful_Move)
	LogFile.WriteLine "Successful Creation of folders: " & UBound(Log_Successful_CreateFolder)
	LogFile.WriteLine "Failed Moves: " & UBound(Log_Failed_Move)
	LogFile.WriteLine "Failed Creation of folders: " & UBound(Log_Failed_CreateFolder)
	LogFile.WriteLine "Ignored Files: " & UBound(Log_Ignored_Files)
	LogFile.WriteLine ""
	
	LogFile.WriteLine "----------------------------------------------------------------------------"
	LogFile.WriteLine "Successful Moves:"
	for each Successful_Moves in Log_Successful_Move
		' Output successful moves.
			LogFile.WriteLine Successful_Moves
	next
	
	LogFile.WriteLine "----------------------------------------------------------------------------"
	LogFile.WriteLine "Successful Creation of folders:"
	for each Successful_CreateFolder in Log_Successful_CreateFolder
		' Output successful creations of folders.
		LogFile.WriteLine Successful_CreateFolder
	next

	LogFile.WriteLine "----------------------------------------------------------------------------"
	LogFile.WriteLine "Failed Moves:"
	for each Failed_Moves in Log_Failed_Move
		' Output failed moves.
		LogFile.WriteLine Failed_Moves
	next

	LogFile.WriteLine "----------------------------------------------------------------------------"
	LogFile.WriteLine "Failed Creation of folders:"
	for each Failed_CreateFolder in Log_Failed_CreateFolder
		' Output failed creations of folders.
		LogFile.WriteLine Failed_CreateFolder
	next	

	LogFile.WriteLine "----------------------------------------------------------------------------"
	LogFile.WriteLine "Ignored Files:"
	for each Ignored_Files in Log_Ignored_Files
		' Output ignored files.
		LogFile.WriteLine Ignored_Files
	next	
	LogFile.WriteLine "----------------------------------------------------------------------------"
	LogFile.WriteLine ""
	LogFile.WriteLine "Script was completed at " & Now 
End Sub

' Flush out the logs
Output_To_LogFile

' Notify the user that we reached the end.
WScript.Echo "Complete. See logfile for more information: " & SourceFolder & "\log-" & LogeFileName & ".txt"