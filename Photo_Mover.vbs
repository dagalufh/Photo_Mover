'#############################################
'########## Created by Mikael Aspehed (dagalufh) ##########
'#############################################


' Define some defaults, these can be changed but user will get prompted about it when starting script.
SourceFolder = InputBox("Enter source directory.", "Photo Mover", "C:\Temp")
TargetFolderRoot = InputBox("Enter target root directory.", "Photo Mover", "C:\Temp")
TargetFolderAppend = InputBox("Enter target subdirectory: " & vbNewLine & vbNewLine & "Valid keywords are: year, month, day. They will be replaced by date taken for the photo." & vbNewLine & vbNewLine & "Root folder is:  [" & TargetFolderRoot & "]","Photo Mover","\Year\Month\Day")
LogeFileName = DatePart("yyyy",Now()) & "" & DatePart("m",Now()) & "" & DatePart("d",Now()) & "-" & DatePart("h",Now()) & "" & DatePart("n",Now()) & "" & DatePart("s",Now())

Set objShell = CreateObject ("Shell.Application")
Set objFolder = objShell.Namespace(SourceFolder)
Set fso = CreateObject("Scripting.FileSystemObject")
Set LogFile = fso.CreateTextFile(SourceFolder & "\log-" & LogeFileName & ".txt", True, True)
LogFile.WriteLine Now & " | Script started."  


For Each strFileName In objFolder.Items
	
	' Check if there is anything in the number 12 of extended properties. This is where DateTaken is stored.
	if (Len(objFolder.GetDetailsOf(strFileName, 12)) > 0) then
		
		' Remove the time from the field as it's only the date we are interested in.
		DateTaken = Split(Mid(objFolder.GetDetailsOf(strFileName, 12), 1, InStr(objFolder.GetDetailsOf(strFileName, 12), " ")-1), "-")
		
		TargetFolderAppend_Temp = replace(TargetFolderAppend, "Year",DateTaken(0),1,-1, 1)
		TargetFolderAppend_Temp = replace(TargetFolderAppend_Temp, "Month",DateTaken(1),1,-1, 1)
		TargetFolderAppend_Temp = replace(TargetFolderAppend_Temp, "Day",DateTaken(2),1,-1, 1)

		CreateTargetFolder TargetFolderRoot & TargetFolderAppend_Temp
		
		if (fso.FileExists(TargetFolderRoot & TargetFolderAppend_Temp & "\" & strFileName)) then

			LogFile.WriteLine Now & " | Source: [" & SourceFolder & "\\" & strFileName & "] Target: [" & TargetFolderRoot & TargetFolderAppend_Temp & "\" & strFileName & "] Photo already exists in target." 
		else
		
			LogFile.WriteLine Now & " | Source: [" & SourceFolder & "\\" & strFileName & "] Target: [" & TargetFolderRoot & TargetFolderAppend_Temp & "\" & strFileName & "] Moving source to target." 
			fso.MoveFile SourceFolder & "\\" & strFileName, TargetFolderRoot & TargetFolderAppend_Temp & "\\"
			
		end if
		
	End if
Next

Sub CreateTargetFolder (path)
	If Not (fso.FolderExists(path)) Then
		CreateTargetFolder fso.GetParentFolderName(path)
		LogFile.WriteLine Now & " | Creating Folder: [" & path & "]"
		fso.CreateFolder(path)
	end if
End Sub
LogFile.WriteLine Now & " | Script ended."
WScript.Echo "Complete. See logfile for more information: " & SourceFolder & "\log-" & LogeFileName & ".txt"