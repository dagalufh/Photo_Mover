SourceFolder = "F:\Testar"
TargetFolderRoot = "F:\\Testar"
StartTime = DatePart("yyyy",Now()) & "" & DatePart("m",Now()) & "" & DatePart("d",Now()) & "-" & DatePart("h",Now()) & "" & DatePart("n",Now()) & "" & DatePart("s",Now())

Set objShell = CreateObject ("Shell.Application")
Set objFolder = objShell.Namespace(SourceFolder)
Set fso = CreateObject("Scripting.FileSystemObject")
Set LogFile = fso.OpenTextFile(SourceFolder & "\log-" & StartTime & ".txt", 2, True)
  
LogFile.Write "Script started at: " & Now
Dim DateTaken, Year, Month, Day
For Each strFileName In objFolder.Items
	if (Len(objFolder.GetDetailsOf(strFileName, 12)) > 0) then
		'Wscript.Echo objFolder.GetDetailsOf(strFileName, 0) & vbTab & "["  & objFolder.GetDetailsOf(strFileName, 12) & "]"
		
		DateTaken = Split(Mid(objFolder.GetDetailsOf(strFileName, 12), 1, InStr(objFolder.GetDetailsOf(strFileName, 12), " ")-1), "-")
		Year = DateTaken(0)
		Month = DateTaken(1)
		Day = DateTaken(2)
		TargetFolderAppend = "\\" & Year & "\\" & Month & "\\" & Day	
		
		'Wscript.Echo Year & " - " & Month & " - " & Day
		
		CreateTargetFolder TargetFolderRoot & TargetFolderAppend
		
		if (fso.FileExists(TargetFolderRoot & TargetFolderAppend & "\" & strFileName)) then
			'Wscript.Echo "Photo already exists in folder. [" & TargetFolderRoot & TargetFolderAppend & "\" & strFileName &"]"
		else
			' Wscript.Echo "Photo needs to be moved to the folder.["&TargetFolderRoot & TargetFolderAppend & "\" & strFileName &"]"
			
			'Wscript.Echo SourceFolder & "\\" & strFileName & "," & TargetFolderRoot & TargetFolderAppend & "\\" 
			fso.MoveFile SourceFolder & "\\" & strFileName, TargetFolderRoot & TargetFolderAppend & "\\"
			
		end if
		
	End if
Next

Sub CreateTargetFolder (path)
	
	If (fso.FolderExists(path)) Then
		'Wscript.Echo "Folder already exists! ["& path &"]"
	else
		'Wscript.Echo "Folder does not exists! ["& path &"]"
		CreateTargetFolder fso.GetParentFolderName(path)
		fso.CreateFolder(path)
	end if
	
End Sub

LogFile.Write "Script ended at: " & Now