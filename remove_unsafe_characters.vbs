' Name: remove_unsafe_characters.vbs
' Author: Darsh
' Version: 0.1.0
' Updated: 11.02.2013
' Target Files: files with extensions .jpg, .gif, .png, .pdf
' Description: This script removes unsafe characters ((,),&,$,!,@,#,%,^,{,},[,],',`,~,;,+,= and spaces)in file names

Dim objFso, startFolder, extRe, unsafeCharsRe, numberOfFiles, numberOfFolders, answer, prefix

Set objFso = CreateObject("Scripting.FileSystemObject")
Set extRe = new regexp
Set unsafeCharsRe = new regexp
Set startFolder = objFso.GetFolder(".")

extRe.IgnoreCase = true
extRe.Pattern = "jpg|gif|png|pdf"

unsafeCharsRe.IgnoreCase = true
unsafeCharsRe.Global = true
unsafeCharsRe.Pattern = "[^A-Za-z0-9.]"

numberOfFiles = 0
numberOfFolders = 1
prefix = "file_"


Function removeChars(folder)
	Dim subFolders, newFolderName
	
	Set subFolders = folder.SubFolders

	If (numberOfFolders = 1) Then
		Call writeToLog("Entering current directory")
	End If
	Call removeCharsInFiles(folder.Files)
	
	
	For Each subfolder in subFolders
		newFolderName = unsafeCharsRe.Replace(subfolder.Name,"")
		
		If (newFolderName<>subfolder.Name) Then
			Call writeToLog(subfolder.Name & " >> " & newFolderName)
			subfolder.Name = newFolderName
		End If
		
		Call writeToLog("Entering " & subfolder.path)
		numberOfFolders = numberOfFolders + 1
		removeChars(subfolder)
	Next
	
End Function

Function removeCharsInFiles(files)
	Dim currFileName, ext
	
	For Each File In files
		currFileName = File.Name
		ext = objFso.GetExtensionName(currFileName)
		
		If (extRe.Test(ext)) Then
			
			currFileName = unsafeCharsRe.Replace(currFileName,"")
			
			If (currFileName<>File.Name) Then
			
				numberOfFiles = numberOfFiles + 1
				
				If (currFileName = "." & ext) Then
					currFileName = prefix & numberOfFiles & "." & ext
				End If
				Call writeToLog(File.Name & " >> " & currFileName)
				File.Move(File.ParentFolder+"\"+currFileName)
				
			End If
			
		End If
		
	Next
End Function

Function writeToLog(logthis)
	sLogFileName = "removed_unsafe_characters.log"
	
	Set logOutput = objFso.OpenTextFile(sLogFileName, 8, True)
	
	logOutput.WriteLine(cstr(Now) + " -" + vbTab + logthis)
	logOutput.Close
	
	Set logOutput = Nothing
End Function

answer = MsgBox("Are you sure that you want to run this script?", 4, "Please confirm")

If answer = 6 Then
	Call removeChars(startFolder)
	MsgBox("Summary :" & chr(13) & chr(13) & "No. of folders checked : "& numberOfFolders & chr(13) & "No. of files affected : " & numberOfFiles)
End If