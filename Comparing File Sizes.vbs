'This script will compare the file sizes of files from two different folders.
'It will then output this into a text file, seperated with a tab. (This allows it to be easily imported into Excel)



'ISSUE: if a subfolder is located within a folder provided by the user, the script will not return anything at all.



'Don't be dumB.
option explicit

'Declaring stupid dumb variable because VBS is a dumb baby language.
dim stupidVariable

'File system handles and variables.
dim objFSO, objFile, objFolder
dim objFileOne
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Collection variables. These contain the two file collections.
Dim collectionOne, collectionTwo

'Current file variables. These hold the current files being examined.
Dim fileOne, fileTwo

'This is the basename of a file. Only the file name, no extension.
dim objFileName

'Defining list text file variable and creating the text file.
dim list
Set list = objFSO.OpenTextFile("LIST.txt", 2, True, 0)

'Creating handle to browse folders.
dim objShell
Set objShell = CreateObject("Shell.Application")

'Creating message box to notify user what the script does and the general "how it works".
stupidVariable = MsgBox("This script will compare the file sizes of files within 2 folders and output matches. This list can be easily exported into Excel. Note: This does not include folders, or zip files.", 1, "NOTICE")
'If "Ok" was pressed.
If stupidVariable = 1 Then
	'Creating browser 
	Set objFolder = objShell.BrowseForFolder(0, "Please choose a folder. This will only look at files, not zips or other folders.", 1, 0)


	If Not (objFolder Is Nothing) Then
		'Making this a proper collection.
		Set collectionOne = objFSO.GetFolder(objFolder.Self.path + "\").Files


		Set objFolder = objShell.BrowseForFolder(0, "Please choose a second folder. This will only look at files, not zips or other folders.", 1, 0)
			If Not (objFolder Is Nothing) Then
				'Making this a proper collection too.
				Set collectionTwo = objFSO.GetFolder(objFolder.Self.path + "\").Files
				
				'For every file in the folder, set the file as "fileOne"
				For Each objFile in collectionOne
					Set fileOne = objFile
						
						'For every file in the folder, set it to fileTwo and compare the size of the 2 files.
						For Each objFileOne in collectionTwo
							Set fileTwo = objFileOne
							If fileOne.Size = fileTwo.Size Then
								list.WriteLine(fileOne & vbTab & fileTwo)
							End If
						Next
					
				Next
				
		
		Else
		'If no input is given, announce so and close the script.
		stupidVariable = MsgBox("No input given! Script closing.", 0, "Comparing Failed!")
		WScript.Quit()
		End If
	Else
		'If no input is given, announce so and close the script.
		stupidVariable = MsgBox("No input given! Script closing.", 0, "Comparing Failed!")
		WScript.Quit()
	End If

	'Announcing that everything is done!
	objFile = objFSO.GetFile("LIST.txt")
	MsgBox("Finished! The text file is located at: " & objFile)
Else
	stupidVariable = MsgBox("Script Aborted", 0, "Aborted")
	Wscript.Quit()
End If


'Closing the text file and the script.
list.Close
Wscript.quit()