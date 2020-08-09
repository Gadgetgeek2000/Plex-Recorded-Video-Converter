Option Explicit
'****************************************************************************
'***  Plex Video Converter.vbs                                            ***
'***  Written 10-13-2016  by Geoff Faulkner                               ***
'***  This program parses a folder for .ts files and then launches a      ***
'***  command line (exe, script, etc.) to convert the video               ***
'****************************************************************************

'class libraries
includeFile "clsFile.vbs"
includeFile "clsFolder.vbs"

'Constants -- change these settings to match your configuration
CONST SOURCE_MEDIA_FOLDER = "\\homepvr\d$\Plex Recorded TV\"
CONST LOGFILE = "\\homepvr\d$\Plex Recorded TV\_executionlog.log"
CONST CONVERTER_COMMAND_LINE = """c:\Program Files (x86)\Plex Video Converter\Handbrake\HandBrakeCLI.exe"" -i ""%INPUT_FILE%"" -t 1 --angle 1 -c 1-11 -o ""%OUTPUT_FILE%""  -f mp4 --width %OUTPUT_WIDTH% --height %OUTPUT_HEIGHT% --crop 0:0:6:4 --loose-anamorphic  --modulus 2 -e x264 -q 20 --vfr -a 1 -E av_aac -6 dpl2 -R Auto -B 160 -D 0 --gain 0 --audio-fallback ac3  --encoder-preset=veryfast  --encoder-level=""4.0""  --encoder-profile=main  --verbose=1"
CONST INPUT_FILE_STRING = "%INPUT_FILE%"
CONST OUTPUT_FILE_STRING = "%OUTPUT_FILE%"
CONST OUTPUT_WIDTH_STRING = "%OUTPUT_WIDTH%"
CONST OUTPUT_HEIGHT_STRING = "%OUTPUT_HEIGHT%"
CONST OUTPUT_FOLDER = "\\homepvr\data\video\television\"
CONST OUTPUT_ERROR_FOLDER = "\\homepvr\d$\Plex Recorded TV Conversion Errors\"
CONST OUTPUT_FILE_TYPE = ".MP4"
CONST INPUT_FILE_TYPE = ".TS"
CONST OUTPUT_WIDTH = "1280"
CONST OUTPUT_HEIGHT = "720"
CONST LOGGING = True
CONST PARSE_SUBFOLDERS = True
CONST IGNORE_HIDDEN_SUBFOLDERS = True
CONST DEBUGGING=FALSE

'Ensure the folders exist
dim objFolder
set objFolder = new Folder
objFolder.path = SOURCE_MEDIA_FOLDER
objFolder.Create
objFolder.path = OUTPUT_FOLDER
objFolder.Create
objFolder.path = OUTPUT_ERROR_FOLDER
objFolder.Create
set objFolder = Nothing

'run the routine for the folder
parsefolder SOURCE_MEDIA_FOLDER

Sub parseFolder (strFolderToParse)
if Debugging then msgbox "ParseFolder " & strFolderToParse

	'variables
	dim objFolder, objFile, objTargetFile, strPath, strShow, strSeason, objDoneFile
	dim objLogFile
	dim strDestination, objOutputFolder, strCommandLine
	dim intExitCode

	'setup the Folder object
'	objLogFile.WriteLog "Setting folder object"
	set objFolder = new Folder
	objFolder.path = strFolderToParse
	
	'Ensure the folder exists
	if objFolder.exists Then
		if not (objFolder.hidden AND IGNORE_HIDDEN_SUBFOLDERS) Then
'			objLogFile.WriteLog "Folder exists and is not hidden. Opening the folder"
			objFolder.Open
			
			'parse through the Files
			Do until objFolder.getFileName = ""
				'get file properties
				set objFile = Nothing
				set objFile = new file
				objFile.name = strFolderToParse & "\" & objFolder.getFileName
				
				'if it's a  recorded video file Then
				If ucase(right(objFile.name,len(INPUT_FILE_TYPE))) = INPUT_FILE_TYPE Then

					'Setup the logfile
					set objLogFile = new File
					objLogFile.name = objFile.path & "\" & objFile.name & ".log"
				
					'Check to see if a "done" file exists, meaning it has recorded the same show twice
					set objDoneFile = nothing
					set objDoneFile = new file
					objDoneFile.name = strFolderToParse & "\" & objFolder.getFileName & ".done"
					
					if objDoneFile.Exists then
						if objDoneFile.size < objFile.size then
							'the new recording is larger. It could be a better HD recording
							'delete the done file and the log.
							objDoneFile.delete
							objDoneFile.name = strFolderToParse & "\" & objFolder.getFileName & ".log"
							objDoneFile.delete
							set objDoneFile = nothing
						else
							'the new recording is smaller. Delete the file, and write to the log.  
							objLogFile.WriteLog "deleted a smaller video file, which was a duplicate."
							objFile.delete
						end if
					end if
					

					'Check to see if the log file exists. If it does, skip the file. Otherwise create the logfile
					If not objLogFile.exists Then
						'Check to see if the file is in use
						if not objFile.locked then 
							objLogFile.WriteFile ""  'Create the log
							objLogFile.WriteLog "Processing file " & objFile.name
							objLogFile.WriteLog "Not locked."
							
							'determine the show name and season from the path
if Debugging then msgbox "Outputfolder" & OUTPUT_FOLDER
							strPath = objFile.path
if Debugging then msgbox "Path: " & strPath						
							strPath = replace(strPath, SOURCE_MEDIA_FOLDER, "", 1, -1, 1)
if Debugging then msgbox "Path: " & strPath													
							strShow = left(strPath, instr(strPath, "\") - 1)
if Debugging then msgbox "Show: " & strShow
							strPath = replace(strPath, ucase(strShow & "\"), "", 1, -1, 1)
if Debugging then msgbox "path: " & strPath
							strSeason = strPath
							strSeason = right(strSeason, len(strSeason) - instrRev(strSeason, " "))
							strSeason = "Season " & right("00" & strSeason, 2)
if Debugging then msgbox "season: " & strSeason
							'capitalize each word of the show name.
							strShow = EachWordUpper(strShow)
							
							objLogFile.WriteLog "Show: " & strShow
							objLogFile.WriteLog "Season: " & strSeason
							
							
							
							
							'set the destination Directory
							strDestination = OUTPUT_FOLDER & strShow & "\" & strSeason & "\" & objFile.name
if Debugging then msgbox "strDestination: " & strDestination							
							objLogFile.WriteLog "Destination: " & strDestination
							
							'make the output Folders if they don't exist
							set objOutputFolder = new Folder
							objOutputFolder.Path = OUTPUT_FOLDER & strShow
							objOutputFolder.Create
							objOutputFolder.Path = OUTPUT_FOLDER & strShow & "\" & strSeason
							objOutputFolder.Create
							objLogFile.WriteLog "Created folder " & objOutputFolder.Path
							set objOutputFolder = Nothing
												
							'assemble the command line
							strCommandLine = CONVERTER_COMMAND_LINE
							strCommandLine = replace(strCommandLine, INPUT_FILE_STRING, objFile.path & "\" & objFile.name, 1, -1, 1)
							strCommandLine = replace(strCommandLine, OUTPUT_FILE_STRING, OUTPUT_FOLDER & EachWordUpper(strShow) & "\" & strSeason & "\" & replace(objFile.Name, INPUT_FILE_TYPE, lcase(OUTPUT_FILE_TYPE), 1, -1, 1), 1, -1, 1)
							strCommandLine = replace(strCommandLine, OUTPUT_WIDTH_STRING, OUTPUT_WIDTH, 1, -1, 1)
							strCommandLine = replace(strCommandLine, OUTPUT_HEIGHT_STRING, OUTPUT_HEIGHT, 1, -1, 1)
							objLogFile.WriteLog strCommandLine
							
							'check that the target file doesn't already exist. If it does, move it to the recycle bin (in case it needs to be recovered later
							set objTargetFile = Nothing
							set objTargetFile = new file
							
							'borrow the targetfile object for a moment, please
							'look for a the "do not overwrite" File. If it does then get out of the folder.
							objTargetFile.Name = OUTPUT_FOLDER & strShow & "\" & strSeason & "\" & "_Do not overwrite.txt"
							objLogFile.WriteLog objTargetFile.path & "\" & objTargetFile.name

							if objTargetFile.exists then
								objLogFile.WriteLog "Don't overwrite exists"
								objLogFile.WriteLog "Folder contains do not overwrite instructions. Exiting this folder."
								objFile.rename objFile.name & ".done"
								Exit Do
							Else
								objLogFile.WriteLog "Don't overwrite doesn't exist"
							End If

							'set the target file
							objTargetFile.name = OUTPUT_FOLDER & strShow & "\" & strSeason & "\" & replace(ucase(objFile.Name), INPUT_FILE_TYPE, lcase(OUTPUT_FILE_TYPE), 1, -1, 1)
							if objTargetFile.exists Then
								objLogFile.WriteLog "The target file " & objTargetFile.path & "\" & objTargetFile.name &  " exists."
								
								'check to see if ReadOnly
								if not objTargetFile.ReadOnly then
									objLogFile.WriteLog "The target file is not read-only. Deleting and converting."
									objTargetFile.delete

									'convert the file
									intExitCode=ShellExecute(strCommandLine, 1, true)
								Else
									objLogFile.WriteLog "The target file is read-only. Taking no action to overwrite."
								End If
							Else
								'convert the file
								objLogFile.WriteLog "The target file does not exist. Converting."
								intExitCode = ShellExecute(strCommandLine, 1, true)
								objLogFile.WriteLog "Exited with code " & intExitCode
							end if
														
							if intExitCode = 0 then
								'delete the source if exists at the target. Otherwise move it to the error folder.
								objLogFile.WriteLog "Checking if target " & objTargetFile.name & " exists:"
								if objTargetFile.exists Then
									objLogFile.WriteLog "Exists. Renaming original to .done"
									objFile.rename objFile.name & ".done"
								Else
									objLogFile.WriteLog "Doesn't exist. Renaming original to .targetdidntexist"
									objFile.rename objFile.name & ".targetdidntexist"
								end if
							else
								objFile.rename objFile.name & ".generatedError"
							end if

						else
	'						objLogFile.WriteLog "Locked. Skipping."
						end if
					end if
				
					set objLogFile = nothing				
				end if

				'delete files over 1 week old
				if ucase(right(objFile.name,5)) = ".DONE" then

					'check date
					if Debugging then msgbox objFile.name & ": created " & objFile.DateCreated
					if Debugging then msgbox "7 days ago: " & dateadd("d",now(),-7)
					if Debugging then msgbox "test to delete: " & objFile.dateCreated < dateadd("d",now(),-7)


					if objFile.dateCreated < dateadd("d",now(),-7) then
						'old file, delete it
						objFile.delete
						
						'delete the log file
						set objLogFile = new File
						objLogFile.name = replace(ucase(objFile.path & "\" & objFile.name), ".DONE", "") & ".log"
						objLogFile.delete
						
						set objFile = nothing
						set objLogFile = nothing

					end if
				end if
				
				objFolder.nextFile
			Loop
			
			'parse through the subfolders
			Do until objFolder.getSubFolderName = ""
'				objLogFile.WriteLog "Found subfolder: " & objFolder.getSubFolderName
				if PARSE_SUBFOLDERS Then
					'ignore the .grab folder, which is where new videos are recorded.
					if lcase(objFolder.getSubFolderName) <> ".grab" then
'						objLogFile.WriteLog "Processing folder " & objFolder.path 
						ParseFolder(objFolder.path & "\" & objFolder.getSubFolderName)
					end if
				End If
				objFolder.nextSubFolder
			Loop		
		Else
'			objLogFile.WriteLog "The folder " & objFolder.path & " is hidden. Ignoring."
		End If
	Else
'		objLogFile.WriteLog "The top-level folder does not exist."
	end if

end sub

'********************************************
'Functions
'********************************************
sub includeFile (ByVal strFileName)
	'This function reads in a library / include file and prepares it for execution.
	dim objFS, objFile, strFileContents
	set objFS = createObject ("Scripting.FileSystemObject")
	
	'check if the full path was given. If not, assume the included file exists
	'in the same folder as the script
	if not instr(strFilename,"\") then strFileName = GetScriptDirectory & "\" & strFileName
	set objFile = objFS.OpenTextFile(strFileName)
	strFileContents = objFile.readall()
	objFile.close
	executeGlobal strFileContents
	set objFile = Nothing
	set objFS = nothing
end sub

Function GetScriptDirectory()
	GetScriptDirectory = Left(wscript.scriptfullname, InStr(1, wscript.scriptfullname, wscript.scriptname) - 1)
End Function

Function ShellExecute(ByVal ShellCommandLine, ByVal ShellOption, ByVal WaitOnReturn)
	if ShellOption < 0 or ShellOption > 10 then
		ShellOption = 3 'Activated and Maximized
	end if

	if WaitOnReturn <> True AND WaitOnReturn <> False then
		WaitOnReturn = True
	end if

	if ShellCommandLine <> "" Then
		dim objShell
		Set objShell = CreateObject("WScript.Shell")
		ShellExecute = objShell.Run(ShellCommandLine, ShellOption, WaitOnReturn)
	end if
End Function

Function EachWordUpper(strText)
	Dim arrText, strWord
	arrText = split(strText, " ")
	For each strWord in arrText
		EachWordUpper = EachWordUpper & ucase(left(strWord,1)) & lcase(mid(strWord,2))& " "
	Next
	EachWordUpper = Trim(EachWordUpper)
End Function
