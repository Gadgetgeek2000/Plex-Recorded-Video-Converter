'*********************************************************************
'Geoff's File Object
'10/13/2016 Geoff Faulkner
'*********************************************************************

'dim objTestFile
'set objTestFile = new File
'objTestFile.name = "c:\users\geoff\desktop\test.txt"


'msgbox objTestFile.exists
'msgbox objTestFile.extension
'objTestFile.writefile "test write" & vbCRLF
'msgbox objTestFile.read
'msgbox objTestFile.drive
'msgbox objTestFile.attributes
'msgbox objTestFile.readonly
'objTestFile.readonly = true
'objTestFile.rename "test1.txt"
'msgbox objTestFile.path
'msgbox objTestFile.name
'msgbox objTestFile.shortname
'objTestFile.delete
'CONST Logging = True
'dim objLogFile : Set objLogFile = new File 'object for logging
'objLogFile.name = GetScriptDirectory & "execution.log"
'if Logging then objLogFile.writelog "write this string to the log"

Class File
    Private strFileName
    Public Property Let name(inputname)
        strFileName = inputname
    End Property
    Public Property Get name()
        If Me.exists Then
            Dim objFS, objFile
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFile = objFS.GetFile(strFileName)
            name = objFile.name
        Else
            name = False
        End If
    End Property
    Public Property Get size()
        If Me.exists Then
            Dim objFS, objFile
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFile = objFS.GetFile(strFileName)
            size = Round(objFile.size / 1024, 2)
        End If
    End Property
    Public Property Get extension()
        extension = Right(Me.name, Len(Me.name) - InStrRev(Me.name, "."))
    End Property
    Public Property Get drive()
        If Me.exists Then
            Dim objFS, objFile
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFile = objFS.GetFile(strFileName)
            drive = objFile.drive
        End If
    End Property
    Public Property Get path()
        If Me.exists Then
            Dim objFS, objFile
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFile = objFS.GetFile(strFileName)
            path = objFile.ParentFolder
        Else
            path = False
        End If
    End Property
    Public Property Get shortName()
        If Me.exists Then
            Dim objFS, objFile
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFile = objFS.GetFile(strFileName)
            shortName = objFile.shortName
        Else
            shortName = False
        End If
    End Property
    Public Property Get shortPath()
        If Me.exists Then
            Dim objFS, objFile
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFile = objFS.GetFile(strFileName)
            shortPath = objFile.shortPath
        Else
            shortPath = False
        End If
    End Property
    Public Property Get exists()
        Dim objFS
        Set objFS = CreateObject("Scripting.FileSystemObject")
        If objFS.FileExists(strFileName) Then
            exists = True
        Else
            exists = False
        End If
    End Property
    Public Property Get dateCreated()
        If Me.exists Then
            Dim objFS, objFile
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFile = objFS.GetFile(strFileName)
            dateCreated = objFile.dateCreated
        Else
            dateCreated = False
        End If
    End Property
    Public Property Get dateModified()
        If Me.exists Then
            Dim objFS, objFile
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFile = objFS.GetFile(strFileName)
            dateModified = objFile.DateLastModified
        Else
            dateModified = False
        End If
    End Property
    Public Property Get dateAccessed()
        If Me.exists Then
            Dim objFS, objFile
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFile = objFS.GetFile(strFileName)
            dateAccessed = objFile.DateLastAccessed
        Else
            dateAccessed = False
        End If
    End Property
    Public Property Get readOnly()
        readOnly = False
        If Me.exists Then
            Dim objFS, objFile, attributes
            attributes = Me.attributes
            If GetBit(attributes, 1) Then readOnly = True
        End If
    End Property
	Public Property Get locked()
		locked = True
		if me.exists Then
			if me.Move(me.path) Then
				locked = False
			end if
		End if
	End Property
    Public Property Let readOnly(value)
        If Me.exists Then
            Dim objFS, objFile, attributes
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFile = objFS.GetFile(strFileName)
            If value = True Then
                    objFile.attributes = SetBit(Me.attributes, 1, 1)
            Else
                    objFile.attributes = SetBit(Me.attributes, 1, 0)
            End If
        End If
    End Property
    Public Property Get hidden()
        hidden = False
        If Me.exists Then
            Dim objFS, objFile, attributes
            attributes = Me.attributes
            If GetBit(attributes, 2) Then hidden = True
        End If
    End Property
    Public Property Let hidden(value)
        If Me.exists Then
            Dim objFS, objFile, attributes
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFile = objFS.GetFile(strFileName)
            If value = True Then
                objFile.attributes = SetBit(Me.attributes, 2, 1)
            Else
                objFile.attributes = SetBit(Me.attributes, 2, 0)
            End If
        End If
    End Property
    Public Property Get system()
        system = False
        If Me.exists Then
            Dim objFS, objFile, attributes
            attributes = Me.attributes
            If GetBit(attributes, 3) Then system = True
        End If
    End Property
    Public Property Get archive()
        archive = False
        If Me.exists Then
            Dim objFS, objFile, attributes
            attributes = Me.attributes
            If GetBit(attributes, 6) Then archive = True
        End If
    End Property
    Public Property Let archive(value)
        If Me.exists Then
            Dim objFS, objFile, attributes
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFile = objFS.GetFile(strFileName)
            If value = True Then
                    objFile.attributes = SetBit(Me.attributes, 6, 1)
            Else
                    objFile.attributes = SetBit(Me.attributes, 6, 0)
            End If
        End If
    End Property
    Public Property Get attributes()
        '-1 - Doesn't exist
        '0 - Normal
        '1 - ReadOnly   (bit 1)
        '2 - Hidden (bit 2)
        '4 - System (bit 3)
        '8 - Volume (bit 4)
        '16 - Directory (bit 5)
        '32 - Archive   (bit 6)
        '1024 - Alias   (bit 9)
        '2048 - Compressed (bit 10)
        If Me.exists Then
            Dim objFS, objFile
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFile = objFS.GetFile(strFileName)
            attributes = objFile.attributes
        Else
            attributes = -1
        End If
    End Property
    Public Function read()
        'Returns the contents of the specified objFile
        Const ForReading = 1, TristateUseDefault = -2
        If Me.exists Then
            On Error Resume Next
            Dim objFS, objFile
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFile = objFS.OpenTextFile(strFileName, ForReading, False, TristateUseDefault)
            If objFile.AtEndOfStream = False Then read = objFile.ReadAll
            objFile.Close
            If Err.Number <> 0 Then read = False
            On Error GoTo 0
        End If
    End Function
    Public Function append(ByVal strToWrite)
        append = False
        Const ForAppending = 8, TristateUseDefault = -2
        Dim objFS, objFile
        Set objFS = CreateObject("Scripting.FileSystemObject")
        Set objFile = objFS.OpenTextFile(strFileName, ForAppending, True, TristateUseDefault)
        objFile.Write strToWrite
        objFile.Close
        append = True
    End Function
    Public Function writeLog(strTextToAdd)
        If strTextToAdd <> "" Then Me.append Now & vbTab & strTextToAdd & vbCrLf
    End Function
    Public Function extractZIP(ByVal strPathToExtract)
        Dim objFS
        Dim objShell
        Dim objFilesInZip
		If me.exists Then
            Set objFS = CreateObject("Scripting.FileSystemObject")
            'if the folder doesn't exist, create it
            If Not objFS.FolderExists(strPathToExtract) Then
                objFS.CreateFolder strPathToExtract
            End If
            'Extract
            Set objShell = CreateObject("Shell.Application")
            Set objFilesInZip = objShell.NameSpace(strFileName).items
            objShell.NameSpace(strPathToExtract).CopyHere (objFilesInZip)
        Else
            MsgBox "File not found"
            extractZIP = False
        End If
    End Function
	Public Function WriteFile(ByVal strToWrite)
        WriteFile = False
        Const ForWriting = 2, TristateUseDefault = -2
        Dim objFS, objFile
        Set objFS = CreateObject("Scripting.FileSystemObject")
        On Error Resume Next
        Set objFile = objFS.OpenTextFile(strFileName, ForWriting, True, TristateUseDefault)
        objFile.Write strToWrite
        objFile.Close
        If Err.Number = 0 Then WriteFile = True
        On Error GoTo 0
    End Function
    Public Function rename(ByVal strNewName)
        rename = False
        If Me.exists Then
            Dim objFS
            Set objFS = CreateObject("Scripting.FileSystemObject")
            On Error Resume Next
            objFS.movefile strFileName, strNewName
            If Err = 0 Then 
		rename = True
		strFileName = strNewName
            end if
            On Error GoTo 0
        End If
    End Function
    Public Function move(ByVal strTargetFilePath)
        If Right(strTargetFilePath, 1) <> "\" Then strTargetFilePath = strTargetFilePath & "\"
        move = False
        If Me.exists Then
            Dim objFS, objFile
            Set objFS = CreateObject("Scripting.FileSystemObject")
            On Error Resume Next
            objFS.MoveFile strFileName, strTargetFilePath
            If Err.Number = 0 Then move = True
            On Error GoTo 0
        End If
    End Function
	Public Function moveOverwrite(ByVal strTargetFilePath)
        If Right(strTargetFilePath, 1) <> "\" Then strTargetFilePath = strTargetFilePath & "\"
        moveOverwrite = False
        If Me.exists Then
            Dim objFS, objFile
            Set objFS = CreateObject("Scripting.FileSystemObject")
            On Error Resume Next
            objFS.copyFile strFileName, strTargetFilePath, True
			objFS.deleteFile strFileName
            If Err.Number = 0 Then moveOverwrite = True
            On Error GoTo 0
        End If
    End Function
    Public Function delete()
        Dim objFS
        Set objFS = CreateObject("Scripting.FileSystemObject")
        On Error Resume Next
        If Me.exists = True Then objFS.DeleteFile (strFileName)
        If Me.exists = False Then
            delete = True
        Else
            delete = False
        End If
        On Error GoTo 0
    End Function
	Public Function copy(strDestinationFileName)
        Dim objFS
        Set objFS = CreateObject("Scripting.FileSystemObject")
        On Error Resume Next
        If Me.exists = True Then objFS.copyfile strFileName, strDestinationFileName
        On Error GoTo 0
    End function
    Private Function GetBit(lngValue, BitNum)
        Dim BitMask
        If BitNum < 32 Then BitMask = 2 ^ (BitNum - 1) Else BitMask = "&H80000000"
        GetBit = CBool(lngValue And BitMask)
    End Function
    Private Function SetBit(lngValue, BitNum, NewValue)
        Dim BitMask
        If BitNum < 32 Then BitMask = 2 ^ (BitNum - 1) Else BitMask = "&H80000000"
        If NewValue Then
            SetBit = lngValue Or BitMask
        Else
            BitMask = Not BitMask
            SetBit = lngValue And BitMask
        End If
    End Function
	Public Function PrepareFilename(strToFormat)
		'This function replaces illegal characters in the filename with acceptable alternatives.
		'Accepts a string and returns a formatted scring.
		strToFormat = replace(strToFormat, ":", "")
		strToFormat = replace(strToFormat, "\", "-")
		strToFormat = replace(strToFormat, "/", "-")
		strToFormat = replace(strToFormat, "*", "")
		strToFormat = replace(strToFormat, "?", "")
		strToFormat = replace(strToFormat, "|", "")
		strToFormat = replace(strToFormat, "", "'")
		strToFormat = replace(strToFormat, "", "'")
		strToFormat = replace(strToFormat, """", "'")
		strToFormat = replace(strToFormat, "<", "(")
		strToFormat = replace(strToFormat, ">", ")")
		PrepareFilename = strToFormat	
    End Function
End Class
