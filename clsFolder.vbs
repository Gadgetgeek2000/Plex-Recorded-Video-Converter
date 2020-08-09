'*********************************************************************
'Geoff's Folder Object
'10/13/2016 Geoff Faulkner
'*********************************************************************

'dim objTestFolder
'set objTestFolder = new Folder
'objTestFolder.name = "c:\users\geoff\desktop\testfolder"
'objTestFolder.Create

'msgbox objTestFolder.exists

'objTestFolder.rename "testfolder2"


'msgbox objTestFolder.attributes
'msgbox objTestFolder.readonly
'objTestFolder.readonly = true
'msgbox objTestFolder.path
'msgbox objTestFolder.name
'msgbox objTestFolder.shortname
'objTestFolder.delete


Class Folder
    Private strFolderPath    
	Private strFolderName
	Private rsFiles
	Private rsSubFolders
	Private bolConnected
	Private Sub Class_Initialize
		strFolderPath = ""
		strFolderName = ""
	End Sub
    Public Property Let path(inputname)
        strFolderPath = inputname
		strFolderName = inputname
		if right(strFolderName,1) = "\" then strFolderName = left(strFolderName, len(strFolderName)-1)
		strFolderName = right(strFolderName, len(strFoldername)-instrrev(strFolderName, "\"))
    End Property
    Public Property Get name()
		name = strFolderName
    End Property
	Public Property Get path()
		path = strFolderPath
	End Property
    Public Property Get drive()
        If Me.exists Then
            Dim objFS, objFolder
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFolder = objFS.GetFolder(strFolderPath)
            drive = objFolder.drive
			set objFolder = Nothing
			set objFS = Nothing
        End If
    End Property
    Public Property Get shortName()
        If Me.exists Then
            Dim objFS, objFolder
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFolder = objFS.GetFolder(strFolderPath)
            shortName = objFolder.shortName
			set objFolder = Nothing
			set objFS = Nothing			
        Else
            shortName = False
        End If
    End Property
    Public Property Get shortPath()
        If Me.exists Then
            Dim objFS, objFolder
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFolder = objFS.GetFolder(strFolderPath)
            shortPath = objFolder.shortPath
			set objFolder = Nothing
			set objFS = Nothing			
        Else
            shortPath = False
        End If
    End Property
    Public Property Get exists()
		dim objFS
		Set objFS = CreateObject("Scripting.FileSystemObject")
		If objFS.FolderExists(strFolderPath) Then	
			exists = True
		Else
			exists = False
		End If
		set objFS = Nothing
    End Property
    Public Property Get dateCreated()
        If Me.exists Then
            Dim objFS, objFolder
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFolder = objFS.GetFolder(strFolderPath)
            dateCreated = objFolder.dateCreated
			set objFolder = Nothing
			set objFS = Nothing			
        Else
            dateCreated = False
        End If
    End Property
    Public Property Get dateModified()
        If Me.exists Then
            Dim objFS, objFolder
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFolder = objFS.GetFolder(strFolderPath)
            dateModified = objFolder.DateLastModified
			set objFolder = Nothing
			set objFS = Nothing			
        Else
            dateModified = False
        End If
    End Property
    Public Property Get dateAccessed()
        If Me.exists Then
            Dim objFS, objFolder
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFolder = objFS.GetFolder(strFolderPath)
            dateAccessed = objFolder.DateLastAccessed
			set objFolder = Nothing
			set objFS = Nothing
        Else
            dateAccessed = False
        End If
    End Property
    Public Property Get readOnly()
        readOnly = False
        If Me.exists Then
            Dim attributes
            attributes = Me.attributes
            If GetBit(attributes, 1) Then readOnly = True
        End If
    End Property
    Public Property Let readOnly(value)
        If Me.exists Then
            Dim objFS, objFolder, attributes
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFolder = objFS.GetFolder(strFolderPath)
            If value = True Then
                    objFolder.attributes = SetBit(Me.attributes, 1, 1)
            Else
                    objFolder.attributes = SetBit(Me.attributes, 1, 0)
            End If
			set objFolder = Nothing
			set objFS = Nothing
        End If
    End Property
    Public Property Get hidden()
        hidden = False
        If Me.exists Then
            Dim attributes
            attributes = Me.attributes
            If GetBit(attributes, 2) Then hidden = True
        End If
    End Property
    Public Property Let hidden(value)
        If Me.exists Then
            Dim objFS, objFolder, attributes
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFolder = objFS.GetFolder(strFolderPath)
            If value = True Then
                objFolder.attributes = SetBit(Me.attributes, 2, 1)
            Else
                objFolder.attributes = SetBit(Me.attributes, 2, 0)
            End If
			set objFolder = Nothing
			set objFS = Nothing
        End If
    End Property
    Public Property Get system()
        system = False
        If Me.exists Then
            Dim attributes
            attributes = Me.attributes
            If GetBit(attributes, 3) Then system = True
        End If
    End Property
    Public Property Get archive()
        archive = False
        If Me.exists Then
            Dim attributes
            attributes = Me.attributes
            If GetBit(attributes, 6) Then archive = True
        End If
    End Property
    Public Property Let archive(value)
        If Me.exists Then
            Dim objFS, objFolder, attributes
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFolder = objFS.GetFolder(strFolderPath)
            If value = True Then
                    objFolder.attributes = SetBit(Me.attributes, 6, 1)
            Else
                    objFolder.attributes = SetBit(Me.attributes, 6, 0)
            End If
			set objFolder = Nothing
			set objFS = Nothing			
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
            Dim objFS, objFolder
            Set objFS = CreateObject("Scripting.FileSystemObject")
            Set objFolder = objFS.GetFolder(strFolderPath)
            attributes = objFolder.attributes
			set objFolder = Nothing
			set objFS = Nothing			
        Else
            attributes = -1
        End If
    End Property
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
    Public Function rename(ByVal strNewName)
		if strFolderPath <> "" Then
			rename = False
			If Me.exists Then
				Dim objFS, objFolder
				Set objFS = CreateObject("Scripting.FileSystemObject")
				Set objFolder = objFS.GetFolder(strFolderPath)
				On Error Resume Next
				objFolder.name = strNewName
				strFolderPath = objFolder.path
				If Err = 0 Then rename = True
				On Error GoTo 0
				set objFolder = Nothing
				set objFS = Nothing
			End If
		Else
			msgbox "You must set the path property before renaming the folder."
		End If			
    End Function
    Public Function move(ByVal strTargetFolderPath) 
		if strFolderPath <> "" Then
			If Right(strTargetFolderPath, 1) <> "\" Then strTargetFolderPath = strTargetFolderPath & "\"
			move = False
			If Me.exists Then
				Dim objFS, objFolder
				Set objFS = CreateObject("Scripting.FileSystemObject")
				On Error Resume Next
				'move the folder
				objFS.MoveFolder strFolderPath, strTargetFolderPath
				
				'update the folder object path to reference the new moved location
				strFolderPath = strTargetFolderPath & strFolderName
				
				If Err.Number = 0 Then move = True
				On Error GoTo 0
				set objFolder = Nothing
				set objFS = Nothing
			End If
		Else
			msgbox "You must set the path property before moving the folder."
		End If	
    End Function
    Public Function delete()
		if strFolderPath <> "" Then
			Dim objFS
			Set objFS = CreateObject("Scripting.FileSystemObject")
			On Error Resume Next
			If Me.exists = True Then objFS.Deletefolder (strFolderPath)
			If Me.exists = False Then
				delete = True
			Else
				delete = False
			End If
			On Error GoTo 0
			set objFolder = Nothing
			set objFS = Nothing
		Else
			msgbox "You must set the path property before deleting the folder."
		End If
	End Function
	Public Function create()
		if strFolderPath <> "" Then
			Dim objFS
			Set objFS = CreateObject("Scripting.FileSystemObject")
			On Error Resume Next
			if me.exists = false Then
				objFS.CreateFolder strFolderPath
				if me.Exists then 
					create = True
				Else
					create = False
				End If
			Else
				create = True
			End If
			on error Goto 0
			set objFS = Nothing
		Else
			msgbox "You must set the path property before creating the folder."
		End If
	End Function
	Public Function open()
		if strFolderPath <> "" Then
			if me.exists Then
				dim objFS, objFolder, objFiles, objFile, objSubFolders, objSubFolder
				set objFS = CreateObject("Scripting.FileSystemObject")


				'setup a database object for the files in the folder
				set rsFiles = Nothing
				set rsFiles = CreateObject("ADOR.Recordset")
				rsFiles.Fields.Append "FileName",200,255
				rsFiles.Open
				
				'get the files in the folder into the recordset
				Set objFolder = objFS.GetFolder(strFolderPath)
				set objFiles = objFolder.Files
				for each objFile in objFiles
					rsFiles.AddNew
					rsFiles("FileName").Value = objFile.Name
					rsFiles.Update
				next
								
				'setup a database object for the folders in the folder
				set rsSubFolders = Nothing
				set rsSubFolders = CreateObject("ADOR.Recordset")
				rsSubFolders.Fields.Append "SubFolderName",200,255
				rsSubFolders.Open
				
				'get the folders in the folder into the recordset
				Set objFolder = objFS.GetFolder(strFolderPath)
				set objSubFolders = objFolder.SubFolders
				for each objSubFolder in objSubFolders
					rsSubFolders.AddNew
					rsSubFolders("SubFolderName").Value = objSubFolder.Name
					rsSubFolders.Update
				next			
				
				'move to the start of each recordset
				if not rsFiles.BOF and not rsFiles.EOF then rsFiles.MoveFirst
				if not rsSubFolders.BOF and not rsSubFolders.EOF then rsSubFolders.MoveFirst
				
				'Set the connected Property
				bolConnected = True
			End If
		Else
			msgbox "You must set the path property before opening the folder."
		End If
	End Function
	Public Property Get getFileName()
        if bolConnected Then
			if not rsFiles.BOF and not rsFiles.EOF then
				getFileName = rsFiles("FileName").value
			Else
				getFileName = ""
			end if
		Else
			msgbox "You must first connect to a folder with the Connect verb."
		End If
    End Property
	Public Function nextFile()
		if bolConnected Then
			if not rsFiles.BOF and not rsFiles.EOF then
				rsFiles.MoveNext
			end if
		Else
			msgbox "You must first connect to a folder with the Connect verb."
		End If
    End Function
	Public Property Get getSubFolderName()
        if bolConnected Then
			if not rsSubFolders.BOF and not rsSubFolders.EOF then
				getSubFolderName = rsSubFolders("SubFolderName").value
			Else
				getSubFolderName = ""
			end if
		Else
			msgbox "You must first connect to a folder with the Connect verb."
		End If
    End Property
	Public Function nextSubFolder()
		if bolConnected Then
			if not rsSubFolders.BOF and not rsSubFolders.EOF then
				rsSubFolders.MoveNext
			end if
		Else
			msgbox "You must first connect to a folder with the Connect verb."
		End If
    End Function		
End Class
