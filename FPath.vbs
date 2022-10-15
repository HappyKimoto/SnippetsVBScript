Option Explicit

' FPaths.vbs
' - Create an array of absolute file paths.

' Reference
' https://www.tutorialspoint.com/vbscript/vbscript_fso_objects.htm
'
' -----------------------------------------------------------------
' File
'   File is an Object, which contains both properties and methods
'   that allow the developers to create, delete or move a file.
'
' - Methods
'   - Copy
'   - Delete
'   - Move
'   - openasTextStream
'
' - Properties
'   - Attributes
'   - DateCreated
'   - DateLastAccessed
'   - DateLastModified
'   - Drive
'   - Name
'   - ParentFolder
'   - Path
'   - ShortName
'   - ShortPath
'   - Size
'   - Type
' -----------------------------------------------------------------
' Files
'   Files is a collection,
'   which provides a list of all files contained within a folder.
'
' - Properties
'   - Count
'   - Item
' -----------------------------------------------------------------
' Folder
'   Folder is an Object, 
'   which contains both properties and methods that allow the developers to 
'   create, delete or move a folder.
'
' - Methods
'   - Copy
'   - Delete
'   - Move
'   - CreateTextFile
'
' - Properties
'   - Attributes
'   - DateCreated
'   - DateLastAccessed
'   - DateLastModified
'   - Drive
'   - Files
'   - IsRootFolder
'   - Name
'   - ParentFolder
'   - Path
'   - ShortName
'   - ShortPath
'   - Size
'   - SubFolders
'   - Type
' -----------------------------------------------------------------
' Folders
'   Folders is an collection of all Folder Objects within a Folder object.
'
' - Methods
'   - Add
'   - Properties
'   - Count
'   - Item

' Purpose: Create an array of File object properties
' Inputs: objFile as File Object
' Return: Variant() of file object properties
Function varFilePropArray(ByRef objFile)
	varFilePropArray = Array( _
	objFile.Path, _
	objFile.Name, _
	objFile.DateLastModified, _
	objFile.Size)
End Function

' | Scope           | Prefix | Example            | 
' | --------------- | ------ | ------------------ | 
' | Procedure-level | None   | dblVelocity        | 
' | Script-level    | s      | sblnCalcInProgress | 
Const sconFilePropertyArrayPath = 0
Const sconFilePropertyArrayName = 1
Const sconFilePropertyArrayDateLastModified = 2
Const sconFilePropertyArraySize = 3

' Purpose: Filter out the file attribute array by file path using RegExp
' Input: ByRef varFileAttr as file attribute array
' 		 ByVal strPattern as RegExp pattern
' Return: newly created trimmed array
Function varFileAttrArrayFilteredByPath(ByRef varFileAttr, ByVal strPattern)
	Dim objRgx: Set objRgx = New RegExp
	With objRgx
		.Pattern = strPattern
		.Global = False	' should not be global for file path pattern check
		.IgnoreCase = False	' windows file path system is case insensitive
	End With
	Dim varReturn: ReDim varReturn(-1)
	Dim lngIdxOrg, lngIdxNew
	lngIdxNew = -1
	For lngIdxOrg = LBound(varFileAttr) To UBound(varFileAttr)
		If objRgx.Test(varFileAttr(lngIdxOrg)(sconFilePropertyArrayPath)) Then
			lngIdxNew = lngIdxNew + 1
			ReDim Preserve varReturn(lngIdxNew)
			varReturn(lngIdxNew) = varFileAttr(lngIdxOrg)
		End if
	Next
	varFileAttrArrayFilteredByPath = varReturn
End Function

' Purpose: Sort file array according to column index with bubble sort algorithm
' Inputs: varFileAttr - Array of File Attributes; each element is varFilePropArray
'         intCol - the column index of varFilePropArray
'                  If varFilePropArray is defined as below and want to sort by file timestamp,
'				   then pass 2 as intCol
'                   -----------------------------------
'					varFilePropArray = Array(
'					objFile.Path, _					(0) Absolute file path
'					objFile.Name, _					(1) File name
'					objFile.DateLastModified, _		(2) File timestamp
'					objFile.Size)					(3) File size
'                   -----------------------------------
'                   It is recommended to use one of the script scope constant
'                   -----------------------------------
'                   - Const sconFilePropertyArrayPath = 0
'                   - Const sconFilePropertyArrayName = 1
'                   - Const sconFilePropertyArrayDateLastModified = 2
'                   - Const sconFilePropertyArraySize = 3
'                   -----------------------------------
' Assumptions: varFileAttr has been already populated
'			   by either MapFilesTopOnly or MapFilesRecursively
Sub SortFileAttrArray(ByRef varFileAttr, ByVal intCol)
	Dim i, j, intSwapCount, varTempAttr
	For i = LBound(varFileAttr) + 1 To UBound(varFileAttr)
		intSwapCount = 0
		For j = LBound(varFileAttr) + 1 To UBound(varFileAttr)
			If varFileAttr(j-1)(intCol) > varFileAttr(j)(intCol) Then
				varTempAttr = varFileAttr(j-1)
				varFileAttr(j-1) = varFileAttr(j)
				varFileAttr(j) = varTempAttr
				intSwapCount = intSwapCount + 1
			End If
		Next
		If intSwapCount = 0 Then Exit For
	Next
End Sub

' Purpose: Create an array of file paths.
' Inputs: strRootDir - absolute path to the root folder
'         varFileAttr - passed as an empty variable, but will become an array of file paths.
' Effect: Map only the root folder.
' Assumption: The caller will use varFileAttr.
Sub MapFilesTopOnly(ByVal strRootDir, ByRef varFileAttr)
	ReDim varFileAttr(-1)
	Dim lngUB
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objRootDir: Set objRootDir = objFSO.GetFolder(strRootDir)
	Dim objFile
	For Each objFile In objRootDir.Files
		lngUB = UBound(varFileAttr) + 1
		ReDim Preserve varFileAttr(lngUB)
		varFileAttr(lngUB) = varFilePropArray(objFile)
	Next
	Set objFile = Nothing
	Set objRootDir = Nothing
	Set objFSO = Nothing
End Sub

' Purpose: To be called from MapFilesRecursively
' Inputs: objDir as Folder object where files exist
'         varFileAttr as array
' Effects: varFileAttr will keep accumulating through recursive calls.
Sub MapFilesRecursivelySub(ByRef objDir, ByRef varFileAttr)
	Dim lngUB
	' Get file attributes.
	Dim objFile: For Each objFile In objDir.Files
		lngUB = UBound(varFileAttr) + 1
		ReDim Preserve varFileAttr(lngUB)
		varFileAttr(lngUB) =  varFilePropArray(objFile)
	Next
	' Call recursively on subfolders.
	Dim objSubDir: For Each objSubDir In objDir.SubFolders
		Call MapFilesRecursivelySub(objSubDir, varFileAttr)
	Next
	Set objFile = Nothing
	Set objSubDir = Nothing
End Sub

' Purpose: Create an array of file paths.
' Inputs: strRootDir - absolute path to the root folder
'         varFileAttr - passed as an empty variable, but will become an array of file paths.
' Effect: Map recursively.
' Assumption: The caller will use varFileAttr.
Sub MapFilesRecursively(ByVal strRootDir, ByRef varFileAttr)
	ReDim varFileAttr(-1)
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objRootDir: Set objRootDir = objFSO.GetFolder(strRootDir)
	Call MapFilesRecursivelySub(objRootDir, varFileAttr)
	Set objRootDir = Nothing
	Set objFSO = Nothing
End Sub

' Purpose: Test FPath.vbs
' Assumption: Run like >cscript FPath.vbs <RootFolder> .*txt
Sub TestMapping()
	Dim strRootDir, varFileAttr, strFilterRgxPattern
	strRootDir = WScript.Arguments(0)
	strFilterRgxPattern = WScript.Arguments(1)
	' Loop through files and gather file attribute
	MapFilesRecursively strRootDir, varFileAttr
	' Sort array
	SortFileAttrArray varFileAttr, sconFilePropertyArrayDateLastModified
	' Filter by regular expression
	varFileAttr = varFileAttrArrayFilteredByPath(varFileAttr, strFilterRgxPattern)
	' Print the result
	Dim i: For i = LBound(varFileAttr) To UBound(varFileAttr)
		WScript.Echo Join(varFileAttr(i), "~")
	Next
End Sub
' TestMapping

' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object
'
'  | Method              | Description                                                                       | 
'  | ------------------- | --------------------------------------------------------------------------------- | 
'  | BuildPath           | Appends a name to an existing path.                                               | 
'  | CopyFile            | Copies one or more files from one location to another.                            | 
'  | CopyFolder          | Copies one or more folders from one location to another.                          | 
'  | CreateFolder        | Creates a new folder.                                                             | 
'  | CreateTextFile      | Creates a text file and returns a TextStream object                               | 
'  |                     | that can be used to read from, or write to the file.                              | 
'  | DeleteFile          | Deletes one or more specified files.                                              | 
'  | DeleteFolder        | Deletes one or more specified folders.                                            | 
'  | DriveExists         | Checks if a specified drive exists.                                               | 
'  | FileExists          | Checks if a specified file exists.                                                | 
'  | FolderExists        | Checks if a specified folder exists.                                              | 
'  | GetAbsolutePathName | Returns the complete path from the root of the drive for the specified path.      | 
'  | GetBaseName         | Returns the base name of a specified file or folder.                              | 
'  | GetDrive            | Returns a Drive object corresponding to the drive in a specified path.            | 
'  | GetDriveName        | Returns the drive name of a specified path.                                       | 
'  | GetExtensionName    | Returns the file extension name for the last component in a specified path.       | 
'  | GetFile             | Returns a File object for a specified path.                                       | 
'  | GetFileName         | Returns the file name or folder name for the last component in a specified path.  | 
'  | GetFolder           | Returns a Folder object for a specified path.                                     | 
'  | GetParentFolderName | Returns the name of the parent folder of the last component in a specified path.  | 
'  | GetSpecialFolder    | Returns the path to some of Windows' special folders.                             | 
'  | GetTempName         | Returns a randomly generated temporary file or folder.                            | 
'  | Move                | Moves a specified file or folder from one location to another.                    | 
'  | MoveFile            | Moves one or more files from one location to another.                             | 
'  | MoveFolder          | Moves one or more folders from one location to another.                           | 
'  | OpenAsTextStream    | Opens a specified file and returns a TextStream object                            | 
'  |                     | that can be used to read from, write to, or append to the file.                   | 
'  | OpenTextFile        | Opens a file and returns a TextStream object that can be used to access the file. | 
'  | WriteLine           | Writes a specified string and new-line character to a TextStream file.            | 
'
'  | Property | Description                                                                                   | 
'  | -------- | --------------------------------------------------------------------------------------------- | 
'  | Drives   | Returns a collection of all Drive objects on the computer.                                    | 
'  | Name     | Sets or returns the name of a specified file or folder.                                       | 
'  | Path     | Returns the path for a specified file, folder, or drive.                                      | 
'  | Size     | For files, returns the size, in bytes, of the specified file;                                 | 
'  |          | for folders, returns the size, in bytes, of all files and subfolders contained in the folder. | 
'  | Type     | Returns information about the type of a file or folder                                        | 
'  |          | (for example, for files ending in .TXT, "Text Document" is returned).                         | 

Sub TestGetNames()
	Dim strFp: strFp = WScript.Arguments(0)
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	WScript.Echo "strFp=" & strFp
	WScript.Echo "objFSO.GetExtensionName(strFp)=" & objFSO.GetExtensionName(strFp)
	WScript.Echo "objFSO.GetBaseName(strFp)=" & objFSO.GetBaseName(strFp)
	WScript.Echo "objFSO.GetDriveName(strFp)=" & objFSO.GetDriveName(strFp)
	WScript.Echo "objFSO.GetFileName(strFp)=" & objFSO.GetFileName(strFp)
	WScript.Echo "objFSO.GetParentFolderName(strFp)=" & objFSO.GetParentFolderName(strFp)
	WScript.Echo "objFSO.GetAbsolutePathName(strFp)=" & objFSO.GetAbsolutePathName(strFp)
	Set objFSO = Nothing
End Sub
TestGetNames
