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
'					Assumptions: varFileAttr has been already populated
'								 by either MapFilesTopOnly or MapFilesRecursively
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
' Assumption: Run like >cscript FPath.vbs <RootFolder>
Sub Test()
	Dim strRootDir, varFileAttr
	strRootDir = WScript.Arguments(0)
	MapFilesRecursively strRootDir, varFileAttr
	
	SortFileAttrArray varFileAttr, 2
	Dim i: For i = LBound(varFileAttr) To UBound(varFileAttr)
		WScript.Echo Join(varFileAttr(i), "~")
	Next
End Sub
Test
