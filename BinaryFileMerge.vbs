Option Explicit

' https://www.w3schools.com/asp/ado_ref_stream.asp
' The ADO Stream Object is used to read, write, and manage a stream of binary data or text.
'
' Syntax
' objectname.property
' objectname.method
'
' | Property      | Description                                                                           | 
' | ------------- | ------------------------------------------------------------------------------------- | 
' | CharSet       | Sets or returns a value that specifies into which character                           | 
' |               | set the contents are to be translated.                                                | 
' |               | This property is only used with text Stream objects (type is adTypeText)              | 
' | EOS           | Returns whether the current position is at the end of the stream or not               | 
' | LineSeparator | Sets or returns the line separator character used in a text Stream object             | 
' | Mode          | Sets or returns the available permissions for modifying data                          | 
' | Position      | Sets or returns the current position (in bytes) from the beginning of a Stream object | 
' | Size          | Returns the size of an open Stream object                                             | 
' | State         | Returns a value describing if the Stream object is open or closed                     | 
' | Type          | Sets or returns the type of data in a Stream object                                   | 
'
' | Constant      | Value | Description        | 
' | ------------- | ----- | ------------------ | 
' | adTypeBinary  | 1     | Binary data        | 
' | adTypeText    | 2     | Default. Text data | 
'
' | Method       | Description                                                                                     | 
' | ------------ | ----------------------------------------------------------------------------------------------- | 
' | Cancel       | Cancels an execution of an Open call on a Stream object                                         | 
' | Close        | Closes a Stream object                                                                          | 
' | CopyTo       | Copies a specified number of characters/bytes from one Stream object into another Stream object | 
' | Flush        | Sends the contents of the Stream buffer to the associated underlying object                     | 
' | LoadFromFile | Loads the contents of a file into a Stream object                                               | 
' | Open         | Opens a Stream object                                                                           | 
' | Read         | Reads the entire stream or a specified number of bytes from a binary Stream object              | 
' | ReadText     | Reads the entire stream, a line, or a specified number of characters from a text Stream object  | 
' | SaveToFile   | Saves the binary contents of a Stream object to a file                                          | 
' | SetEOS       | Sets the current position to be the end of the stream (EOS)                                     | 
' | SkipLine     | Skips a line when reading a text Stream                                                         | 
' | Write        | Writes binary data to a binary Stream object                                                    | 
' | WriteText    | Writes character data to a text Stream object                                                   | 


' Purpose: Merge files in binary mode
Sub MergeFiles(ByVal strDirIn, ByVal strDirOut, ByVal strFName)
    const adTypeBinary = 1
    ' Input stream
	Dim objStreamIn: Set objStreamIn = WScript.CreateObject("ADODB.Stream")
	objStreamIn.Open
	objStreamIn.type = adTypeBinary
    ' output stream
	Dim objStreamOut: Set objStreamOut = WScript.CreateObject("ADODB.Stream")
	objStreamOut.Open
	objStreamOut.type = adTypeBinary
    ' Loop
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objDirIn: Set objDirIn = objFSO.GetFolder(strDirIn)
	Dim objFile: For Each objFile In objDirIn.Files
        objStreamIn.LoadFromFile(objFile.Path)
        objStreamOut.Write = objStreamIn.Read()
	Next
    ' Save to File
    objStreamOut.SaveToFile(strDirOut & "\" & strFName)
    ' Garbage collection
    Set objStreamIn = Nothing
    Set objStreamOut = Nothing
    Set objDirIn = Nothing
    Set objFile = Nothing
    Set objFSO = Nothing
End Sub

' cscript BinaryFileMerge.vbs <DirIn> <DirOut> <FName>
Sub Test()
    Dim strDirIn: strDirIn = WScript.Arguments(0)
    Dim strDirOut: strDirOut = WScript.Arguments(1)
    Dim strFName: strFName = WScript.Arguments(2)
    MergeFiles strDirIn, strDirOut, strFName
End Sub
Test
