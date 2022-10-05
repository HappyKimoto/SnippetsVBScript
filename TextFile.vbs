Option Explicit

'
' TextFile.vbs
' Read and/or write text files.
'

' Scripting.FileSystemObject
' .OpenTextFile (filename, [ iomode, [ create, [ format ]]])
'
' Part      Description
' --------  ----------------------------------------------------------------------------
' object    Required. Always the name of a FileSystemObject.
' --------  ----------------------------------------------------------------------------
' filename  Required. String expression that identifies the file to open.
' --------  ----------------------------------------------------------------------------
' iomode    Optional.
'           Indicates input/output mode. 
'           Can be one of three constants: ForReading, ForWriting, or ForAppending.
' --------  ----------------------------------------------------------------------------
' create    Optional.
'           Boolean value that indicates whether a new file can be created
'            if the specified filename doesn't exist.
'           The value is True if a new file is created; False if it isn't created.
'           The default is False.
' --------  ----------------------------------------------------------------------------
' format    Optional.
'           One of three Tristate values used to indicate the format of the opened file.
'           If omitted, the file is opened as ASCII.
'
'  Constant      Value  Description                                    
'  ------------  -----  ---------------------------------------------  
'  ForReading    1      Open a file for reading only.                 
'  ForWriting    2      Open a file for writing only.                 
'  ForAppending  8      Open a file and write to the end of the file. 
'
'  Constant            Value  Description                                  
'  ------------------  -----  -------------------------------------------  
'  TristateUseDefault  -2     Opens the file by using the system default.  
'  TristateTrue        -1     Opens the file as Unicode.                   
'  TristateFalse       0      Opens the file as ASCII.                     

' Purpose: Read text contents
' Inputs: strFp - absolute file path to the input file
' Returns: string of the file text in ASCII encoding
Function strReadTextASCII(ByVal strFp)
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objTS: Set objTS = objFSO.OpenTextFile(strFp, 1)	'1: ForReading
	strReadTextASCII = objTS.ReadAll()
	objTS.Close
	Set objFSO = Nothing
End Function

' Purpose: Write text contents to the output file
' Inputs: strFp - absolute file path to the output file
'         strTxt - Text contents
' Assumptions: strTxt is encoded in ASCII
' Effects: strTxt will be written to strFp
Sub WriteTextASCII(ByVal strFp, ByVal strTxt)
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objTS: Set objTS = objFSO.OpenTextFile(strFp, 2, True)	' 2: ForWriting
	Call objTS.Write(strTxt)
	objTS.Close
	Set objFSO = Nothing
	WScript.Echo "WriteTextASCII Err.Number=" & Err.Number
End Sub

' Purpose: Append text contents to the output file
' Inputs: strFp - absolute file path to the output file
'         strTxt - Text contents
' Assumptions: strTxt is encoded in ASCII
' Effects: strTxt will be appended to strFp
Sub AppendTextASCII(ByVal strFp, ByVal strTxt)
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objTS: Set objTS = objFSO.OpenTextFile(strFp, 8, True)	'8: ForAppending
	Call objTS.Write(strTxt)
	objTS.Close
	Set objFSO = Nothing
End Sub

' https://www.w3schools.com/asp/ado_ref_stream.asp
' The ADO Stream Object is used to read, write, and manage a stream of binary data or text.
'
'   Property        Description                                                                
'   -------------   -------------------------------------------------------------------------  
'   CharSet         Sets or returns a value that specifies into                                
'                   which character set the contents are to be translated.                     
'                   This property is only used with text Stream objects (type is adTypeText)   
'   EOS             Returns whether the current position is at the end of the stream or not    
'   LineSeparator   Sets or returns the line separator character used in a text Stream object  
'   Mode            Sets or returns the available permissions for modifying data               
'   Position        Sets or returns the current position (in bytes)                            
'                   from the beginning of a Stream object                                      
'   Size            Returns the size of an open Stream object                                  
'   State           Returns a value describing if the Stream object is open or closed          
'   Type            Sets or returns the type of data in a Stream object                        
'
'  Method         Description                                                                                      
'  ------------   -----------------------------------------------------------------------------------------------  
'  Cancel         Cancels an execution of an Open call on a Stream object                                          
'  Close          Closes a Stream object                                                                           
'  CopyTo         Copies a specified number of characters/bytes from one Stream object into another Stream object  
'  Flush          Sends the contents of the Stream buffer to the associated underlying object                      
'  LoadFromFile   Loads the contents of a file into a Stream object                                                
'  Open           Opens a Stream object                                                                            
'  Read           Reads the entire stream or a specified number of bytes from a binary Stream object               
'  ReadText       Reads the entire stream, a line, or a specified number of characters from a text Stream object   
'  SaveToFile     Saves the binary contents of a Stream object to a file                                           
'  SetEOS         Sets the current position to be the end of the stream (EOS)                                      
'  SkipLine       Skips a line when reading a text Stream                                                          
'  Write          Writes binary data to a binary Stream object                                                     
'  WriteText      Writes character data to a text Stream object                                                    

' Purpose: Read text contents
' Inputs: strFp - absolute file path to the input file
' Returns: string of the file text in Shift-JIS encoding
Function strReadTextShiftJIS(ByVal strFp)
	Dim objStream: Set objStream = CreateObject("ADODB.Stream")
	With objStream
		.CharSet = "shift-jis"
		.Open
		.LoadFromFile(strFp)
		strReadTextShiftJIS = .ReadText()
	End With
	Set objStream = Nothing
End Function

' Purpose: Read text contents
' Inputs: strFp - absolute file path to the input file
'         strCharSet - let the user define the encoding
' Returns: string of the file text read in the specified encoding
Function strReadText(ByVal strFp, ByVal strCharSet)
	Dim objStream: Set objStream = CreateObject("ADODB.Stream")
	With objStream
		.CharSet = strCharSet
		.Open
		.LoadFromFile(strFp)
		strReadText = .ReadText()
	End With
	Set objStream = Nothing
End Function

' Purpose: Write text contents to the output file
' Inputs: strFp - absolute file path to the output file
'         strTxt - Text contents
' Assumptions: strTxt is encoded in ShiftJIS
' Effects: strTxt will be written to strFp
Sub WriteTextShiftJIS(ByVal strFp, ByVal strTxt)
	Dim objStream: Set objStream = CreateObject("ADODB.Stream")
	With objStream
		.CharSet = "shift-jis"
		.Open
		.WriteText strTxt
		.SaveToFile strFp
	End With
	Set objStream = Nothing
End Sub

' Purpose: Write text contents to the output file
' Inputs: strFp - absolute file path to the output file
'         strTxt - Text contents
'         strCharSet - Encoding to write with
' Assumptions: strTxt should agree with strCharSet
' Effects: strTxt will be written to strFp
Sub WriteText(ByVal strFp, ByVal strTxt, ByVal strCharSet)
	Dim objStream: Set objStream = CreateObject("ADODB.Stream")
	With objStream
		.CharSet = strCharSet
		.Open
		.WriteText strTxt
		.SaveToFile strFp
	End With
	Set objStream = Nothing
End Sub

Sub TestRead()
	'Arguments: (0): File Path, (1): Encoding
	WScript.Echo strReadText(WScript.Arguments(0), WScript.Arguments(1))
End Sub
TestRead

