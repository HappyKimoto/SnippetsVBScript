Rem - https://blogs.iis.net/robert_mcmurray/creating-quot-pretty-quot-xml-using-xsl-and-vbscript

Sub PrettifyXml(ByVal strFpInXml)

	Dim strFpOutTxt: strFpOutTxt = strFpInXml & ".txt"

	Dim objInputFile, objOutputFile, strXML
	Dim objFSO : Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	Dim objXML : Set objXML = WScript.CreateObject("Msxml2.DOMDocument")
	Dim objXSL : Set objXSL = WScript.CreateObject("Msxml2.DOMDocument")

	' ****************************************
	' Put whitespace between tags. (Required for XSL transformation.)
	' ****************************************

	Rem - Using [-2] as system default for encoding.
	Rem  I thought this could cause some vulnerability.
	Rem  However, this whole script worked even when it is set to ANSI.
	Rem  There are some underlying technology that I do not understand.
	Rem  At least, it works on my laptop.
	Set objInputFile = objFSO.OpenTextFile(strFpInXml,1,False,-2)
	Set objOutputFile = objFSO.CreateTextFile(strFpOutTxt,True,False)
	strXML = objInputFile.ReadAll
	strXML = Replace(strXML,"><",">" & vbCrLf & "<")
	objOutputFile.Write strXML
	objInputFile.Close
	objOutputFile.Close

	' ****************************************
	' Create an XSL stylesheet for transformation.
	' ****************************************

	Dim strStylesheet : strStylesheet = _
		"<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">" & _
		"<xsl:output method=""xml"" indent=""yes""/>" & _
		"<xsl:template match=""/"">" & _
		"<xsl:copy-of select="".""/>" & _
		"</xsl:template>" & _
		"</xsl:stylesheet>"

	' ****************************************
	' Transform the XML.
	' ****************************************

	objXSL.loadXML strStylesheet
	objXML.load strFpOutTxt
	objXML.transformNode objXSL
	objXML.save strFpOutTxt

End Sub 

Class Mapper
	Private m_FSO
	Private m_Path
	
	Public Sub ProcessFile(ByVal strFp)
		If UCase(m_FSO.GetExtensionName(strFp)) = "XML" Then
			PrettifyXml(strFp)
			WScript.Echo "Prettified: " & strFp
		End If
	End Sub
	
	Public Sub CheckPath(ByVal strPath)
		If m_FSO.FileExists(strPath) Then
			Call ProcessFile(strPath)
		ElseIf m_FSO.FolderExists(strPath) Then
			Call MapFiles(strPath)
		End If
	End Sub
	
	Public Sub MapFiles(ByVal strDirRoot)
		Dim objDirRoot: Set objDirRoot = m_FSO.GetFolder(strDirRoot)
		
		Dim objDirSub
		For Each objDirSub In objDirRoot.SubFolders
			Call MapFiles(objDirSub.Path)
		Next
		
		Dim objFile
		For Each objFile In objDirRoot.Files
			Call ProcessFile(objFile.Path)
		Next
		
	End Sub

	Private Sub Class_Initialize()
		Set m_FSO = CreateObject("Scripting.FileSystemObject")
	End Sub

	Private Sub Class_Terminate()
		Set m_FSO = Nothing
	End Sub

End Class

Sub Main
	Dim strPath: strPath = WScript.Arguments(0)
	
	Dim objMapper: Set objMapper = New Mapper
	
	Call objMapper.CheckPath(strPath)
	
	Set objMapper = Nothing
End Sub

Call Main
