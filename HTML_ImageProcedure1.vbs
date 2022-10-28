Option Explicit

Function blnIsExtensionMatched(ByVal strFp, ByRef varExtensions)
    Dim strExt, intIndex
    strExt = LCase(CreateObject("Scripting.FileSystemObject").GetExtensionName(strFp))
    For intIndex = LBound(varExtensions) To UBound(varExtensions)
        If strExt = varExtensions(intIndex) Then
            blnIsExtensionMatched = True
            Exit Function
        End If
    Next
    blnIsExtensionMatched = False
End Function

Sub WriteTextASCII(ByVal strFp, ByVal strTxt)
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objTS: Set objTS = objFSO.OpenTextFile(strFp, 2, True)	' 2: ForWriting
	Call objTS.Write(strTxt)
	objTS.Close
	Set objFSO = Nothing
End Sub

Sub Main()
    Dim strDirImg: strDirImg = WScript.Arguments(0)
    Dim strDirOut: strDirOut = WScript.Arguments(1)
    Dim varImageExtensionsLowCase: varImageExtensionsLowCase = Array("png", "jpg", "gif")
    Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
    ' Include
    Dim strDirInclude: strDirInclude = objFSO.BuildPath(strDirOut, "include")
    If Not objFSO.FolderExists(strDirInclude) Then objFSO.CreateFolder(strDirInclude)
    ' CSS
    Dim strFpCss: strFpCss = objFSO.BuildPath(strDirInclude, "styles.css")
    Dim strTxtCss: strTxtCss = Join(Array( _
    "body {background-color: black; color: white;}", _
    "table {width: 70%;}", _
    ".lang1 {color: yellow;}", _
    ".lang2 {color: blue;}"), vbCrlf)
    ' HTML
    Dim strFpHtml: strFpHtml = objFSO.BuildPath(strDirOut, "index.html")
    Dim strTxtHtml: strTxtHtml = Join(Array( _
    "<!DOCTYPE html>", _
    "<html>", _
    "<head>", _
    "<title></title>", _
    "<meta charset=""UTF-8"">", _
    "<link rel=""stylesheet"" href=""./include/style.css"">", _
    "</head>", _
    "<body>", _
    "<h1></h1>"), vbCrlf) & vbCrlf & vbCrlf
    Dim objFile, strFpImgOut, strFpImgOutRelative
    For Each objFile In objFSO.GetFolder(strDirImg).Files
        If blnIsExtensionMatched(objFile.Path, varImageExtensionsLowCase) Then
            strFpImgOut = objFSO.BuildPath(strDirInclude, objFile.Name)
            objFSO.CopyFile objFile.Path, strFpImgOut, True
            strFpImgOutRelative = Replace(strFpImgOut, strDirOut & "\", "")
            strFpImgOutRelative = Replace(strFpImgOutRelative, "\", "/")
            strTxtHtml = strTxtHtml & Join(Array( _
            "<h2>" & objFSO.GetBaseName(objFile.Name) & "</h2>", _
            "<img src=""" & strFpImgOutRelative & """ alt=""" & objFile.Name & """>", _
            "<table><tr>", _
            "<td class=""lang1"">", _
            "<td class=""lang2"">", _
            "</tr></table>"), vbCrlf) & vbCrlf & vbCrlf
        End If
    Next
    strTxtHtml = strTxtHtml & Join(Array( _
    "</body>", _
    "</html>"), vbCrlf) & vbCrlf
    ' Write files
    WriteTextASCII strFpCss, strTxtCss
    WriteTextASCII strFpHtml, strTxtHtml
    ' Clean
    Set objFile = Nothing
    Set objFSO = Nothing
End Sub
Main
