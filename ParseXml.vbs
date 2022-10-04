Option Explicit

Const conXmlFp = "Sample.xml"

Sub Test1()
    ' Create DOM object
    Dim objDOM, objNode, objNodes
    Set objDOM = CreateObject("Msxml2.DOMDocument.6.0")
    WScript.Echo TypeName(objDOM)   ' DOMDocument60

    ' Load file
    objDOM.Load conXmlFp

    ' Single node
    Set objNode = objDOM.SelectSingleNode( _
    "/Root/Body/Message [@lang='English' and @index='1']")
    WScript.Echo TypeName(objNode)  ' IXMLDOMElement

    ' Node text
    WScript.Echo objNode.Text, TypeName(objNode.Text)   ' Hello World String

    ' Single node without specific attribute
    Set objNode = objDOM.SelectSingleNode( _
    "/Root/Header/Date")

    ' Node text
    WScript.Echo objNode.Text, TypeName(objNode.Text)   ' 2021/9/25 String

    ' Get attribute value
    WScript.Echo objNode.getAttribute("description")    ' Start

    ' Select Nodes
    Set objNodes = objDOM.SelectNodes("Root/Body/Message")
    WScript.echo TypeName(objNodes) ' IXMLDOMSelection

    For Each objNode In objNodes
        WScript.Echo objNode.Text, objNode.getAttribute("lang"), objNode.getAttribute("index")
    Next

    ' Test escape characters
    Set objNodes = objDOM.SelectSingleNode("Root/Escapes").childNodes
    For Each objNode In objNodes
        ' Get node name, attribute, text
        WScript.Echo Join(Array(objNode.nodeName, objNode.getAttribute("name"), objNode.Text), "//")
    Next

End Sub
Test1
