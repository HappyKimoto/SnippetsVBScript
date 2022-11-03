Attribute VB_Name = "mdlShapeText"
Option Explicit


'
' https://learn.microsoft.com/en-us/office/vba/api/office.msoshapetype
'
' | Name                 | Value | Description           |
' | -------------------- | ----- | --------------------- |
' | mso3DModel           |    30 | 3D model              |
' | msoAutoShape         |     1 | AutoShape             |
' | msoCallout           |     2 | Callout               |
' | msoCanvas            |    20 | Canvas                |
' | msoChart             |     3 | Chart                 |
' | msoComment           |     4 | Comment               |
' | msoContentApp        |    27 | Content Office Add-in |
' | msoDiagram           |    21 | Diagram               |
' | msoEmbeddedOLEObject |     7 | Embedded OLE object   |
' | msoFormControl       |     8 | Form control          |
' | msoFreeform          |     5 | Freeform              |
' | msoGraphic           |    28 | Graphic               |
' | msoGroup             |     6 | Group                 |
' | msoIgxGraphic        |    24 | SmartArt graphic      |
' | msoInk               |    22 | Ink                   |
' | msoInkComment        |    23 | Ink comment           |
' | msoLine              |     9 | Line                  |
' | msoLinked3DModel     |    31 | Linked 3D model       |
' | msoLinkedGraphic     |    29 | Linked graphic        |
' | msoLinkedOLEObject   |    10 | Linked OLE object     |
' | msoLinkedPicture     |    11 | Linked picture        |
' | msoMedia             |    16 | Media                 |
' | msoOLEControlObject  |    12 | OLE control object    |
' | msoPicture           |    13 | Picture               |
' | msoPlaceholder       |    14 | Placeholder           |
' | msoScriptAnchor      |    18 | Script anchor         |
' | msoShapeTypeMixed    |    -2 | Mixed shape type      |
' | msoSlicer            |    25 | Slicer                |
' | msoTable             |    19 | Table                 |
' | msoTextBox           |    17 | Text box              |
' | msoTextEffect        |    15 | Text effect           |
' | msoWebVideo          |    26 | Web video             |
'

Sub OverflowShapeText()
    
    ' Loop through shapes
    Dim thisShape As Shape
    For Each thisShape In ActiveSheet.Shapes
        
        ' Print Shape Type for debug
        Debug.Print "AutoShapeType=" & thisShape.AutoShapeType
        
        ' If the shape has text, let the text overflow in both vertially and horizontally.
        If Not thisShape.TextFrame Is Nothing Then
            If thisShape.TextFrame2.HasText Then
                Debug.Print thisShape.TextFrame.Characters.Text
                thisShape.TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
                thisShape.TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow
            End If
        End If
        
    Next
End Sub
