Attribute VB_Name = "SVGParse"
Option Explicit
Public Type pointD
    x As Double
    y As Double
    noCut As Byte
End Type

Public Type typLine
    Points() As pointD
    SpecialNumPoints As Long
    
    Fillable As Boolean ' Only works for closed paths
    
    ContainedBy As Long ' ID to containing poly
    
    xCenter As Double
    yCenter As Double
    
    Optimized As Boolean
    
    greyLevel As Byte ' 0 to GREYLEVELS level of grey, higher is lighter
    
    LayerID As String
    
    PathCode As String
    
    LevelNumber As Long 'How many levels deep is this
    
    isDel As Boolean ' Deleted on next iteration
End Type

Public containList As New Scripting.Dictionary

Public pData() As typLine
Public currentLine As Long

Public layerInfo As New Scripting.Dictionary


Public Const PI = 3.141592654
Public GLOBAL_DPI As Double

Public EXPORT_EXTENTS_X As Double, EXPORT_EXTENTS_Y As Double
Public LastExportPath As String
Public CurrentFile As String

Dim hasUnfinishedLine As Boolean


Function parseSVG(inFile As String)

    Dim SVG As New ChilkatXml
    Dim x As ChilkatXml
    Dim i As Long, j As Long
    
    Dim realW As Double
    Dim realH As Double
    Dim realDPI As Double
    
    Dim S() As String
    
    
    ReDim pData(0)
    currentLine = 0
    
    realDPI = 90
    
    SVG.LoadXmlFile inFile
    
    If SVG Is Nothing Then
        MsgBox "Could not load SVG"
        Exit Function
    End If
    
    
    'For i = 0 To SVG.childNodes.length - 1
    '    Set x = SVG.childNodes(i)
    '    If x.nodeName = "svg" Then Exit For
    'Next
    
    If SVG.Tag = "svg" Then
    
        '   width="8.5in"
        '   height="11in"
        '   viewBox="0 0 765.00001 990.00002"
    
        ' Read these numbers to determine the scale of the data inside the file.
        ' width and height are the real-world widths and heights
        ' viewbox is how we're going to scale the numbers in the file (expressed in pixels) to the native units of this program, which is inches
        
        realW = Val(SVG.GetAttrValue("width"))
        ' Read the unit
        Select Case LCase(Replace(SVG.GetAttrValue("width"), realW, ""))
            Case "in" ' no conversion needed
            Case "mm", "" ' convert from mm
                realW = realW / 25.4
            Case "cm" ' convert from cm
                realW = realW / 2.54
        
        End Select
        
        realH = Val(SVG.GetAttrValue("height"))
        ' Read the unit
        Select Case LCase(Replace(SVG.GetAttrValue("height"), realH, ""))
            Case "in" ' no conversion needed
            Case "mm", "" ' convert from mm
                realH = realH / 25.4
            Case "cm" ' convert from cm
                realH = realH / 2.54
        End Select
        
        'MsgBox "Size in inches: " & realW & ", " & realH
        
        ' The 'ViewBox' is how we scale an inch to a pixel.  The default is 90dpi but it may not be.
        
        Dim ttt As String
        'ttt = InputBox("Detected with: " & realW & " inches.  Change it?", "Width", realW)
        'If ttt <> "" Then
        '    realW = Val(ttt)
        'End If
        
        
        S = Split(SVG.GetAttrValue("viewBox"), " ")
        If UBound(S) = 3 Then
            ' Get the width in pixels
            If realW = 0 Then
                realDPI = 300
            Else
                realDPI = Val(S(2)) / realW
            End If
        End If
        
        
        If realDPI = 1 Then realDPI = 72
        
        'ttt = InputBox("Detected DPI: " & realDPI & ".  Change it?", "DPI")
        'If ttt <> "" Then
        '    realDPI = Val(ttt)
        'End If
        
        
        GLOBAL_DPI = realDPI
        
               
        parseSVGKids SVG
    End If
    
    ' Scale by the DPI
    For i = 1 To UBound(pData)
        With pData(i)
            For j = 1 To UBound(.Points)
                With .Points(j)
                    .x = .x / realDPI
                    .y = .y / realDPI
                End With
            Next
        End With
    Next
    
' Fix the extents
    Dim minX As Double
    Dim minY As Double
    
    minX = 1000000
    minY = 1000000
    
    ' Calculate the extents
    For i = 1 To UBound(pData)
        With pData(i)
            For j = 1 To UBound(.Points)
                With .Points(j)
                    minX = Min(minX, .x)
                    minY = Min(minY, .y)
                End With
            Next
        End With
    Next
    
 
    ' Now fix the points by removing space at the left and top
    
    For i = 1 To UBound(pData)
        With pData(i)
            For j = 1 To UBound(.Points)
                With .Points(j)
                    .x = .x - minX
                    .y = .y - minY
                End With
            Next
        End With
    Next


End Function


Function parseSVGKids(inEle As ChilkatXml, Optional currentLayer As String)

    ' Loop through my kids and figure out what to do!
    Dim i As Long
    Dim x As ChilkatXml
    Dim beforeLine As Long
    Dim j As Long
    
    Dim cX As Double
    Dim cY As Double
    Dim cW As Double
    Dim cH As Double
    
    Dim beforeGroup As Long
    Dim layerName As String
    
    If currentLayer = "" Then currentLayer = "BLANK"
    
    
    Debug.Print "PARSING A KIDS:", currentLayer
    
    
    Set x = inEle.FirstChild
    Do Until x Is Nothing
        
        'MsgBox X.nodeName
        
        Select Case LCase(x.Tag)
            Case "g" ' g is GROUP
                beforeGroup = currentLine
                
                ' Is this group a layer?
                layerName = getAttr(x, "inkscape:label", "")
                If layerName = "" Then
                    If InStr(1, getAttr(x, "id", ""), "layer", vbTextCompare) > 0 Then
                        layerName = getAttr(x, "id", "")
                    End If
                End If
                
                If layerName = "" Then layerName = currentLayer
                
                'If layerName = "" Then layerName = getAttr(x, "id", "")
                
                parseSVGKids x, layerName
                    
                If getAttr(x, "transform", "") <> "" Then
                    ' Transform these lines
                    For j = beforeGroup + 1 To currentLine
                        transformLine j, getAttr(x, "transform", "")
                    Next
                End If
            
            Case "switch" ' stupid crap
                parseSVGKids x
             
            ' SHAPES
            Case "rect", "path", "line", "polyline", "circle", "polygon", "ellipse"
                beforeLine = currentLine
                
                Select Case LCase(x.Tag)
                    Case "rect" ' RECTANGLE
                        
                        newLine currentLayer
                        cX = Val(getAttr(x, "x", ""))
                        cY = Val(getAttr(x, "y", ""))
                        cW = Val(getAttr(x, "width", ""))
                        cH = Val(getAttr(x, "height", ""))
                        addPoint cX, cY
                        addPoint cX + cW, cY
                        addPoint cX + cW, cY + cH
                        addPoint cX, cY + cH
                        addPoint cX, cY
                        finishLine
                        
                        pData(currentLine).Fillable = True
                    
                    Case "path"
                        
                        ' Parse the path.
                        Dim thePath As String
                        thePath = getAttr(x, "d", "")
                        If x.GetAttrValue("fill") <> "" And x.GetAttrValue("fill") <> "none" Then  ' For some reason Illustrator doesn't close paths that are filled
                            If Len(thePath) > 0 Then
                                If LCase(Right(thePath, 1)) = "z" Then
                                    ' ALready closed
                                Else
                                    thePath = thePath & "z"
                                End If
                            End If
                        End If
                        
                        parsePath thePath, currentLayer
                        
                        
                        
                        
                    Case "line"
                        ' Add this line
                        newLine currentLayer
                        addPoint Val(getAttr(x, "x1", "")), Val(getAttr(x, "y1", ""))
                        addPoint Val(getAttr(x, "x2", "")), Val(getAttr(x, "y2", ""))
                        finishLine
                        
                    Case "polyline"
                        newLine currentLayer
                        parsePolyLine getAttr(x, "points", "")
                        finishLine
                        
                    Case "polygon"
                        newLine currentLayer
                        parsePolyLine getAttr(x, "points", "")
                        finishLine
                        
                        pData(currentLine).Fillable = True
                        
                        
                    Case "circle"
                        ' Draw a circle.
                        newLine currentLayer
                        parseCircle Val(getAttr(x, "cx", "")), Val(getAttr(x, "cy", "")), Val(getAttr(x, "r", ""))
                    
                    Case "ellipse" ' Draw an ellipse
                        newLine currentLayer
                        '   cx="245.46707"
                        '   cy = "469.48389"
                        '   rx = "13.131983"
                        '   ry="14.142136" />
                        
                        parseEllipse Val(getAttr(x, "cx", "")), Val(getAttr(x, "cy", "")), Val(getAttr(x, "rx", "")), Val(getAttr(x, "ry", ""))
                End Select
                
                ' Shape transformations
                If getAttr(x, "transform", "") <> "" Then
                    ' Transform these lines
                    For j = beforeLine + 1 To currentLine
                        transformLine j, getAttr(x, "transform", "")
                    Next
                End If
        End Select
        Set x = x.NextSibling
    Loop
    
    

End Function

Function parseCircle(cX As Double, cY As Double, Radi As Double)

    Dim A As Double
    Dim x As Double, y As Double
    Dim rr As Long
    
    rr = 2
    If Radi > 100 Then rr = 1
    
    
    For A = 0 To 360 Step rr
        
        x = Cos(A * (PI / 180)) * Radi + cX
        y = Sin(A * (PI / 180)) * Radi + cY
        
        addPoint x, y
        
        
    Next
    
    pData(currentLine).Fillable = True

End Function


Function parseEllipse(cX As Double, cY As Double, RadiX As Double, RadiY As Double)

    Dim A As Double
    Dim x As Double, y As Double
    Dim rr As Long
    
    rr = 2
    If RadiX > 100 Or RadiY > 100 Then rr = 1
    
    
    For A = 0 To 360 Step rr
        
        x = Cos(A * (PI / 180)) * RadiX + cX
        y = Sin(A * (PI / 180)) * RadiY + cY
        
        addPoint x, y
        
    Next
    
    pData(currentLine).Fillable = True

End Function

Function transformLine(lineID As Long, transformText As String)

    ' Parse the transform text
    Dim e As Long, f As Long
    Dim j As Long
    
    Dim func As String
    Dim params As String
    Dim pSplit
    Dim Ang As Double
    
    
    With pData(lineID)
        
        e = InStr(1, transformText, "(")
        If e > 0 Then
            func = left(transformText, e - 1)
            f = InStr(e + 1, transformText, ")")
            If f > 0 Then params = Mid(transformText, e + 1, f - e - 1)
            
            Select Case LCase(func)
                Case "translate"
                    ' Just move everything
                    pSplit = Split(params, ",")
                    
                    ' Translate is
                    ' [ 1  0  tx ]
                    ' [ 0  1  ty ]
                    ' [ 0  0  1  ]
                    
                    If UBound(pSplit) = 0 Then
                        multiplyLineByMatrix lineID, 1, 0, 0, 1, Val(pSplit(0)), 0
                    Else
                        multiplyLineByMatrix lineID, 1, 0, 0, 1, Val(pSplit(0)), Val(pSplit(1))
                    End If
            
                Case "matrix"
                    pSplit = Split(params, ",")
                    If UBound(pSplit) = 0 Then pSplit = Split(params, " ")
                    multiplyLineByMatrix lineID, Val(pSplit(0)), Val(pSplit(1)), Val(pSplit(2)), Val(pSplit(3)), Val(pSplit(4)), Val(pSplit(5))
            
                Case "rotate"
                    
                    pSplit = Split(params, ",")
                    Ang = Deg2Rad(Val(pSplit(0)))
                    
                    multiplyLineByMatrix lineID, Cos(Ang), Sin(Ang), -Sin(Ang), Cos(Ang), 0, 0
                    
                Case "scale" ' scale(-1,-1)
                    pSplit = Split(params, ",")
                    If UBound(pSplit) = 0 Then pSplit = Split(params, " ")
                    If UBound(pSplit) = 0 Then
                        ' Handle shitty SVG, such as not having two parameters
                        ReDim Preserve pSplit(1)
                        pSplit(1) = pSplit(0)
                    End If
                    multiplyLineByMatrix lineID, Val(pSplit(0)), 0, 0, Val(pSplit(1)), 0, 0
                    
                
            End Select
        
        End If
    End With
    
End Function

Function multiplyLineByMatrix(polyID As Long, A As Double, b As Double, c As Double, D As Double, e As Double, f As Double)
    ' Miltiply a line/poly by a transformation matrix
    ' [ A C E ]
    ' [ B D F ]
    ' [ 0 0 1 ]
    
    ' http://www.w3.org/TR/SVG11/coords.html#TransformMatrixDefined
    'X1 = AX + CY + E
    'Y1 = BX + DY + F
    Dim j As Long
    Dim oldPoint As pointD
    
    With pData(polyID)
        For j = 1 To UBound(.Points)
            oldPoint = .Points(j)
            .Points(j).x = (A * oldPoint.x) + (c * oldPoint.y) + e
            .Points(j).y = (b * oldPoint.x) + (D * oldPoint.y) + f
        Next
    End With
    
End Function

Function parsePolyLine(inLine As String)
    ' Parse a polyline
    Dim pos As Long
    Dim char As String
    Dim token1 As String, token2 As String, token3 As String, token4 As String
    Dim currX As Double
    Dim currY As Double
    inLine = Replace(inLine, vbCr, " ")
    inLine = Replace(inLine, vbLf, " ")
    
    pos = 1
    Do Until pos > Len(inLine)
        skipWhiteSpace inLine, pos
        token1 = extractToken(inLine, pos)
        skipWhiteSpace inLine, pos
        token2 = extractToken(inLine, pos)
    
        If token1 <> "" And token2 <> "" Then
            addPoint Val(token1), Val(token2)
        End If
    Loop
        
        
    ' Close the shape.
    If UBound(pData(currentLine).Points) > 0 Then addPoint (pData(currentLine).Points(1).x), (pData(currentLine).Points(1).y)
    
    
End Function

Function parsePath(inPath As String, currentLayer As String)

    

    
    ' Parse an SVG path.
    Dim pos As Long
    Dim char As String
    Dim token1 As String, token2 As String, token3 As String, token4 As String
    Dim token5 As String, token6 As String, token7 As String, token8 As String
    
    
    Dim isRelative As Boolean
    Dim gotFirstItem As Boolean
    
    Dim currX As Double
    Dim currY As Double
    
    Dim pt0 As pointD
    Dim pt1 As pointD
    Dim pt2 As pointD
    Dim pt3 As pointD
    Dim pt4 As pointD
    Dim pt5 As pointD
    
    Dim ptPrevPoint As pointD
    Dim hasPrevPoint As Boolean
    
    Dim lastUpdate As Long
    
    
    
    
    
    Dim startX As Double
    Dim startY As Double
    
    Dim pInSeg As Double
    Dim lastChar As String
    
    
    
    'M209.1,187.65c-0.3-0.2-0.7-0.4-1-0.4c-0.3,0-0.7,0.2-0.9,0.4c-0.3,0.3-0.4,0.6-0.4,0.9c0,0.4,0.1,0.7,0.4,1
    'c0.2,0.2,0.6,0.4,0.9,0.4c0.3,0,0.7-0.2,1-0.4c0.2-0.3,0.3-0.6,0.3-1C209.4,188.25,209.3,187.95,209.1,187.65z

    ' Get rid of enter presses
    inPath = Replace(inPath, vbCr, " ")
    inPath = Replace(inPath, vbLf, " ")
    inPath = Replace(inPath, vbTab, " ")
    
    ' Start parsing
    pos = 1
    Do Until pos > Len(inPath)
        char = Mid(inPath, pos, 1)
        pos = pos + 1
        isRelative = False
        
        Select Case char
            Case "M", "m", "L", "l", "C", "c", "V", "v", "A", "a", "H", "h", "S", "s", "Z", "z", "q", "Q", "T", "t"
                ' Accepted character.
                lastChar = char
            Case " "
            
            Case Else
                ' No accepted, must be a continuation.
                char = lastChar
                If char = "m" Then char = "l" ' Continuous moveto becomes lineto
                If char = "M" Then char = "L" ' Continuous moveto becomes lineto not relative
                pos = pos - 1
        End Select
        
        
        Select Case char
            Case " " ' Skip spaces
            
            Case "M", "m"   ' MOVE TO
                If LCase(char) = char Then isRelative = True    ' Lowercase means relative co-ordinates
                If Not gotFirstItem Then isRelative = False 'Relative not valid for first item
                
                
                ' Extract two co-ordinates
                skipWhiteSpace inPath, pos
                token1 = extractToken(inPath, pos)
                skipWhiteSpace inPath, pos
                token2 = extractToken(inPath, pos)
                
                ' Set our "current" co-ordinates to this
                If isRelative Then
                    currX = currX + Val(token1)
                    currY = currY + Val(token2)
                Else
                    currX = Val(token1)
                    currY = Val(token2)
                End If
            
                ' Start a new line, since we moved
                'If Not isRelative Then
                newLine currentLayer
                'pData(currentLine).PathCode = Right(inPath, Len(inPath) - pos)
                    
                ' Add the start point to this line
                addPoint currX, currY
                
                
                'pData(currentLine).PathCode = pData(currentLine).PathCode & "Move to " & currX & ", " & currY & vbCrLf
                
                
                'If Not gotFirstItem Then
                startX = currX: startY = currY
                gotFirstItem = True
                hasPrevPoint = False
                
            Case "L", "l"   ' LINE TO
                If LCase(char) = char Then isRelative = True    ' Lowercase means relative co-ordinates
                If Not gotFirstItem Then isRelative = False 'Relative not valid for first item
                        
                        
                ' Extract two co-ordinates
                skipWhiteSpace inPath, pos
                token1 = extractToken(inPath, pos)
                skipWhiteSpace inPath, pos
                token2 = extractToken(inPath, pos)
                
                ' Set our "current" co-ordinates to this
                If isRelative Then
                    currX = currX + Val(token1)
                    currY = currY + Val(token2)
                Else
                    currX = Val(token1)
                    currY = Val(token2)
                End If
    
                ' Add this point to the line
                addPoint currX, currY
                
                ''pData(currentLine).PathCode = pData(currentLine).PathCode & "Line to " & currX & ", " & currY & vbCrLf
                
                If Not gotFirstItem Then startX = currX: startY = currY
                gotFirstItem = True
                hasPrevPoint = False
                
            Case "V", "v"   ' VERTICAL LINE TO
                If LCase(char) = char Then isRelative = True    ' Lowercase means relative co-ordinates
                If Not gotFirstItem Then isRelative = False 'Relative not valid for first item
                        
                ' Extract one co-ordinate
                skipWhiteSpace inPath, pos
                token1 = extractToken(inPath, pos)
                
                ' Set our "current" co-ordinates to this
                If isRelative Then
                    currY = currY + Val(token1)
                Else
                    currY = Val(token1)
                End If
    
                ' Add this point to the line
                addPoint currX, currY
                
                'pData(currentLine).PathCode = pData(currentLine).PathCode & "Vertical to " & currX & ", " & currY & vbCrLf
                
                If Not gotFirstItem Then startX = currX: startY = currY
                gotFirstItem = True
                hasPrevPoint = False
                
            Case "H", "h"   ' HORIZONTAL LINE TO
                If LCase(char) = char Then isRelative = True    ' Lowercase means relative co-ordinates
                If Not gotFirstItem Then isRelative = False 'Relative not valid for first item
                        
                ' Extract one co-ordinate
                skipWhiteSpace inPath, pos
                token1 = extractToken(inPath, pos)
                
                ' Set our "current" co-ordinates to this
                If isRelative Then
                    currX = currX + Val(token1)
                Else
                    currX = Val(token1)
                End If
    
                ' Add this point to the line
                addPoint currX, currY
                'pData(currentLine).PathCode = pData(currentLine).PathCode & "Horiz to " & currX & ", " & currY & vbCrLf
                
                If Not gotFirstItem Then startX = currX: startY = currY
                gotFirstItem = True
                hasPrevPoint = False
            
            Case "A", "a"       ' PARTIAL ARC TO
                If LCase(char) = char Then isRelative = True    ' Lowercase means relative co-ordinates
                If Not gotFirstItem Then isRelative = False 'Relative not valid for first item
            
                    '(rx ry x-axis-rotation large-arc-flag sweep-flag x y)+
                
                ' Radii X and Y
                skipWhiteSpace inPath, pos
                token1 = extractToken(inPath, pos)
                skipWhiteSpace inPath, pos
                token2 = extractToken(inPath, pos)
                
                ' X axis rotation
                skipWhiteSpace inPath, pos
                token3 = extractToken(inPath, pos)
                
                ' Large arc flag
                skipWhiteSpace inPath, pos
                token4 = extractToken(inPath, pos)
                
                ' Sweep flag
                skipWhiteSpace inPath, pos
                token5 = extractToken(inPath, pos)
                
                ' X and y
                skipWhiteSpace inPath, pos
                token6 = extractToken(inPath, pos)
                skipWhiteSpace inPath, pos
                token7 = extractToken(inPath, pos)
                
                ' Start point
                pt0.x = currX
                pt0.y = currY
                
                ' Set our "current" co-ordinates to this
                If isRelative Then
                    currX = currX + Val(token6)
                    currY = currY + Val(token7)
                Else
                    currX = Val(token6)
                    currY = Val(token7)
                End If
                
                pt1.x = currX
                pt1.y = currY
                
                parseArcSegment Val(token1), Val(token2), Val(token3), pt0, pt1, (token4 = "1"), (token5 = "1")
                
                'pData(currentLine).PathCode = pData(currentLine).PathCode & "Partial Arc to " & currX & ", " & currY & vbCrLf
                
                If Not gotFirstItem Then startX = currX: startY = currY
                gotFirstItem = True
                hasPrevPoint = False
                
            Case "C", "c"       ' CURVE TO
                If LCase(char) = char Then isRelative = True    ' Lowercase means relative co-ordinates
                If Not gotFirstItem Then isRelative = False 'Relative not valid for first item
               
                pt0.x = currX
                pt0.y = currY
                
                ' Extract two co-ordinates
                skipWhiteSpace inPath, pos
                token1 = extractToken(inPath, pos)
                skipWhiteSpace inPath, pos
                token2 = extractToken(inPath, pos)
                               
                ' Set into point 0
                pt1.x = IIf(isRelative, currX, 0) + Val(token1)
                pt1.y = IIf(isRelative, currY, 0) + Val(token2)
                
                
                ' Extract next two co-ordinates
                skipWhiteSpace inPath, pos
                token1 = extractToken(inPath, pos)
                skipWhiteSpace inPath, pos
                token2 = extractToken(inPath, pos)
                               
                ' Set into point 1
                pt2.x = IIf(isRelative, currX, 0) + Val(token1)
                pt2.y = IIf(isRelative, currY, 0) + Val(token2)
                
                ' Extract next two co-ordinates
                skipWhiteSpace inPath, pos
                token1 = extractToken(inPath, pos)
                skipWhiteSpace inPath, pos
                token2 = extractToken(inPath, pos)
                               
                ' Set into point 2
                currX = IIf(isRelative, currX, 0) + Val(token1)
                currY = IIf(isRelative, currY, 0) + Val(token2)
                pt3.x = currX
                pt3.y = currY
                
'
                pInSeg = getPinSeg(pt0, pt3)
                
                
                
                ' Run the bezier code with 4 points
                AddBezier pInSeg, pt0, pt1, pt2, pt3
                
                ' Reflect this point about pt3
                
                ptPrevPoint = reflectAbout(pt2, pt3)
                hasPrevPoint = True
                
                'pData(currentLine).PathCode = pData(currentLine).PathCode & "Bezier to " & currX & ", " & currY & vbCrLf
                
                If Not gotFirstItem Then startX = currX: startY = currY
                gotFirstItem = True
                
            Case "S", "s"      ' CURVE TO with 3 points
                If LCase(char) = char Then isRelative = True    ' Lowercase means relative co-ordinates
                If Not gotFirstItem Then isRelative = False 'Relative not valid for first item
               
                pt0.x = currX
                pt0.y = currY
                
                ' Extract two co-ordinates
                skipWhiteSpace inPath, pos
                token1 = extractToken(inPath, pos)
                skipWhiteSpace inPath, pos
                token2 = extractToken(inPath, pos)
                               
                ' Set into point 0
                pt1.x = IIf(isRelative, currX, 0) + Val(token1)
                pt1.y = IIf(isRelative, currY, 0) + Val(token2)
                
                ' Extract next two co-ordinates
                skipWhiteSpace inPath, pos
                token1 = extractToken(inPath, pos)
                skipWhiteSpace inPath, pos
                token2 = extractToken(inPath, pos)
                               
                ' Set into point 1
                currX = IIf(isRelative, currX, 0) + Val(token1)
                currY = IIf(isRelative, currY, 0) + Val(token2)
                pt2.x = currX
                pt2.y = currY
                
                pInSeg = getPinSeg(pt0, pt2)
                
                
                If Not hasPrevPoint Then
                    ' Same as pt1
                    ptPrevPoint = pt1
                End If
                
                AddBezier pInSeg, pt0, ptPrevPoint, pt1, pt2
            
                ptPrevPoint = reflectAbout(pt1, pt2)
                hasPrevPoint = True
                
            
                'pData(currentLine).PathCode = pData(currentLine).PathCode & "3Bezier to " & currX & ", " & currY & vbCrLf
            
                If Not gotFirstItem Then startX = currX: startY = currY
                gotFirstItem = True
                
            Case "Q", "q"      ' Quadratic Bezier TO with 3 points
                If LCase(char) = char Then isRelative = True    ' Lowercase means relative co-ordinates
                If Not gotFirstItem Then isRelative = False 'Relative not valid for first item
               
                pt0.x = currX
                pt0.y = currY
                
                ' Extract two co-ordinates
                skipWhiteSpace inPath, pos
                token1 = extractToken(inPath, pos)
                skipWhiteSpace inPath, pos
                token2 = extractToken(inPath, pos)
                               
                ' Set into point 0
                pt1.x = IIf(isRelative, currX, 0) + Val(token1)
                pt1.y = IIf(isRelative, currY, 0) + Val(token2)
                
                ' Extract next two co-ordinates
                skipWhiteSpace inPath, pos
                token1 = extractToken(inPath, pos)
                skipWhiteSpace inPath, pos
                token2 = extractToken(inPath, pos)
                               
                ' Set into point 1
                currX = IIf(isRelative, currX, 0) + Val(token1)
                currY = IIf(isRelative, currY, 0) + Val(token2)
                pt2.x = currX
                pt2.y = currY
                
                pInSeg = getPinSeg(pt0, pt2)
                
                
                'If Not hasPrevPoint Then
                '    ' Same as pt1
                '    ptPrevPoint = pt1
                'End If
                
                AddQuadBezier pInSeg, pt0, pt1, pt2
            
                ptPrevPoint = reflectAbout(pt1, pt2)
                hasPrevPoint = True
                
                'pData(currentLine).PathCode = pData(currentLine).PathCode & "3Bezier to " & currX & ", " & currY & vbCrLf
            
                If Not gotFirstItem Then startX = currX: startY = currY
                gotFirstItem = True
                                
            Case "T", "t"      ' Quadratic Bezier TO with 3 points, but use reflection of last
                If LCase(char) = char Then isRelative = True    ' Lowercase means relative co-ordinates
                If Not gotFirstItem Then isRelative = False 'Relative not valid for first item
               
                pt0.x = currX
                pt0.y = currY
                
                ' Extract two co-ordinates
                skipWhiteSpace inPath, pos
                token1 = extractToken(inPath, pos)
                skipWhiteSpace inPath, pos
                token2 = extractToken(inPath, pos)
                               
                ' Set into point 0
                pt1.x = IIf(isRelative, currX, 0) + Val(token1)
                pt1.y = IIf(isRelative, currY, 0) + Val(token2)
                
                pInSeg = getPinSeg(pt0, pt1)
                
                
                
                If Not hasPrevPoint Then
                    ' Same as pt1
                    ptPrevPoint = pt0 ' SHOULD NEVER HAPPEN
                End If
                
                AddQuadBezier pInSeg, pt0, ptPrevPoint, pt1
            
                ptPrevPoint = reflectAbout(ptPrevPoint, pt1)
                hasPrevPoint = True
                
                'pData(currentLine).PathCode = pData(currentLine).PathCode & "3Bezier to " & currX & ", " & currY & vbCrLf
            
                If Not gotFirstItem Then startX = currX: startY = currY
                gotFirstItem = True
                                                
            Case "z", "Z"
            
                hasPrevPoint = False
                
                ' z means end the shape
                ' Draw a line back to start of shape
                addPoint startX, startY
                currX = startX
                currY = startY
                
                
                ' Since this is a closed path, mark it as fillable.
                pData(currentLine).Fillable = True
                
                'gotFirstItem = False
                
                
                'pData(currentLine).PathCode = pData(currentLine).PathCode & "End Shape" & vbCrLf
            
                
            
            Case Else
                Debug.Print "UNSUPPORTED PATH CODE: ", char
                
            
        End Select
        
        
        If pos > lastUpdate + 2000 Then
            lastUpdate = pos
            frmInterface.Caption = "Parsing path: " & pos & " / " & Len(inPath)
            DoEvents
        End If
        
    Loop
    
    
    

End Function

Function getPinSeg(pStart As pointD, pEnd As pointD)
    Dim D As Double
    D = pointDistance(pStart, pEnd) / GLOBAL_DPI
    'MsgBox "distance: " & D
    
    'Select Case d
    '    Case Is > 20
    '        getPinSeg = 0.1
    '    Case Is > 10
    '        getPinSeg = 0.2
    '    Case Is > 5
    '        getPinSeg = 0.25
    '    Case Else
    '        getPinSeg = 0.3
    'End Select
               
               
    ' with a resolution of 500 dpi, the curve should be split into 500 segments per inch. so a distance of 1 should be 500 segments, which is 0.002
    Dim segments As Double
    segments = 250 * D
    
    If segments = 0 Then segments = 1
    
    
    getPinSeg = Max(0.01, 1 / segments)
    
               

End Function




Function reflectAbout(ptReflect As pointD, ptOrigin As pointD) As pointD
    ' Reflect ptReflect 180 degrees around ptOrigin
    
    
    reflectAbout.x = (-(ptReflect.x - ptOrigin.x)) + ptOrigin.x
    reflectAbout.y = (-(ptReflect.y - ptOrigin.y)) + ptOrigin.y
    
    
End Function

Function parseArcSegment(RX As Double, RY As Double, rotAng As Double, _
                            P1 As pointD, P2 As pointD, _
                            largeArcFlag As Boolean, sweepFlag As Boolean)
    
    ' Parse "A" command in SVG, which is segments of an arc
    ' P1 is start point
    ' P2 is end point
        
    Dim centerPoint As pointD
    Dim Theta As Double
    Dim P1Prime As pointD
    Dim P2Prime As pointD
    
    Dim CPrime As pointD
    Dim Q As Double
    Dim qTop As Double
    Dim qBot As Double
    Dim c As Double
    
    Dim startAng As Double
    Dim endAng As Double
    Dim Ang As Double
    Dim AngStep As Double
    
    Dim tempPoint As pointD
    Dim tempAng As Double
    Dim tempDist As Double
    
    
    
    Dim Theta1 As Double
    Dim ThetaDelta As Double
    
    
    ' Turn the degrees of rotation into radians
    Theta = Deg2Rad(rotAng)
        
    ' Calculate P1Prime
    P1Prime.x = (Cos(Theta) * ((P1.x - P2.x) / 2)) + (Sin(Theta) * ((P1.y - P2.y) / 2))
    P1Prime.y = (-Sin(Theta) * ((P1.x - P2.x) / 2)) + (Cos(Theta) * ((P1.y - P2.y) / 2))
    
    P2Prime.x = (Cos(Theta) * ((P2.x - P1.x) / 2)) + (Sin(Theta) * ((P2.y - P1.y) / 2))
    P2Prime.y = (-Sin(Theta) * ((P2.x - P1.x) / 2)) + (Cos(Theta) * ((P2.y - P1.y) / 2))
    
    qTop = ((RX ^ 2) * (RY ^ 2)) - ((RX ^ 2) * (P1Prime.y ^ 2)) - ((RY ^ 2) * (P1Prime.x ^ 2))
    
    If qTop < 0 Then ' We've been given an invalid arc. Calculate the correct value.
        
        c = Sqr(((P1Prime.y ^ 2) / (RY ^ 2)) + ((P1Prime.x ^ 2) / (RX ^ 2)))
        
        RX = RX * c
        RY = RY * c
        
        qTop = 0
    End If
    
    qBot = ((RX ^ 2) * (P1Prime.y ^ 2)) + ((RY ^ 2) * (P1Prime.x ^ 2))
    If qBot <> 0 Then
    Q = Sqr((qTop) / (qBot))
    Else
        Q = 0
    End If
    ' Q is negative
    If largeArcFlag = sweepFlag Then Q = -Q
    
    ' Calculate Center Prime
    CPrime.x = 0
    
    If RY <> 0 Then CPrime.x = Q * ((RX * P1Prime.y) / RY)
    If RX <> 0 Then CPrime.y = Q * -((RY * P1Prime.x) / RX)
    
    ' Calculate center point
    centerPoint.x = ((Cos(Theta) * CPrime.x) - (Sin(Theta) * CPrime.y)) + ((P1.x + P2.x) / 2)
    centerPoint.y = ((Sin(Theta) * CPrime.x) + (Cos(Theta) * CPrime.y)) + ((P1.y + P2.y) / 2)
    
    ' TEMPTEMP
    
    frmInterface.Zoom = 2
    frmInterface.panX = 140
    frmInterface.panY = 140
    
    
    frmInterface.Picture1.Circle ((centerPoint.x + frmInterface.panX) * frmInterface.Zoom, (centerPoint.y + frmInterface.panY) * frmInterface.Zoom), 10, vbBlue
    frmInterface.Picture1.Circle ((P1.x + frmInterface.panX) * frmInterface.Zoom, (P1.y + frmInterface.panY) * frmInterface.Zoom), 10, vbGreen
    frmInterface.Picture1.Circle ((P2.x + frmInterface.panX) * frmInterface.Zoom, (P2.y + frmInterface.panY) * frmInterface.Zoom), 10, vbRed
    
    Debug.Print "Circle"
    
    ' Calculate Theta1
    
    Theta1 = angleFromPoint(P1Prime, CPrime)
    ThetaDelta = angleFromPoint(P2Prime, CPrime)
    
    Theta1 = Theta1 - PI
    ThetaDelta = ThetaDelta - PI
    
    'Theta1 = angleFromVect(((P1Prime.X - CPrime.X) / RX), ((P1Prime.Y - CPrime.Y) / RY), (P1Prime.X - CPrime.X), (P1Prime.Y - CPrime.Y))
    'ThetaDelta = angleFromVect(((-P1Prime.X - CPrime.X) / RX), ((-P1Prime.Y - CPrime.Y) / RY), (-P1Prime.X - CPrime.X), (-P1Prime.Y - CPrime.Y))
    
    'Theta1 = Theta1 - (PI / 2)
    'ThetaDelta = ThetaDelta - (PI / 2)

    'If Theta1 = ThetaDelta Then ThetaDelta = ThetaDelta + (PI * 2)
    
    'Debug.Print Theta1
        
    
    If sweepFlag Then ' Sweep is going POSITIVELY
        If ThetaDelta < Theta1 Then ThetaDelta = ThetaDelta + (PI * 2)
    Else    ' Sweep  is going NEGATIVELY
        'If ThetaDelta < 0 Then ThetaDelta = ThetaDelta + (PI * 2)
        If ThetaDelta > Theta1 Then ThetaDelta = ThetaDelta - (PI * 2)
    End If
    
    
    startAng = Theta1
    endAng = ThetaDelta
    
    
    AngStep = (PI / 180)
    If Not sweepFlag Then AngStep = -AngStep ' Sweep flag indicates a positive step
    
    Debug.Print "Start angle", Rad2Deg(startAng), " End angle ", Rad2Deg(endAng), "Step ", Rad2Deg(AngStep)
    
    'Theta = Deg2Rad(-40)
    
    ' Hackhack
    'startAng = startAng + AngStep * 2
    
    
    Ang = startAng
    Do
        ' X   =   RX
        'pt4.X = (pt1.X * Cos(Ang))
        'pt4.Y = (pt1.Y * Sin(Ang))

        'pt4.X = (Cos(Theta) * pt4.X) + (-Sin(Theta) * pt4.Y)
        'pt4.Y = (Sin(Theta) * pt4.X) + (Cos(Theta) * pt4.Y)

        '         X      CX
        'pt4.X = pt4.X + pt3.X
        'pt4.Y = pt4.Y + pt3.Y

        tempPoint.x = (RX * Cos(Ang)) + centerPoint.x
        tempPoint.y = (RY * Sin(Ang)) + centerPoint.y
        
        tempAng = angleFromPoint(centerPoint, tempPoint) + Theta
        tempDist = pointDistance(centerPoint, tempPoint)
        
        tempPoint.x = (tempDist * Cos(tempAng)) + centerPoint.x
        tempPoint.y = (tempDist * Sin(tempAng)) + centerPoint.y
        
        
        
        
        
        'tempPoint.X = (Cos(Theta) * tempPoint.X) + (-Sin(Theta) * tempPoint.Y)
        'tempPoint.Y = (Sin(Theta) * tempPoint.X) + (Cos(Theta) * tempPoint.Y)
        

        addPoint tempPoint.x, tempPoint.y
        

        Ang = Ang + AngStep
    Loop Until (Ang >= endAng And AngStep > 0) Or (Ang <= endAng And AngStep < 0)

    ' Add the final point

    addPoint P2.x, P2.y
    
    
End Function

Function rotatePoint(inPoint As pointD, Theta As Double, centerPoint As pointD) As pointD

    rotatePoint = inPoint
    
    rotatePoint.x = rotatePoint.x - centerPoint.x
    rotatePoint.y = rotatePoint.y - centerPoint.y
    
    rotatePoint.x = (Cos(Theta) * rotatePoint.x) + (-Sin(Theta) * rotatePoint.y)
    rotatePoint.y = (Sin(Theta) * rotatePoint.x) + (Cos(Theta) * rotatePoint.y)
    
    rotatePoint.x = rotatePoint.x + centerPoint.x
    rotatePoint.y = rotatePoint.y + centerPoint.y
    
    

End Function


Function Rad2Deg(inRad As Double) As Double
    Rad2Deg = inRad * (180 / PI)
End Function

Function Deg2Rad(inDeg As Double) As Double
    Deg2Rad = inDeg / (180 / PI)
End Function

Function angleFromVect(vTop As Double, vBot As Double, diffX As Double, diffY As Double) As Double
    ' Not sure if this working
    
    If vBot = 0 Then
        angleFromVect = IIf(vTop > 0, PI / 2, -PI / 2)
    ElseIf diffX >= 0 Then
        angleFromVect = Atn(vTop / vBot)
    Else
        angleFromVect = Atn(vTop / vBot) - PI
    End If

End Function

Function angleFromPoint(pCenter As pointD, pPoint As pointD) As Double
    ' Calculate the angle of a point relative to the center
    
    ' Slope is rise over run
    Dim slope As Double
    
    If pPoint.x = pCenter.x Then
        ' Either 90 or 270
        angleFromPoint = IIf(pPoint.y > pCenter.y, PI / 2, -PI / 2)
        
    ElseIf pPoint.x > pCenter.x Then
        ' 0 - 90 and 270-360
        slope = (pPoint.y - pCenter.y) / (pPoint.x - pCenter.x)
        angleFromPoint = Atn(slope)
    Else
        ' 180-270
        slope = (pPoint.y - pCenter.y) / (pPoint.x - pCenter.x)
        angleFromPoint = Atn(slope) + PI
    End If
    
    If angleFromPoint < 0 Then angleFromPoint = angleFromPoint + (PI * 2)
    
    
    
    
End Function

Function newLine(Optional theLayer As String)
    
    If hasUnfinishedLine Then finishLine
    
    
    
    currentLine = UBound(pData) + 1
    ' Set up this line
    ReDim Preserve pData(currentLine)
    ReDim pData(currentLine).Points(0)

    pData(currentLine).LayerID = theLayer
    

End Function

Function finishLine()
    If hasUnfinishedLine Then
        hasUnfinishedLine = False
        
        ' Remove the excess
        ReDim Preserve pData(currentLine).Points(pData(currentLine).SpecialNumPoints)
    End If
    
End Function

Function addPoint(x As Double, y As Double, Optional noCutLineSegment As Boolean)
    
    Dim n As Long
    With pData(currentLine)
        
        If .Points(UBound(.Points)).x = x And .Points(UBound(.Points)).y = y And UBound(.Points) > 0 Then
            ' No point to add
            'Debug.Print "same as last point"
            
        Else
        
            ' Once we get over 5000 points, we enter a special allocation mode.
            If UBound(.Points) > 5000 Then
                hasUnfinishedLine = True
                
                ' Allocate in blocks of 5000 at a time.
                n = .SpecialNumPoints + 1
                If n > UBound(.Points) Then ReDim Preserve .Points(UBound(.Points) + 5000)
                
            Else
                n = UBound(.Points) + 1
                ReDim Preserve .Points(n)
            End If
            
        
            .Points(n).x = x
            .Points(n).y = y
            .SpecialNumPoints = n
            If noCutLineSegment Then .Points(n).noCut = 1
        End If
    End With
    

End Function

Function skipWhiteSpace(ByRef inPath As String, ByRef pos As Long)
    ' Skip any white space.
    Dim char As String
    
    Do Until pos > Len(inPath)
        char = Mid(inPath, pos, 1)
        Select Case char
            Case " ", ",", vbTab ' List all white space characters here
                ' Continue
            Case Else
                Exit Function
        End Select
                
        pos = pos + 1
    Loop
End Function


Function extractToken(ByRef inPath As String, ByRef pos As Long) As String

    ' Exract until we get a space or a comma
    Dim char As String
    Dim build As String
    Dim seenMinus As Boolean
    Dim startPos As Long
    Dim seenE As Boolean
    
    startPos = pos
    
    
    Do Until pos > Len(inPath)
        char = Mid(inPath, pos, 1)
        
        Select Case char
            ' Only accept numbers
            Case "-"
                If seenE Then
                    build = build & char
                    pos = pos + 1
                ElseIf seenMinus Or pos > startPos Then
                    Exit Do
                Else
                    ' We already saw a minus sign
                    seenMinus = True
                    build = build & char
                    pos = pos + 1
                End If
                
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "."
                build = build & char
                pos = pos + 1
                ',6.192 -10e-4,12.385
            Case "e" ' Exponent
                seenE = True
                build = build & char
                pos = pos + 1
            Case Else
                Exit Do
        End Select
    Loop
    extractToken = build

End Function

Function isNumChar(char As String) As Boolean
    Select Case char
        ' Only accept numbers
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "-", "."
            isNumChar = True
    End Select



End Function


Function getAttr(attr As ChilkatXml, attrName As String, Optional DefaultValue)

    getAttr = attr.GetAttrValue(attrName)

End Function

Function pointIsInPoly(polyID As Long, x As Double, y As Double)

    ' Determine if this point is inside the polygon.
    
    
    
    Dim i As Long
    Dim j As Long
    
    With pData(polyID)
        j = UBound(.Points)
        
        For i = 1 To UBound(.Points)
        
            If (.Points(i).y < y And .Points(j).y >= y _
                Or .Points(j).y < y And .Points(i).y >= y) Then
                    If (.Points(i).x + (y - .Points(i).y) / (.Points(j).y - .Points(i).y) * (.Points(j).x - .Points(i).x) < x) Then
                        pointIsInPoly = Not pointIsInPoly
                    End If
            End If
        
            j = i
        Next
        
    End With

'  int      i, j=polySides-1 ;
'  boolean  oddNodes=NO      ;
'
'  for (i=0; i<polySides; i++) {
'    if (polyY[i]<y && polyY[j]>=y
'    ||  polyY[j]<y && polyY[i]>=y) {
'      if (polyX[i]+(y-polyY[i])/(polyY[j]-polyY[i])*(polyX[j]-polyX[i])<x) {
'        oddNodes=!oddNodes; }}
'    j=i; }
'
'  return oddNodes; }
    
    
'    Dim nPol As Long
'    Dim i As Long, j As Long
'
'    Dim counter As Long
'
'    Dim p1 As pointD
'    Dim p2 As pointD
'    Dim p As pointD
'    Dim n As Long
'    Dim xinters As Double
'
'    p.X = X
'    p.Y = Y
'
'
'  'double xinters;
'  'Point p1,p2;
'    With pData(polyID)
'        n = UBound(.Points)
'        p1 = .Points(1)
'        For i = 1 To n
'            p2 = .Points(i Mod n)
'
'            If (p.Y > Min(p1.Y, p2.Y)) Then
'                If (p.Y <= Max(p1.Y, p2.Y)) Then
'                    If (p.X <= Max(p1.X, p2.X)) Then
'                        If (p1.Y <> p2.Y) Then
'                            xinters = (p.Y - p1.Y) * (p2.X - p1.X) / (p2.Y - p1.Y) + p1.X
'                            If (p1.X = p2.X Or p.X <= xinters) Then counter = counter + 1
'                        End If
'                    End If
'                End If
'            End If
'            p1 = p2
'        Next
'
'    End With
'
'    If counter Mod 2 = 0 Then
'        pointIsInPoly = False
'    Else
'        pointIsInPoly = True
'    End If
'
'
    
    
    
'    Dim Inside As Boolean
'
'    With pData(polyID)
'        nPol = UBound(.Points) ' Number of points
'
'        j = nPol ' Starts at the last point
'        For i = 1 To nPol
'            If .Points(j).Y - .Points(i).Y > 0 Then
'                If ((((.Points(i).Y <= Y) And (Y < .Points(j).Y)) Or _
'                    ((.Points(j).Y <= Y) And (Y < .Points(i).Y))) And _
'                    (X < (.Points(j).X - .Points(i).X) * (Y - .Points(i).Y) / (.Points(j).Y - .Points(i).Y) + .Points(i).X)) Then
'                        Inside = Not Inside
'                End If
'            End If
'            j = i
'        Next
'    End With
'
'    pointIsInPoly = Inside
    
'int pnpoly(int npol, float *xp, float *yp, float x, float y)
'    {
'      int i, j, c = 0;
'      for (i = 0, j = npol-1; i < npol; j = i++) {
'        if ((((yp[i] <= y) && (y < yp[j])) ||
'             ((yp[j] <= y) && (y < yp[i]))) &&
'            (x < (xp[j] - xp[i]) * (y - yp[i]) / (yp[j] - yp[i]) + xp[i]))
'          c = !c;
'      }
'      return c;
'    }


End Function

Function getPolyBounds(polyID As Long, ByRef minX As Double, ByRef minY As Double, ByRef maxX As Double, ByRef maxY As Double)

    Dim j As Long
    
    minX = 1000000
    minY = 1000000
    maxX = 0
    maxY = 0
    
    ' Calculate the extents
    With pData(polyID)
        For j = 1 To UBound(.Points)
            With .Points(j)
                minX = Min(minX, .x)
                minY = Min(minY, .y)
                maxX = Max(maxX, .x)
                maxY = Max(maxY, .y)
            End With
        Next
    End With


End Function

Function getExtents(ByRef maxX As Double, ByRef maxY As Double, Optional ByRef minX As Double, Optional ByRef minY As Double)

    Dim i As Long
    Dim j As Long
    Dim setMin As Boolean
        
    ' Calculate the extents
    For i = 1 To UBound(pData)
        With pData(i)
            For j = 1 To UBound(.Points)
                With .Points(j)
                    If setMin Then
                        minX = Min(minX, .x)
                        minY = Min(minY, .y)
                    Else
                        setMin = True
                        minX = .x
                        minY = .y
                    End If
                    maxX = Max(maxX, .x)
                    maxY = Max(maxY, .y)
                End With
            Next
        End With
    Next


End Function

Function canPolyFitInside(smallPoly As Long, bigPoly As Long)
    ' See if smallPoly will fit inside bigPoly
    
    ' In theory, if all of smallPoly's points are inside bigPoly, then the whole poly is inside bigpoly.
    Dim i As Long
    With pData(smallPoly)
        For i = 1 To UBound(.Points)
            With .Points(i)
                If Not pointIsInPoly(bigPoly, .x, .y) Then
                    ' This point is outside.
                    Exit Function
                Else
                    canPolyFitInside = True
                End If
            End With
        Next
    End With
    
    
        
End Function

Function getPolyArea(polyID As Long) As Double
    ' Get the area of this polygon
    Dim minX As Double, maxX As Double
    Dim minY As Double, maxY As Double
    
    getPolyBounds polyID, minX, minY, maxX, maxY
        
    ' For now, we are just using the bounding box. Todo: proper area calculation
    getPolyArea = (maxX - minX) * (maxY - minY)
    
End Function

Function pointIsInPolyWithContain(polyID As Long, x As Double, y As Double) As Boolean

    ' Checks if the point is or isn't in the poly and deals with contained poly's also
    Dim cl As Collection
    Dim i As Long
    Dim isIn As Boolean
    If containList.Exists(polyID) Then Set cl = containList(polyID) ' A list of polygons that I contain
    
    isIn = pointIsInPoly(polyID, x, y)
    
    ' Check if it's in any of my kids. If so, it could be that it's NOT inside me.
    If Not cl Is Nothing Then
        For i = 1 To cl.count
            If pointIsInPolyWithContain(cl(i), x, y) Then
                ' It's in my kid.
                Exit Function
            End If
        Next
    End If
    
    pointIsInPolyWithContain = isIn
    
    

End Function

Sub rasterDocument(yStep As Double, currentLayer As String)

    Dim maxX As Double, maxY As Double, minX As Double, minY As Double
    Dim p As Long
    Dim totalResult() As pointD
    Dim result() As pointD
    Dim n As Long
    Dim y As Double
    Dim i As Long
    Dim goingRight As Boolean
    
    getExtents maxX, maxY, minX, minY
        
    ' Here's how this works:
    ' We draw a line from left to right, and then right to left, through the entire document. All shapes.
    ' We create a giant list of all the places where it intersects.
    ' And we take that and create a single line with many on/off points.
    
    y = minY
    Do Until y >= maxY
        
        ReDim totalResult(0)
            
        For p = 1 To UBound(pData)
            If pData(p).ContainedBy = 0 And pData(p).Fillable Then
        
                
                
                ' Draw a line from the X left to the X right, and fill in every second line segment.
                result = lineIntersectPoly(newPoint(minX - 50, y), newPoint(maxX + 50, y), p)
        
        
                If UBound(result) > 0 Then
                    ' Copy into TotalResult
                    n = UBound(totalResult)
                    ReDim Preserve totalResult(n + UBound(result))
                    For i = 1 To UBound(result)
                        totalResult(n + i) = result(i)
                    Next
                End If
                
            End If
        Next
        
        If UBound(totalResult) > 0 Then

            newLine currentLayer
            
            orderArray totalResult, goingRight
            goingRight = Not goingRight ' TEMP
            
            i = 1
            ' Add a beginning point
            addPoint totalResult(i).x + IIf(goingRight, -0.5, 0.5), totalResult(i).y, True
            Do Until i > UBound(totalResult)
                ' Start point
                addPoint totalResult(i).x, totalResult(i).y, i Mod 2 = 0
                i = i + 1
            Loop
            ' And an end point
            addPoint totalResult(i - 1).x + IIf(goingRight, 0.5, -0.5), totalResult(i - 1).y, True

        End If
        
        y = y + yStep
        '    frmInterface.Caption = "Progress : " & Round(y / maxY * 100) & " %"
        '    DoEvents
        ''End If
            
            
    Loop




End Sub

Sub rasterLinePoly(lineID As Long, yStep As Double, currentLayer As String)

    ' Fill this polygon with raster lines from top to bottom
    
    Dim maxX As Double, maxY As Double
    Dim minX As Double, minY As Double
    Dim x As Double, y As Double
    Dim prevX As Double
    Dim Xadd As Double
    Dim result() As pointD
    Dim draw() As pointD
    
    Dim i As Long
    Dim j As Long
    Dim lastPoint As pointD
    Dim cap As String
    cap = frmInterface.Caption
    
    
    
    Dim goingRight As Boolean ' The laser moves either left or right. Alternate directions smartly.
    
    'yStep = 0.008

    ' Get the bounds of this shape.
    
    getPolyBounds lineID, minX, minY, maxX, maxY
    
    y = minY
    Do Until y >= maxY
        
        ' Draw a line from the X left to the X right, and fill in every second line segment.
        result = lineIntersectPoly(newPoint(-10, y), newPoint(maxX + 50, y), lineID)

        If UBound(result) > 0 Then

            orderArray result, goingRight
            goingRight = Not goingRight
            i = 1
            Do Until i > UBound(result)
            

                ' Start point
                If i + 1 <= UBound(result) Then
                    newLine currentLayer
                    addPoint result(i).x, result(i).y
                    addPoint result(i + 1).x, result(i + 1).y
                End If
                
                i = i + 2
            Loop
        End If
        'TEMP
        'yStep = yStep * 1.05
        y = y + yStep
        
        'If CLng(Y) Mod 10 = 0 Then
            'frmInterface.Caption = "Progress : " & Round(y / maxY * 100) & " %"
            DoEvents
        'End If
        
    Loop
    
End Sub

Function lineThroughPolygon(polyID As Long, startPoint As pointD, endPoint As pointD) As pointD()

    ' Return an array of line segments to draw with this line
    Dim out() As pointD
    Dim draw() As pointD
    Dim result() As pointD
    Dim i As Long
    Dim j As Long, K As Long, k2 As Long
    Dim cl As Collection
    If containList.Exists(polyID) Then Set cl = containList(polyID) ' A list of polygons that I contain
    
    
    result = lineIntersectPoly(startPoint, endPoint, polyID)
    
    
    If UBound(result) = 0 Then 'No intersections
    
        ' Return just the segment unchanged
        ReDim out(1)
        out(0) = startPoint
        out(1) = endPoint
    Else
        ' Build a new set of lines based on the result.
        
        ' Order the points from left to right
        orderArray result, True
        
        ' THIS array should be odd!
        ReDim out(0)
        
        out(0) = startPoint
        For i = 1 To UBound(result) Step 2
            If i + 1 <= UBound(result) Then
                
                ' Check the kids of this shape.
                If Not cl Is Nothing Then
                    For K = 1 To 1 'cl.count
                        draw = lineThroughPolygon(cl(K), result(i), result(i + 1))
                        
                        ' Add this
                        For k2 = 0 To UBound(draw) Step 2
                            If k2 + 1 <= UBound(draw) Then
                                ReDim Preserve out(UBound(out) + 2)
                                out(UBound(out) - 1) = draw(k2)
                                out(UBound(out)) = draw(k2 + 1)
                            End If
                        Next
                    Next
                Else
                    ' Add two points
                    ReDim Preserve out(UBound(out) + 2)
                    out(UBound(out) - 1) = result(i)
                    out(UBound(out)) = result(i + 1)
                End If
                
                
            End If
        Next
        ' Last point
        ReDim Preserve out(UBound(out) + 1)
        out(UBound(out)) = endPoint
    End If
    
    
    lineThroughPolygon = out

End Function

Function orderArray(inRes() As pointD, Ascending As Boolean)

    ' Order the return array of points.
    Dim i As Long
    Dim b As Double
    Dim sorted As Boolean
    Do
        sorted = False
        For i = 1 To UBound(inRes) - 1
            
            If (inRes(i).x > inRes(i + 1).x And Not Ascending) Or (inRes(i).x < inRes(i + 1).x And Ascending) Then
                ' swap
                b = inRes(i).x
                inRes(i).x = inRes(i + 1).x
                inRes(i + 1).x = b
                sorted = True
            End If
        Next
    Loop Until Not sorted
    
End Function

Function sortByLayers()

    Dim i As Long
    Dim sorted As Boolean
    Dim bb As typLine
    
    Do
        sorted = False
        For i = 1 To UBound(pData) - 1
            If pData(i).LayerID > pData(i + 1).LayerID Then
                sorted = True
                bb = pData(i + 1)
                pData(i + 1) = pData(i)
                pData(i) = bb
                
            End If
        Next
    Loop Until sorted = False
    
End Function

Function mergeConnectedLines()
    
    Dim i As Long, j As Long
    Dim n As Long
    Dim iCount As Long
    Dim doMerge As Boolean
    Dim doFlip As Boolean
    Dim didMerge As Boolean
    
    ' Looks for polygons that begin/end exactly at the beginning/end of another polygon and merges them into one polygon.
    
    For i = 1 To UBound(pData)
        pData(i).Optimized = False
    Next
    
    ' Step 2: Loop through the unoptimized polygons
    Do
        didMerge = False
        For i = 1 To UBound(pData) - 1
            
            
            If Not pData(i).Optimized Then
                iCount = UBound(pData(i).Points)
                
                frmInterface.Caption = "Optimizing " & i & " / " & UBound(pData)
                If i Mod 50 = 0 Then DoEvents
                
                doMerge = False
                For j = 1 To UBound(pData)
                    If j <> i And pData(j).LayerID = pData(i).LayerID Then
                        If pData(i).Points(iCount).x = pData(j).Points(1).x And _
                           pData(i).Points(iCount).y = pData(j).Points(1).y Then
                            
                            ' OK, this shape starts where my shape ends.
                            Debug.Print "SHAPE " & i & " AND " & j & " X: ", pData(i).Points(iCount).x, pData(j).Points(1).x
                            Debug.Print "SHAPE " & i & " AND " & j & " Y: ", pData(i).Points(iCount).y, pData(j).Points(1).y
                            
                            doMerge = True
                            doFlip = False
                            Exit For
                        End If
                            
                        If pData(i).Points(iCount).x = pData(j).Points(UBound(pData(j).Points)).x And _
                           pData(i).Points(iCount).y = pData(j).Points(UBound(pData(j).Points)).y Then
                            ' OK, this shape ends where my shape ends.
                            doMerge = True
                            doFlip = True ' Since its the end that matched, we need to flip it first.
                            Exit For
                        End If
                    End If
                Next
                
                If doMerge Then
                    Debug.Print "MERGING SHAPE ", j, "INTO ", i
                    didMerge = True
                    If doFlip Then ' Flip it around first.
                        flipPolyStartEnd j
                    End If
                    
                    ' Merge the points from j into i
                    ReDim Preserve pData(i).Points(iCount + UBound(pData(j).Points))
                    
                    For n = 1 To UBound(pData(j).Points)
                        pData(i).Points(iCount + n) = pData(j).Points(n)
                    Next
                    ' Delete shape j since we don't need it anymore
                    For n = j To UBound(pData) - 1
                        pData(n) = pData(n + 1)
                    Next
                    ReDim Preserve pData(UBound(pData) - 1)
                    
                    ' Then start the loop again.
                    Debug.Print "COUNT IS NOW ", UBound(pData)
                    Exit For ' Start the loop again
                Else
                    ' Alright we're done with this one
                    pData(i).Optimized = True
                End If
            End If
        Next
    Loop Until Not didMerge ' Continue looping until there's no more merging
    
    ' Finally, look for polygons that have a start and end point at the same co-ordinate and mark them as fillable.
    For i = 1 To UBound(pData)
        With pData(i)
            If .Points(1).x = .Points(UBound(.Points)).x And _
               .Points(1).y = .Points(UBound(.Points)).y Then
                
                    ' End of shape matches start
                    ' Therefore it is fillable.
                    .Fillable = True
               
            End If
        End With
    Next
    
End Function

Function optimizePolys()

    
    Dim i As Long
    Dim j As Long
    
    Dim dist As Double
    Dim bestDist As Double
    Dim bestLine As Long
    Dim bestIsEnd As Boolean ' Is the best match actually the END of another line?
    
    
    ' Run through the list of polygons. Order them so that when we reach the end of one,
    ' we immediately find the nearest next line.
    
    ' Step 1: Mark all of the polygons as "unordered"
    
    
    For i = 1 To UBound(pData)
        pData(i).Optimized = False
    Next
    
    
    ' Step 2: Loop through the unoptimized polygons
    For i = 1 To UBound(pData) - 1
        If pData(i).Optimized = False Then
            
            frmInterface.Caption = "Optimizing " & i & " / " & UBound(pData)
            If i Mod 50 = 0 Then DoEvents
            
            ' Find the next polygon that ends nearest this one.
            bestDist = 10000000
            bestLine = 0
            
            
            For j = 1 To UBound(pData)
                If j <> i And pData(j).Optimized = False And pData(j).LayerID = pData(i).LayerID Then
                    ' Calculate the distance
                    dist = pointDistance(pData(i).Points(UBound(pData(i).Points)), pData(j).Points(1))
                    If dist < bestDist Then
                         bestDist = dist
                         bestLine = j
                         bestIsEnd = False
                    End If
                    
                    ' Try the End of the line, since the line can be flipped if this makes more sense
                    dist = pointDistance(pData(i).Points(UBound(pData(i).Points)), pData(j).Points(UBound(pData(j).Points)))
                    If dist < bestDist Then
                         bestDist = dist
                         bestLine = j
                         bestIsEnd = True
                    End If
                    
                End If
            Next
            
            ' Now we know which line is best to go NEXT.
            ' So, move this line so that it is the next line after this one.
            If bestLine > 0 Then
                
                If bestIsEnd Then
                    ' We've got to flip the line around, since it's END point is closest to our end.
                    flipPolyStartEnd bestLine
                End If
                
                ' For now, we just swap the desired line with the next one.
                SwapLine pData(i + 1), pData(bestLine)
                
                
            End If
            
            'Mark ourselves as optimized
            pData(i).Optimized = True
        
        End If
    Next
        
End Function

Public Sub SwapLine(ByRef A As typLine, ByRef b As typLine)
    Dim c As typLine
    c = A
    A = b
    b = c

End Sub

Function exportGCODE(outFile As String, feedRate As Double, PlungeZ As Boolean, PPIMode As Boolean, PPIVal As Long, LoopMode As Boolean, Loops As Long, RaiseDist As Double)


    ' Export GCODE!
    Dim i As Long
    Dim j As Long
    Dim f As Long
    Dim scalar As Long
    Dim tLayer As String
    Dim t As String
    
        
    Dim minFeedRate As Long
    Dim maxFeedRate  As Long
    maxFeedRate = 200
    minFeedRate = 15
    
    
    f = FreeFile
    ' Draw the lines.
    
    If Dir(outFile) <> "" Then Kill outFile
    Open outFile For Append As f
        
        
        ' Get the extents
        Dim maxX As Double
        Dim maxY As Double
        
        Dim greyLevel As Double
        
        Dim isDefocused As Boolean
        Dim wasDefocused As Boolean
        
        Dim cutCount As Long
        Dim cuts As Long ' Defocusde cuts cut the same thing many times
        
        
        
        
    
        maxX = EXPORT_EXTENTS_X
        maxY = EXPORT_EXTENTS_Y
        
    
        ' Make it 5 inches high
        scalar = 1
        'scalar = 0.01
        
        
        ' Go to the corners
        Print #f, "G20 (Units are in Inches)"
        Print #f, "F" & Format(feedRate, "0.00000")
        Print #f, "G61 (Go to exact corners)" ' Added Sep 21, 2016
        
        If PPIMode Then
            Print #f, "S" & PPIVal & " (PPI mode with this many pulses per inch)"
        End If
        
        If LoopMode Then
        
            Print #f, "#201 = " & Loops & " (number of passes)"
            Print #f, "#200 = " & Format(RaiseDist * 0.0393701, "0.000000") & " (move the bed up incrementally by this much in inches)"
            Print #f, "#300 = 0 (bed movement distance storage variable)"
            Print #f, "#100 = 1 (layer number storage variable)"
            
            Print #f, "G1 W0.00000 (make sure bed is 0.0000 before you cut first pass)"
            Print #f, "o101 WHILE [#100 LE #201] (the number of passes is that the number after LE, LE = less or equal to)"
        
        End If
        
        
        ' Turn on the spindle
        'Print #f, "M3 S1"
        
        'Print #F, "G1 X0 Y0"
        'Print #F, "G1 X" & Round(maxX * scalar, 5) & " Y0"
        'Print #F, "G1 X" & Round(maxX * scalar, 5) & " Y" & Round(maxY * scalar, 5)
        'Print #F, "G1 X0 Y" & Round(maxY * scalar, 5)
    
        tLayer = "---"
    
        For i = 1 To UBound(pData)
            With pData(i)
                If UBound(.Points) > 0 Then
                    ' Set the feed rate.
                    'greyLevel = .greyLevel / GREYLEVELS
                    'Print #f, "F" & CLng((maxFeedRate - minFeedRate) * greyLevel) + minFeedRate
                
                    If .LayerID <> "Cut Boxes" Then
                    
                        If tLayer <> .LayerID Then
                            
                            wasDefocused = isDefocused
                            isDefocused = False
                            If layerInfo.Exists(.LayerID) Then
                                
                                If layerInfo.Item(.LayerID).Exists("pausebefore") Then
                                    Print #f, "(MSG,Change Laser Power!)"
                                    Print #f, "M0"
                                End If
                                
                                ' Are we defocused on this layer?
                                If layerInfo(.LayerID).Exists("defocused") Then
                                    isDefocused = True
                                    
                                    ' Bring it down
                                    Print #1, "F100 (Increated feed rate for defocused cuts)"
                                    Print #1, "G0 W-" & layerInfo(.LayerID)("defocused")
                                    
                                End If
                            
                            End If
                            
                            If wasDefocused And Not isDefocused Then
                                ' Bring the W back up
                                Print #1, "G0 W0"
                                ' Reset the feed rate
                                Print #f, "F" & Format(feedRate, "0.00000")
                            End If
                            
                            tLayer = .LayerID
                        End If
                        
                        Dim lastCutting As Boolean
                        
                        lastCutting = False
                        cutCount = 1
                        If isDefocused Then cutCount = 20
                        
                        
                        For cuts = 1 To cutCount
                            
                            For j = 1 To UBound(.Points)
                                With .Points(j)
                                    
                                    If j = 1 Then ' First point, just GO there.
                                        Print #f, "G0 X" & Format(.x * scalar, "0.00000") & " Y" & Format((maxY - .y) * scalar, "0.00000")
                                        'Print #f, "G1 z-0.0010"
                                        
                                        ' Turn on the spindle
                                        If PPIMode Then
                                            Print #f, "M3"
                                        Else
                                            Print #f, "M3 S1"
                                        End If
                                        'Print #f, "G0 Z -0.0100"
                                    Else
                                        t = "G1 X" & Format(.x * scalar, "0.00000") & " Y" & Format((maxY - .y) * scalar, "0.00000")
                                        
                                        ' Are we CUTTING to this point, or not?
                                        If lastCutting And pData(i).Points(j - 1).noCut = 1 Then
                                            
                                            If PlungeZ Then
                                                Print #f, "G0 Z 0.2"
                                            Else
                                                t = t & " M63 P0" ' STOP cutting
                                            End If
                                            
                                            
                                            lastCutting = False
                                        ElseIf Not lastCutting And pData(i).Points(j - 1).noCut = 0 Then
                                            
                                            If PlungeZ Then
                                                Print #f, "G0 Z -0.5"
                                            Else
                                                t = t & " M62 P0" ' START cutting
                                            End If
                                            
                                            lastCutting = True
                                        End If
                                        Print #f, t
                                    End If
                                End With
                            Next
                            
                            If isDefocused Then
                                ' Run the same line backwards again
                                For j = UBound(.Points) To 1 Step -1
                                    With .Points(j)
                                        If j = UBound(pData(i).Points) Then ' First point, just GO there.
                                            Print #f, "G0 X" & Format(.x * scalar, "0.00000") & " Y" & Format((maxY - .y) * scalar, "0.00000")
                                        Else
                                            t = "G1 X" & Format(.x * scalar, "0.00000") & " Y" & Format((maxY - .y) * scalar, "0.00000")
                                            If lastCutting And pData(i).Points(j - 1).noCut = 1 Then
                                                t = t & " M63 P0" ' STOP cutting
                                                lastCutting = False
                                            ElseIf Not lastCutting And pData(i).Points(j - 1).noCut = 0 Then
                                                t = t & " M62 P0" ' START cutting
                                                lastCutting = True
                                            End If
                                            Print #f, t
                                        End If
                                    End With
                                Next
                            End If
                        Next
                        
                        'Print #F, "G0 Z0.0010"
                        ' Turn off the spindle
                        Print #f, "M5"
                        If PlungeZ Then Print #f, "G0 Z 0.2"
                        
                        'Print #f, "G0 Z 0.0100"
                        
                        'Print #f, "G1 Z0.0010"
                        Print #f, ""
                    End If
                    
                End If
                
            End With
        Next
            
        Print #f, "M5"
        If PlungeZ Then Print #f, "G0 Z 0.2"
        
        If LoopMode Then
            Print #f, "#300 = [#200*#100]"
            Print #f, "G1 W#300 (move the bed up according to the layer its on)"
            Print #f, "#100 = [#100+1] (add one to the layer counter)"
            Print #f, "o101 ENDWHILE"
        End If
        
        Print #f, "M30"
    Close #f

End Function

Public Function MoveLayerToEnd(LayerID As String)
    ' Make a new list of just the lines not in this layer, then put these at the end
    
    Dim pNew() As typLine
    ReDim pNew(0)
    Dim i As Long
    Dim j As Long
    Dim n As Long
    For i = 1 To UBound(pData)
        If pData(i).LayerID = LayerID Then
            ' Put this aside
            n = UBound(pNew) + 1
            ReDim Preserve pNew(n)
            pNew(n) = pData(i)
        Else
            j = j + 1
            pData(j) = pData(i)
        End If
    Next
    
    ' Now add to end
    For i = 1 To n
        j = j + 1
        pData(j) = pNew(i)
    Next
    
    ' All done
    
    
End Function
