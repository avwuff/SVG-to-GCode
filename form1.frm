VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbalTbar6.ocx"
Begin VB.Form frmInterface 
   AutoRedraw      =   -1  'True
   Caption         =   "SVG to GCODE by Avatar-X"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   19305
   LinkTopic       =   "Form1"
   ScaleHeight     =   734
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1287
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cFeedRate 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      TabIndex        =   5
      Top             =   300
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2715
      Left            =   660
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   4
      Top             =   60000
      Width           =   3855
   End
   Begin VB.PictureBox picRulers 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4200
      Index           =   1
      Left            =   0
      ScaleHeight     =   280
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   3
      Top             =   840
      Width           =   300
   End
   Begin VB.PictureBox picRulers 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   300
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   2
      Top             =   540
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog COMDLG 
      Left            =   14580
      Top             =   10260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   6300
      Left            =   11460
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   840
      Width           =   2835
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   7815
      Left            =   300
      ScaleHeight     =   517
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   737
      TabIndex        =   0
      Top             =   840
      Width           =   11115
   End
   Begin vbalTBar6.cReBar cReBar1 
      Left            =   0
      Top             =   0
      _ExtentX        =   5212
      _ExtentY        =   873
   End
   Begin vbalTBar6.cToolbar TB1 
      Height          =   435
      Left            =   3000
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   767
   End
   Begin vbalIml6.vbalImageList vbalImageList1 
      Left            =   960
      Top             =   9420
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   24
      IconSizeY       =   24
      ColourDepth     =   32
      Size            =   31980
      Images          =   "form1.frx":0000
      Version         =   131072
      KeyCount        =   13
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Zoom As Double
Public panX As Double
Public panY As Double
Dim mX As Single
Dim mY As Single
Dim oX As Double
Dim oY As Double
Dim mouseDown As Boolean



Private Sub cmdScale_Click()
    Dim w As Double
    Dim maxX As Double, maxY As Double
    Dim scalar As Double
    Dim i As Long
    Dim j As Long
    
    w = Val(InputBox("Scale width to ? (inches)", "Scale", "5"))
    
    getExtents maxX, maxY
    
    If w > 0 Then
    
        scalar = w / maxY
    
        For i = 1 To UBound(pData)
            With pData(i)
                For j = 1 To UBound(.Points)
                    With .Points(j)
                        .x = scalar * .x
                        .y = scalar * .y
                    End With
                Next
            End With
        Next
    End If
    
    drawLines
    
    
End Sub

Private Sub Command1_Click()
    parseSVG App.path & "\wuffy-fill.svg"
    'parseSVG App.Path & "\drawing-3.svg"
    
    Debug.Print "Drawing ", UBound(pData)
    
    drawLines
    
    updateList
    
    
End Sub

Private Sub cmd1_Click()

    parsePath "M 402.85714,489.50504 L -94.285714,92.362183", ""
    
    drawLines
    
End Sub


Private Sub drawLines()
    
    Picture1.Cls
    
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim cOrig As Long
    
    ' Draw the lines.
    
    'Debug.Print panX, panY
    
    Dim polyPoints() As POINTAPI
    Dim isDefocused As Boolean
    Dim drawNonCut As Boolean
    
    Dim lastX As Double, lastY As Double
    lastX = -10000
    lastY = -10000
    
    
    drawNonCut = TB1.ButtonChecked("noncut")
    
    
    For i = 1 To UBound(pData)
        With pData(i)
            Picture1.ForeColor = vbBlack
            Picture1.DrawWidth = 1
            Picture1.DrawStyle = 0
            c = .greyLevel * (255 / GREYLEVELS)
            c = RGB(c, c, c)
            cOrig = c
            isDefocused = False
            If layerInfo.Exists(.LayerID) Then
                isDefocused = layerInfo(.LayerID).Exists("defocused")
            End If
            
            Picture1.DrawWidth = IIf(isDefocused, 5, 1)
            
            ' Draw a line from the last point
            If lastX <> -10000 And lastY <> -10000 And drawNonCut And .LayerID <> "Cut Boxes" Then
                If UBound(.Points) > 0 Then
                
                    ' Dashed line to here
                    Picture1.DrawStyle = DrawStyleConstants.vbDashDot
                    Picture1.Line ((.Points(1).x + panX) * Zoom, _
                                    (.Points(1).y + panY) * Zoom)-( _
                                    lastX, _
                                    lastY), RGB(200, 200, 200)
                    Picture1.DrawStyle = DrawStyleConstants.vbSolid
                End If
            End If
            
            For j = 1 To UBound(.Points) - 1
            
                c = cOrig
                If .Points(j).noCut = 1 Then c = RGB(150, 0, 0)
                If isDefocused Then c = RGB(0, 200, 0)
                If .LayerID = "Cut Boxes" Then
                    c = RGB(255, 0, 255)
                    Picture1.DrawStyle = DrawStyleConstants.vbDot
                End If
                
                
            
                Picture1.Line ((.Points(j).x + panX) * Zoom, _
                                (.Points(j).y + panY) * Zoom)-( _
                                (.Points(j + 1).x + panX) * Zoom, _
                                (.Points(j + 1).y + panY) * Zoom), c
                             
                             
            Next
            
            If UBound(.Points) > 0 And .LayerID <> "Cut Boxes" Then
                lastX = (.Points(UBound(.Points)).x + panX) * Zoom
                lastY = (.Points(UBound(.Points)).y + panY) * Zoom
            End If
            
            
            
            
            
            
            If .Fillable And .ContainedBy = 0 And False Then
                Picture1.FillStyle = 0
                Picture1.ForeColor = vbBlue
                Picture1.DrawWidth = 1
                Picture1.DrawStyle = 5
                
                
                ReDim polyPoints(UBound(.Points) - 1)
                For j = 1 To UBound(.Points)
                    polyPoints(j - 1).x = (.Points(j).x + panX) * Zoom
                    polyPoints(j - 1).y = (.Points(j).y + panY) * Zoom
                Next
                
                ' Add any that are fillable.
                addFillPolies polyPoints, i
                
                
                
                Polygon Picture1.hDC, polyPoints(0), UBound(polyPoints) 'call the polygon function
            End If
            
        End With
    Next
    
    Dim A As Long
    
    
    For i = 1 To List1.ListCount
        If List1.Selected(i - 1) Then
            A = List1.ItemData(i - 1)
            If A > 0 And A <= UBound(pData) Then
        
                With pData(A)
                    
                
                    Picture1.Circle ((.Points(1).x + panX) * Zoom, (.Points(1).y + panY) * Zoom), 5, vbGreen
                
                    For j = 1 To UBound(.Points) - 1
                    
                        Picture1.ForeColor = vbRed
                        Picture1.DrawWidth = 3
                        
                        Picture1.Line ((.Points(j).x + panX) * Zoom, _
                                        (.Points(j).y + panY) * Zoom)-( _
                                        (.Points(j + 1).x + panX) * Zoom, _
                                        (.Points(j + 1).y + panY) * Zoom)
                    
                        'If j > 1 Then Picture1.Circle ((.Points(j).x + panX) * Zoom, (.Points(j).y + panY) * Zoom), 5, vbBlue
                    
                    Next
                
                    Picture1.DrawWidth = 1
                    
                    Picture1.Circle ((.Points(UBound(.Points)).x + panX) * Zoom, (.Points(UBound(.Points)).y + panY) * Zoom), 5, vbRed
                
                End With
            End If
        End If
    Next
    
            
    updateRulers
    

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

'
'    ' Calculate the X and Y center of each polygon
'    Dim i As Long
'    Dim sorted As Boolean
'    For i = 1 To UBound(pData)
'        calcPolyCenter i, pData(i).xCenter, pData(i).yCenter
'    Next
'
'    ' Sort!
'    Do
'        sorted = False
'        For i = 1 To UBound(pData) - 1
'
'            If pData(i).yCenter > pData(i + 1).yCenter Then
'                sorted = True
'                Swap pData(i), pData(i + 1)
'            ElseIf pData(i).yCenter = pData(i + 1).yCenter Then
'                If pData(i).xCenter < pData(i + 1).xCenter Then
'                    sorted = True
'                    Swap pData(i), pData(i + 1)
'                End If
'            End If
'
'        Next
'    Loop Until Not sorted
'
'    drawLines
'    updateList
    

End Sub




Private Sub Command7_Click()

    'rasterLinePoly Val(Text1)
    'drawLines
    
    Dim x As Double
    Dim y As Double
    Dim Ang As Double
    Dim n As Double
    Dim nSeg As Double
    
    ReDim pData(0)
    
    
    'newLine
    'For x = 0 To 1400 Step 1
    '    y = (Sin(x / 40) * 100) + 100
    '    addPoint x, y
    'Next
    
    'newLine
    'For x = 0 To 1400 Step 1
    '    y = (Sin(x / 10) * 100) + 310
    '    addPoint x, y
    'Next

    'circle
    newLine
    For Ang = 0 To 3.14159 * 2 Step 0.05
        x = (Cos(Ang) * 200) + 250
        y = (Sin(Ang) * 200) + 220
        addPoint x, y
    Next
    
    newLine
    For Ang = 0 To 3.14159 * 2 Step 0.1
        x = (Cos(Ang) * 200) + 750
        y = (Sin(Ang) * 200) + 220
        addPoint x, y
    Next
    

    newLine
    For x = 0 To 1400 Step 4
        y = (Sin(x / 160) * 100) + 530
        addPoint x, y
    Next
    
    'circle
    newLine
    For Ang = 0 To 3.14159 * 2 Step 0.01
        x = (Cos(Ang) * 200) + 250
        y = (Sin(Ang) * 200) + 900
        addPoint x, y
    Next

    
    nSeg = 20
    For n = 1 To nSeg Step 2
        newLine
        
        addPoint 700, 900
        
        For Ang = ((3.14159 * 2) / nSeg) * n To ((3.14159 * 2) / nSeg) * (n + 1) Step 0.01
            x = (Cos(Ang) * 200) + 700
            y = (Sin(Ang) * 200) + 900
            addPoint x, y
        Next
        
        addPoint 700, 900
        
        
    Next
    
    
    
    nSeg = 70
    For n = 1 To nSeg Step 2
        newLine
        
        addPoint 1200, 900
        
        For Ang = ((3.14159 * 2) / nSeg) * n To ((3.14159 * 2) / nSeg) * (n + 1) Step 0.01
            x = (Cos(Ang) * 200) + 1200
            y = (Sin(Ang) * 200) + 900
            addPoint x, y
        Next
        
        addPoint 1200, 900
        
        
    Next
    
    
    drawLines
    updateList
    
    

End Sub

Private Sub Command8_Click()


End Sub

Private Sub Command9_Click()

    
End Sub

Private Sub Form_Load()
    Zoom = 1
    ReDim pData(0)
    
    
    Me.Caption = "Av's SVG to GCODE v " & App.Major & "." & App.Minor & "." & App.Revision
    
    
     With TB1
        .ImageSource = CTBExternalImageList
        .SetImageList vbalImageList1.hIml
        .CreateToolbar 24, True, True, True
        
        .AddButton "Open", 0, , , "Open", CTBAutoSize, "open"
        .AddButton "Scale", 1, , , "Scale", CTBAutoSize, "scale"
        .AddButton "Export", 2, , , "Export", CTBAutoSize, "export"
        
        .AddButton "Specify the Feed Rate used to cut your design", 10, , , "Feed Rate:", CTBAutoSize, "rate"
        .AddControl cFeedRate.Hwnd
        
        .AddButton "Zoom In", 3, , , "Zoom In", CTBAutoSize, "zoomin"
        .AddButton "Zoom Out", 4, , , "Zoom Out", CTBAutoSize, "zoomout"
        
        .AddButton "Display the Non-Cut Paths", 6, , , "Show NonCuts", CTBAutoSize Or CTBCheck, "noncut"
        .AddButton "Plunge the Z (for engraver)", 9, , , "Z Plunge", CTBAutoSize Or CTBCheck, "zplunge"
        .ButtonChecked("noncut") = True
        
        .AddButton "Raster-fill the shapes by line.", 11, , , "Raster Fill", CTBAutoSize, "fill"
        
        
        
        
        
        .AddButton "Generate a puzzle", 5, , , "Puzzle", CTBAutoSize, "puzz"
        .AddButton "Split into multiple GCode files by page", 7, , , "Split by Pages", CTBAutoSize, "pages"
        .AddButton "Duplicate the design in multiple rows and columns", 8, , , "Tile", CTBAutoSize, "tile"
        
        .AddButton "Rotate 90", 12, , , "Rotate 90", CTBAutoSize, "rotate90"
        
        
        '.AddButton "box", 5, , , "box", CTBAutoSize, "box"
        
        
        
'        If browseMode = eB_Structure And Not chooserMode Then
'            .AddButton "Top Node", 10, , , "[Application]", CTBAutoSize Or CTBDropDown, "AppDrop"
'            .AddButton "Type Browser", 4, , , "Types", CTBAutoSize, "Types"
'            .AddButton "Relationship Browser", 9, , , "Relationships", CTBAutoSize, "Relationships"
'            .AddButton "New Object", 11, , , "New Object", CTBAutoSize, "NewObject"
'
'            .AddButton "Display tabs as web-based forms", 12, , , "Web Tabs", CTBAutoSize Or CTBCheck, "WebForms"
'
'
'            .AddButton , , , , , CTBSeparator Or CTBAutoSize, "Sep"
'            .AddButton "System", 3, , , "System", CTBAutoSize Or CTBDropDown, "System"
'            .AddButton "About", 13, , , "About", CTBAutoSize, "About"
'        End If
'
'        .AddButton "Sort", 6, , , "Sort", CTBAutoSize Or CTBDropDown, "Sort"
'
'        .ButtonEnabled("Back") = False
'        .ButtonEnabled("Forward") = False
        
    End With
    
    
    
    With cReBar1
        
        ' a) Create the rebar:
        .ImageSource = CRBLoadFromFile
        .CreateRebar Me.Hwnd
        
        ' b) Add the toolbar & combo boxes.
        ' When you add a band, the rebar automatically sets the IdealWidth
        ' to the size of the object you've added, and makes the Minimum
        ' size the same.  By allowing a smaller minimum size, the rebar
        ' will show a chevron when the band is reduced.
        
        ' i) Add the 24x24 toolbar with text:
        .AddBandByHwnd TB1.Hwnd, , , , "Toolbar1"
        .BandChildMinWidth(.BandCount - 1) = 24
        
    End With
    
    With cFeedRate
        .AddItem "20"
        .AddItem "40"
        .AddItem "60"
        .AddItem "80"
        .AddItem "100"
    End With
    
    cFeedRate.Text = 20
    
    
    
    
End Sub

Private Sub Form_Resize()
    
    List1.left = Me.ScaleWidth - List1.Width
    Picture1.Width = List1.left - Picture1.left
    Picture1.Height = Me.ScaleHeight - Picture1.tOp
    List1.Height = Picture1.Height
    
    cReBar1.RebarSize
    
    
    
    picRulers(0).Width = Picture1.Width
    picRulers(1).Height = Picture1.Height
    
    drawLines
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub List1_Click()
    drawLines
    Picture1.Refresh
    
End Sub

Private Sub List1_DblClick()
    Dim A As Long
    A = getListLine
    If A > 0 Then
        pData(A).Fillable = Not pData(A).Fillable
        updateList
    End If
    
    
End Sub

Function getListLine() As Long

    Dim A As Long
    A = List1.ListIndex
    If A > -1 Then
        getListLine = List1.ItemData(A)
    End If

End Function

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim A As Long
    Dim i As Long
    Dim j As Long
    
    Dim lI As Long
    Dim doDel As Boolean
    
    
    For lI = 1 To List1.ListCount
        If List1.Selected(lI - 1) Then
            A = List1.ItemData(lI - 1)
           
            If KeyCode = vbKeyDelete Then
                pData(A).isDel = True
                doDel = True
            End If
        End If
    Next
               
    If doDel Then
        
        j = 0
        For i = 1 To UBound(pData)
            If pData(i).isDel Then
                ' Skip this one
            Else
                j = j + 1
                pData(j) = pData(i)
            End If
        Next
        ReDim Preserve pData(j)
        
        drawLines
        updateList
        On Error Resume Next
        List1.ListIndex = A - 1
    End If

End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim A As Long
    Dim R As Long
    Dim i As Long
    
    Dim b As Scripting.Dictionary
    Dim selLines As New Collection
    Dim newLayer As String
    
    
    Dim mc As New mcPopupMenu
    If Button = 2 Then
        
        For i = 1 To List1.ListCount
            If List1.Selected(i - 1) Then
                A = List1.ItemData(i - 1)
        
                If A > 0 Then
                    selLines.Add A
                End If
            End If
        Next
                
        If selLines.count = 1 Then
            A = selLines(1)
            With pData(A)
                mc.Add 0, "Layer: " & .LayerID, , , mceGrayed
                mc.Add 1, "Fillable", , .Fillable
                
            
                'mc.Add 1, "",,pdata(i).
                
                If Not layerInfo.Exists(.LayerID) Then
                    layerInfo.Add .LayerID, New Scripting.Dictionary
                End If
                
                Set b = layerInfo.Item(.LayerID)
                
                mc.Add 0, "-"
                mc.Add 10, "Pause before layer", , b.Exists("pausebefore")
                mc.Add 11, "Defocused Cut Layer", , b.Exists("defocused")
                mc.Add 12, "Move Layer to End"
                mc.Add 0, "-"
                mc.Add 20, "Remove Last Segment"
                mc.Add 30, "Set Layer"
                mc.Add 40, "DEBUG: Path Data"
                
                R = mc.Show

            End With
            
        
        ElseIf selLines.count > 1 Then ' Multiple lines selected
            
            With pData(selLines(1))
                mc.Add 100, selLines.count & " objects selected"
                mc.Add 0, "-"
                mc.Add 1, "Fillable", , .Fillable
                
                mc.Add 30, "Set Layer"
            End With
            
            R = mc.Show
            
            
        End If
        
        Select Case R
            Case 30
                newLayer = InputBox("Set layer to?", "Set Layer")
                
        End Select
        
        For i = 1 To selLines.count
            A = selLines(i)
            With pData(A)
                Select Case R
                    Case 0
                        Exit Sub
                    Case 1
                        .Fillable = Not .Fillable
                    Case 10
                        If b.Exists("pausebefore") Then
                            b.Remove "pausebefore"
                        Else
                            b.Add "pausebefore", True
                        End If
                    Case 11
                        If b.Exists("defocused") Then
                            ' turn it off
                            b.Remove "defocused"
                        Else
                            R = Val(InputBox("How many inches to move down?", "Defocus Cuts", 3))
                            If R > 0 Then
                                b.Add "defocused", R
                            End If
                        End If
                    Case 12 ' Move Layer to End
                        newLayer = .LayerID
                    Case 20
                        ' Remove last secmet
                        ReDim Preserve .Points(UBound(.Points) - 1)
                        drawLines
                        
                    Case 30 ' Set layer
                        .LayerID = newLayer
                    
                    Case 40
                        Clipboard.Clear
                        Clipboard.SetText .PathCode
                End Select
            End With
            
            Select Case R ' Outside WITH block
                Case 12 ' Move Layer to End
                    ' Move all lines that are on this layer to the end.
                    MoveLayerToEnd newLayer
            End Select
        Next
        
        Select Case R
            Case 30
                optimizePolys
        End Select
        
        updateList
    End If
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    

    mX = x
    mY = y
    oX = panX
    oY = panY
        
    mouseDown = True
    
    
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'    Dim p1 As pointD
'    Dim p2 As pointD
'    Dim a As Double
'
'    p1.x = 200
'    p1.y = 200
'    p2.x = CDbl(x)
'    p2.y = CDbl(y)
'
'    a = angleFromPoint(p1, p2)
'    Debug.Print a * (180 / PI)
'
'    Picture1.Cls
'    Picture1.Line (200, 200)-(x, y)
'
    

'        Dim result() As pointD
'        Dim i As Long
'
'    If Button = 1 Then
'        drawLines
'        Picture1.ForeColor = vbBlue
'
'        Picture1.Line (X, Y)-(mX, mY)
'
'        result = lineIntersectPoly(newPoint(CDbl(X), CDbl(Y)), newPoint(CDbl(mX), CDbl(mY)), 3)
'
'        Debug.Print UBound(result)
'        For i = 1 To UBound(result)
'            Picture1.Circle (result(i).X, result(i).Y), 5
'        Next
'
'        Picture1.Refresh
'
'    End If
'
'Exit Sub

    If mouseDown Then
        panX = oX + ((x - mX) / Zoom)
        panY = oY + ((y - mY) / Zoom)
        drawLines
        Picture1.Refresh
        
        
    End If
End Sub


Private Sub fillPoly(lineID As Long, useBlack As Boolean)

    ' Get the bounds of this shape.
    Dim maxX As Double, maxY As Double
    Dim minX As Double, minY As Double
    Dim x As Double, y As Double
    
    getPolyBounds lineID, minX, minY, maxX, maxY
    
    For x = minX To maxX
        For y = minY To maxY
        
            If pointIsInPoly(lineID, x, y) Then
                Picture1.PSet (x, y), IIf(useBlack, vbRed, vbWhite)
                
            End If
        Next
        DoEvents
        
    Next
    

End Sub


Private Sub doFills()

    ' Iterate through the polygons that can be filled
    ' First, figure out if any polygons are inside any other polygons.
        
    Dim i As Long
    Dim j As Long
    Dim matchArea As Double
    Dim bestMatchArea As Double
    Dim bestMatchID As Long
    
    Set containList = New Scripting.Dictionary
    
    
    For i = 1 To UBound(pData)
        With pData(i)
            If .Fillable Then
                
                
                
                bestMatchArea = 0
                bestMatchID = 0
                
                ' Find which possible polygons might contain me
                For j = 1 To UBound(pData)
                    If i <> j And pData(j).Fillable And pData(j).LayerID = .LayerID Then
                        
                        
                        If canPolyFitInside(i, j) Then
                            
                            ' Figure out how big this polygon is
                            matchArea = getPolyArea(j)
                            If matchArea < bestMatchArea Or bestMatchID = 0 Then
                                bestMatchArea = matchArea
                                bestMatchID = j
                            End If
                        End If
                    End If
                    
                Next
                
                If bestMatchID > 0 Then
                    ' Found a match.
                    .ContainedBy = bestMatchID
                    
                    If Not containList.Exists(.ContainedBy) Then
                        containList.Add .ContainedBy, New Collection
                    End If
                    
                    containList(.ContainedBy).Add i
                End If
            End If
        End With
        
        Me.Caption = "Checking inside " & i & " / " & UBound(pData)
        If i Mod 20 = 0 Then DoEvents
        
    Next
    
    
    ' Now that we have the shapes and who contains them, reorder the list so that the lowest level of contained ones are cut first.
    SetLevelNumber 0, 0
    
    
    
    ' Now sort by level number going down
    Me.Caption = "Sorting by level number..."
    Dim sorted As Boolean
    Do
        sorted = False
        For i = 1 To UBound(pData) - 1
            If pData(i).LevelNumber < pData(i + 1).LevelNumber And pData(i).LayerID = pData(i + 1).LayerID Then ' Swap!
                SwapLine pData(i + 1), pData(i)
                sorted = True
            End If
        Next
    Loop Until Not sorted
    'fillAll 0, True
    
    
    'updateList
    
    
End Sub

Private Sub SetLevelNumber(Container As Long, LevelNum As Long)
    
    Dim i As Long
    For i = 1 To UBound(pData)
        With pData(i)
            If .ContainedBy = Container Then
                .LevelNumber = LevelNum
                SetLevelNumber i, LevelNum + 1
            End If
        End With
    Next



End Sub

Function fillAll(containerID As Long, fillWith As Boolean)

    ' Fill all poly's with containerID as specified
    Dim i As Long
    For i = 1 To UBound(pData)
        With pData(i)
            If .Fillable And .ContainedBy = containerID Then
            
                fillPoly i, fillWith
                
                fillAll i, Not fillWith
            
            End If
        End With
    Next
End Function

Function updateList()
    Dim i As Long
    Dim tLayer As String
        
    Dim b As Scripting.Dictionary
    
    tLayer = "---"
    List1.Clear
    For i = 1 To UBound(pData)
        If pData(i).LayerID <> tLayer Then
        
            If Not layerInfo.Exists(pData(i).LayerID) Then
                layerInfo.Add pData(i).LayerID, New Scripting.Dictionary
            End If
            
            Set b = layerInfo.Item(pData(i).LayerID)
            
            List1.AddItem "[Layer " & pData(i).LayerID & " " & IIf(b.Exists("pausebefore"), "PAUSE", "") & "]"
            tLayer = pData(i).LayerID
        End If
        List1.AddItem "   Line " & i & ": " & IIf(pData(i).Fillable, "F", "") & " (" & UBound(pData(i).Points) & " segs) in " & pData(i).ContainedBy
        List1.ItemData(List1.NewIndex) = i
    Next

End Function

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseDown = False
    
End Sub

Private Sub Timer1_Timer()

    Dim x As Long
    Dim y As Long
    Dim lx As Long, ly As Long
    Dim phase As Long
    
    phase = Timer * 200
    
    Picture1.Cls

    For x = 1 To 200
        y = (Sin((x + phase) / 20) * 50) + 50
        Picture1.Line (lx, ly)-(x, y)
        lx = x
        ly = y
    Next
    
    lx = 0
    ly = 0
    

    For x = 1 To 200
        y = (Sin((x + phase) / 30) * 50) + 50
        Picture1.Line (lx, ly)-(x, y)
        lx = x
        ly = y
    Next

End Sub



Private Sub cmdOpenFile()
    
    
    
    Dim fName As String
    With COMDLG
        .Filter = "Supported Files (*.svg, Images)|*.svg;*.bmp;*.jpg;*.gif"
        .CancelError = False
        .DialogTitle = "Open File"
        .ShowOpen
        fName = .FileName
    End With

    If fName <> "" Then
        
        
        
        ' Reset pan and zoom settings
        
        Select Case getFileExten(fName)
            Case "svg"
                
                layerInfo.RemoveAll
                
                parseSVG fName
                
                Debug.Print "Drawing ", UBound(pData)
                
                ' Optimize the shapes for best drawing
                'sortByLayers
                
                mergeConnectedLines
                
                optimizePolys
            
            Case "jpg", "bmp", "gif"
                ' Open and rasterize this file
                rasterFile fName
                
            
            Case Else
                
        End Select
        
        panX = 0
        panY = 0
        Zoom = 1
        
        doFills
        
        zoomToFit
        
        getExtents EXPORT_EXTENTS_X, EXPORT_EXTENTS_Y
        
        
        updateList
        
        
    End If
    
    
    Me.Caption = "Av's SVG to GCODE v " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub cmdExport()
    Dim fName As String

    
    On Error GoTo done
    
    With COMDLG
        
        .FileName = getFolderNameFromPath(.FileName) & "\" & getFileNameNoExten(getFileNameFromPath(.FileName)) & ".ngc"
        
        .Filter = "GCODE Files (*.ngc)|*.ngc"
        
        .DialogTitle = "Export GCODE"
        .ShowSave
        fName = .FileName
    End With
    
    On Error GoTo 0
    
    
    If fName <> "" Then
        exportGCODE fName, Val(cFeedRate.Text), TB1.ButtonChecked("zplunge")
    End If

done:

End Sub

Sub updateRulers()
    
    ' Draw the rulers.
    
    Dim pixelStep As Double
    Dim insideStep As Long
    Dim rStart As Double, rEnd As Double, rStep As Double
    Dim i As Double, n As Long
    Dim n1 As Double
    Dim emph As Double
    Dim p As Long
    
    pixelStep = 1
    Do
        n1 = measureToRuler(pixelStep, True) - measureToRuler(0, True)
        If n1 < 30 Then pixelStep = pixelStep + 1
    Loop Until n1 >= 30
    
    ' Calculate the inside step?
    n1 = measureToRuler(pixelStep, True) - measureToRuler(0, True)
    
    'addtolog "The raw pixel difference between each interval is: ", n1
    
    insideStep = 20: emph = 10
    If n1 < 70 Then insideStep = 10: emph = 5
    If n1 < 35 Then insideStep = 5
    
    'Exit Sub
    
    
    ' HORIZONTAL RULER
    picRulers(0).Cls
    rStart = rulerToMeasure(0, True)
    rEnd = rulerToMeasure(picRulers(0).Width, True)
    ' Round down and up to find the nearest ruler points.
    rStart = Int(rStart / pixelStep) * pixelStep
    rEnd = (Int(rEnd / pixelStep) + 1) * pixelStep
    For i = rStart To rEnd Step (pixelStep / insideStep)
        n = measureToRuler(i, True)
        If (i * 10000) Mod (pixelStep * 10000) = 0 Then
            picRulers(0).Line (n, 0)-(n, 16)
            picRulers(0).PSet (n + 1, 2), picRulers(0).BackColor
            picRulers(0).Print Abs(Round(i))
        Else
            If Round(i / (pixelStep / insideStep)) Mod emph = 0 Then
                picRulers(0).Line (n, 11)-(n, 16)
            Else
                picRulers(0).Line (n, 13)-(n, 16)
            End If
            
        End If
    Next
    
    ' VERTICAL RULER
    picRulers(1).Cls
    rStart = rulerToMeasure(0, False)
    rEnd = rulerToMeasure(picRulers(1).Height, False)
    ' Round down and up to find the nearest ruler points.
    rStart = Int(rStart / pixelStep) * pixelStep
    rEnd = (Int(rEnd / pixelStep) + 1) * pixelStep
    For i = rStart To rEnd Step (pixelStep / insideStep)
        n = measureToRuler(i, False)
        If (i * 10000) Mod (pixelStep * 10000) = 0 Then
            picRulers(1).Line (0, n)-(16, n)
            picRulers(1).PSet (0, n + 1), picRulers(0).BackColor

            picRulers(1).Print addEnters(CStr(Abs(Round(i))))
        Else
            If Round(i / (pixelStep / insideStep)) Mod emph = 0 Then
                picRulers(1).Line (11, n)-(16, n)
            Else
                picRulers(1).Line (13, n)-(16, n)
            End If
        End If
    Next
    
    
    picRulers(0).Refresh
    picRulers(1).Refresh

End Sub

Function addEnters(inSt As String) As String
    ' Add a line feed after each letter
    Dim i As Long
    For i = 1 To Len(inSt)
        addEnters = addEnters & Mid(inSt, i, 1) & vbCrLf
    Next
End Function

Function measureToRuler(inMeas As Double, isX As Boolean) As Long
    ' Turn a measure value into an on-screen pixel value for the ruler.
    If isX Then
        measureToRuler = (inMeas + panX) * Zoom
    Else
        measureToRuler = (inMeas + panY) * Zoom
    End If

End Function

Function rulerToMeasure(inPx As Long, isX As Boolean) As Double

    ' Calculate the measure at the specified ruler value, based on the scroller position and such.
    If isX Then
        rulerToMeasure = (inPx / Zoom) - panX
    Else
        rulerToMeasure = (inPx / Zoom) - panY
    End If
    

End Function


Function cmdScale()
    
    ' Scale the work
    
    Dim w As Double, H As Double
    Dim maxX As Double, maxY As Double
    Dim scalarW As Double, scalarH As Double
    Dim i As Long
    Dim j As Long
    
    getExtents maxX, maxY
    
    If maxX = 0 Or maxY = 0 Then Exit Function
    
    Load frmScale
    frmScale.originalAspect = maxX / maxY
    frmScale.originalW = maxX
    frmScale.originalH = maxY
    frmScale.updatingValue = True
    frmScale.txtWidth = Round(maxX, 5)
    frmScale.txtHeight = Round(maxY, 5)
    frmScale.updatingValue = False
    frmScale.Show vbModal, Me
    
    w = frmScale.setW
    H = frmScale.setH
    Unload frmScale
    
    
    
    If w > 0 And H > 0 Then
    
        scalarW = w / maxX
        scalarH = H / maxY
    
        For i = 1 To UBound(pData)
            With pData(i)
                For j = 1 To UBound(.Points)
                    With .Points(j)
                        .x = scalarW * .x
                        .y = scalarH * .y
                    End With
                Next
            End With
        Next
    End If
    
    zoomToFit

End Function

Function zoomToFit()

    ' Fit the object on the screen
    Dim maxX As Double, maxY As Double
    getExtents maxX, maxY
    
    If maxX = 0 Or maxY = 0 Then Exit Function
    
    Zoom = Min(Picture1.Width / maxX, Picture1.Height / maxY) * 0.95
    
    ' Set the pans to center it
    panY = ((Picture1.Height / 2) - ((maxY * Zoom) / 2)) / Zoom
    panX = ((Picture1.Width / 2) - ((maxX * Zoom) / 2)) / Zoom
    
    drawLines
    
    
End Function

Private Sub TB1_ButtonClick(ByVal lButton As Long)


    
    
    Select Case TB1.ButtonKey(lButton)
        Case "open"
            cmdOpenFile
        Case "export"
            cmdExport
            
        
        Case "zoomin"
            
            panX = panX - (Picture1.Width / 4) / Zoom
            panY = panY - (Picture1.Height / 4) / Zoom
            
            Zoom = Zoom * 2
            drawLines
        
        Case "zoomout"
            
            
            Zoom = Zoom / 2
            
            panX = panX + (Picture1.Width / 4) / Zoom
            panY = panY + (Picture1.Height / 4) / Zoom
            
            drawLines
        
        Case "scale"
            cmdScale
            
        Case "fill"
            cmdFill
            
        Case "puzz"
            
            If MsgBox("Create a Puzzle design? This will erase your current design.", vbYesNo Or vbQuestion Or vbDefaultButton2) = vbYes Then
                doPuzzle
            End If
        
        Case "pages"
            
            doPages
            
        Case "box"
            
            'doBox
        
        Case "tile"
            frmTile.Show
        
        Case "noncut"
            drawLines
            
        Case "rotate90"
                
            Dim i As Long, j As Long
            
                For i = 1 To UBound(pData)
                    For j = 1 To UBound(pData(i).Points)
                        With pData(i).Points(j)
                            Swap pData(i).Points(j).x, pData(i).Points(j).y
                            ' Invert the Y
                            pData(i).Points(j).y = 17 - pData(i).Points(j).y
                        End With
                    Next
                Next
    
            drawLines
            
            
            
            
    End Select
End Sub

Sub cmdFill()

    ' Calculate who is in what
    doFills
    
    Dim D As Double
    
    Dim maxX As Double, maxY As Double
    getExtents maxX, maxY
        
    If maxX > 24 Or maxY > 18 Then
        MsgBox "Your document is too big to fit on the laser. Please scale first.", vbCritical
        
        Exit Sub
    End If
    
    D = Val(InputBox("How far apart should the raster lines be in inches?", "Raster DPI", "0.004"))
    If D < 0 Then Exit Sub
    
    
    Dim i As Long
    'For i = 1 To UBound(pData)
    '    If pData(i).ContainedBy = 0 And pData(i).Fillable Then
    '        Debug.Print "Rasterizing ", i
    '        rasterLinePoly i, D, "Fill"
    '    End If
    'Next
    
    rasterDocument D, "Fill"
    
     
    
    drawLines
    updateList
    
End Sub

Sub doPuzzle()

    Dim puzzW As Double, puzzH As Double
    Dim piecesW As Double, piecesH As Double
    
    ' Piece width and height
    Dim pieceW As Double, pieceH As Double
    
    Dim pX As Double, pY As Double
    Dim x As Double
    Dim y As Double
    
    Dim multi As Double
    
    Dim offset As Double
    Dim pieceTypes As Long
    pieceTypes = 6
    
    ' Temp defs
    puzzW = 5
    puzzH = 5
    piecesW = 4
    piecesH = 4
    
    pieceW = puzzW / piecesW
    pieceH = puzzH / piecesH
    
    ReDim pData(0)
    
    
    Randomize
    For y = 1 To piecesH
        For x = 1 To piecesW
        
            If y = piecesH And x = piecesW Then Exit For
            
            pX = x * pieceW
            pY = (y - 1) * pieceH
            
            ' Go to the middle
            newLine
            
            If x < piecesW Then
                
                addPoint pX, pY
                
                drawPuzzleEdge Int(Rnd * pieceTypes) + 1, True, (Rnd < 0.5), pX, pY, pieceW, pieceH
                
                
            Else
                
            End If
            
            ' Vertical pieces
            pY = pY + pieceH
            
            addPoint pX, pY
            
            If y < piecesH Then
                ' Bottom row doesn't get bottom pieces
                drawPuzzleEdge Int(Rnd * pieceTypes) + 1, False, (Rnd < 0.5), pX, pY, pieceW, pieceH
            End If
        
            addPoint pX - pieceW, pY
            
        
        Next
    Next
    
    ' Cut out the box
    newLine
    addPoint 0, 0
    addPoint puzzW, 0
    addPoint puzzW, puzzH
    addPoint 0, puzzH
    addPoint 0, 0
    
    
    optimizePolys
    zoomToFit
    
    updateList


End Sub

Function drawPuzzleEdge(pShape As Long, isHoriz As Boolean, flipYes As Boolean, _
                        sX As Double, sY As Double, pW As Double, pH As Double)

    ' Define the puzzle shapes
    Dim puzzShapes(8)
    Dim puzzPieces
    Dim i As Long
    Dim pont
    Dim scalar As Double
    Dim offset As Double
    
    ' Shapes are defined horizontally starting from the right and protrusion going up
    ' Diagonal box
    puzzShapes(1) = "40,0; 40,10; 30,20; 50,40; 70,20; 60,10; 60,0; 100,0"
    
    ' Regular box
    puzzShapes(2) = "40,0; 40,10; 30,10; 30,30; 70,30; 70,10; 60,10; 60,0; 100,0"
        
    ' U joint
    puzzShapes(3) = "40,0; 40,10; 30,10; 30,40; 45,40; 45,30; 55,30; 55,40; 70,40; 70,10; 60,10; 60,0; 100,0"
        
    ' Arrow
    puzzShapes(4) = "40,0; 40,10; 30,10; 50,40; 70,10; 60,10; 60,0; 100,0"
    
    ' Edge pieces
    puzzShapes(5) = "10,0; 10,40; 20,40; 20,-40; 10,-50; 50,-50; 40,-40; 40,40; 50,40; 50,0; 100,0"
    
    '
    puzzShapes(6) = "30,0; 70,40; 90,20; 80,10; 70,20; 50, 0"
    
    
    puzzPieces = Split(puzzShapes(pShape), "; ")
    
    scalar = (Rnd * 0.2) + 0.4
    offset = (Rnd * 0.4) + 0.2
    
    
    ' If horizontal, go down
    If isHoriz Then
        For i = 0 To UBound(puzzPieces)
            pont = Split(puzzPieces(i), ",")
            ' Add this point
            addPoint (Val(pont(1)) / 100 * pW * scalar) * IIf(flipYes, -1, 1) + sX, (Val(pont(0)) / 100 * pH * scalar) + sY + (pH * scalar * offset)
        Next
    
    Else
        For i = 0 To UBound(puzzPieces)
            pont = Split(puzzPieces(i), ",")
            ' Add this point
            addPoint ((100 - Val(pont(0))) / 100 * pW * scalar) + sX - pW + (pW * scalar * offset), (Val(pont(1)) / 100 * pH * scalar) * IIf(flipYes, -1, 1) + sY
        Next
    
    
    End If
    

End Function


Function doPages()

    ' Look for a layer called 'Cut Boxes'
    Dim i As Long
    Dim hasCutBoxes As Boolean
    Dim p As Long
    Dim fName As String
    Dim ppX As Double, ppY As Double, ppW As Double, ppH As Double ' Not actually W and H, but X2 and Y2 really.
    Dim pCount As Long ' Page count
    
    Dim j As Long
    Dim n As Long, n2 As Long
    Dim lastPointInside As Boolean
    Dim intersect As pointD
    Dim testPoint As pointD
    Dim testPoint2 As pointD
    Dim pBackup() As typLine
    Dim lastLineLeft As String
    Dim thisLineLeft As String
    Dim thePage() As typLine
    
    Dim fileSavePath As String
    
   On Error GoTo doPages_Error

    For i = 1 To UBound(pData)
        If pData(i).LayerID = "Cut Boxes" Then hasCutBoxes = True: Exit For
    Next
    
    If Not hasCutBoxes Then
        MsgBox "This function will take the shapes on the layer 'Cut Boxes' and will export a GCode file for each one, with only the lines that occur inside that box.  No 'Cut Boxes' layer was found in your document.", vbInformation
        Exit Function
    End If
    
    
    
    ' Ask where to save the files

    With COMDLG
        .FileName = getFolderNameFromPath(.FileName) & "\" & getFileNameNoExten(getFileNameFromPath(.FileName)) & "-PageNN.ngc"
        .Filter = "GCODE Files (*.ngc)|*.ngc"
        .DialogTitle = "Export GCODE - NN will be replaced with page number."
        .ShowSave
        fName = .FileName
    End With
    
    If fName = "" Then Exit Function
    If InStr(1, fName, "NN.ngc", vbTextCompare) = 0 Then
        MsgBox "The filename must end with 'NN.ngc'.  The NN will be replaced with the page number for each page saved.", vbExclamation
        Exit Function
    End If
    
    ' Do each page in Cut Boxes.
    For p = 1 To UBound(pData)
        If pData(p).LayerID = "Cut Boxes" Then

            getPolyBounds p, ppX, ppY, ppW, ppH
            
            pCount = pCount + 1
            
            ' Clear the array
            ReDim thePage(0)
            
            ' Loop through each polygon and copy it onto our page.
            For i = 1 To UBound(pData)
                
                If pData(i).LayerID <> "Cut Boxes" Then
                    n = 0
                    lastPointInside = False
                    lastLineLeft = ""
                    
                    ' Loop through each line segment of the polygon. If the line segment fits onto the page, copy it into our new polygon.
                    For j = 1 To UBound(pData(i).Points)
                        
                    
                        With pData(i).Points(j)
                            If .x >= ppX And .x <= ppW And .y >= ppY And .y <= ppH Then
                                
                                If n = 0 Then
                                    ' Create a new poly
                                    n = UBound(thePage) + 1
                                    ReDim Preserve thePage(n)
                                    ReDim thePage(n).Points(0)
                                End If
                                
                                ' Was the last point NOT inside?
                                If lastPointInside = False And j > 1 Then ' Also, this can't be the first point.
                                
                                    ' Test 1: Top segment
                                    testPoint.x = ppX:  testPoint.y = ppY: testPoint2.x = ppW: testPoint2.y = ppY
                                    intersect = lineIntersectLine(pData(i).Points(j - 1), pData(i).Points(j), testPoint, testPoint2)
                                    thisLineLeft = "T"
                                    If intersect.x = -6666 Then
                                        
                                        ' Right side
                                        testPoint.x = ppW:  testPoint.y = ppY: testPoint2.x = ppW: testPoint2.y = ppH
                                        intersect = lineIntersectLine(pData(i).Points(j - 1), pData(i).Points(j), testPoint, testPoint2)
                                        thisLineLeft = "R"
                                        If intersect.x = -6666 Then
                                            ' Bottom
                                            testPoint.x = ppX:  testPoint.y = ppH: testPoint2.x = ppW: testPoint2.y = ppH
                                            intersect = lineIntersectLine(pData(i).Points(j - 1), pData(i).Points(j), testPoint, testPoint2)
                                            thisLineLeft = "B"
                                            If intersect.x = -6666 Then
                                                ' Left
                                                testPoint.x = ppX:  testPoint.y = ppY: testPoint2.x = ppX: testPoint2.y = ppH
                                                intersect = lineIntersectLine(pData(i).Points(j - 1), pData(i).Points(j), testPoint, testPoint2)
                                                thisLineLeft = "L"
                                            End If
                                        End If
                                    End If
                                    
                                    If intersect.x <> -6666 Then ' Did intersect.
                                        If (lastLineLeft = "T" And thisLineLeft = "L") Or (lastLineLeft = "L" And thisLineLeft = "T") Then addPoint3 thePage(n), ppX, ppY
                                        If (lastLineLeft = "L" And thisLineLeft = "B") Or (lastLineLeft = "B" And thisLineLeft = "L") Then addPoint3 thePage(n), ppX, ppH
                                        If (lastLineLeft = "B" And thisLineLeft = "R") Or (lastLineLeft = "R" And thisLineLeft = "B") Then addPoint3 thePage(n), ppW, ppH
                                        If (lastLineLeft = "T" And thisLineLeft = "R") Or (lastLineLeft = "R" And thisLineLeft = "T") Then addPoint3 thePage(n), ppW, ppY
                                        addPoint2 thePage(n), intersect
                                    End If
                                End If
                                addPoint2 thePage(n), pData(i).Points(j)
                                lastPointInside = True
                            Else
                                
                                ' Was the point previous to this one inside?
                                If lastPointInside Then
                                    ' Figure out where the line between this point and the last point intersects the borders.
                                    
                                    ' Which segment did it intersect?
                                    
                                    ' Test 1: Top segment
                                    testPoint.x = ppX:  testPoint.y = ppY: testPoint2.x = ppW: testPoint2.y = ppY
                                    intersect = lineIntersectLine(pData(i).Points(j - 1), pData(i).Points(j), testPoint, testPoint2)
                                    lastLineLeft = "T"
                                    If intersect.x = -6666 Then
                                        
                                        ' Right side
                                        testPoint.x = ppW:  testPoint.y = ppY: testPoint2.x = ppW: testPoint2.y = ppH
                                        intersect = lineIntersectLine(pData(i).Points(j - 1), pData(i).Points(j), testPoint, testPoint2)
                                        lastLineLeft = "R"
                                        If intersect.x = -6666 Then
                                            ' Bottom
                                            testPoint.x = ppX:  testPoint.y = ppH: testPoint2.x = ppW: testPoint2.y = ppH
                                            intersect = lineIntersectLine(pData(i).Points(j - 1), pData(i).Points(j), testPoint, testPoint2)
                                            lastLineLeft = "B"
                                            If intersect.x = -6666 Then
                                                ' Left
                                                testPoint.x = ppX:  testPoint.y = ppY: testPoint2.x = ppX: testPoint2.y = ppH
                                                intersect = lineIntersectLine(pData(i).Points(j - 1), pData(i).Points(j), testPoint, testPoint2)
                                                lastLineLeft = "L"
                                            End If
                                        End If
                                    End If
                                    
                                    If intersect.x <> -6666 Then ' Did intersect.
                                        addPoint2 thePage(n), intersect
                                    End If
                                
                                End If
                                lastPointInside = False
                                
                            End If
                            
                        End With
                    Next
                End If
            Next
            
            ' Add the page outline to the document.
            n = UBound(thePage) + 1
            ReDim Preserve thePage(n)
            ReDim thePage(n).Points(0)
            
            addPoint3 thePage(n), ppX, ppY
            addPoint3 thePage(n), ppX, ppH
            addPoint3 thePage(n), ppW, ppH
            addPoint3 thePage(n), ppW, ppY
            addPoint3 thePage(n), ppX, ppY
            
            ' Backup the shapes
            pBackup = pData
            ' save out just this page
            pData = thePage
            
            ' Subtract the X and Y
            For i = 1 To UBound(pData)
                For j = 1 To UBound(pData(i).Points)
                    With pData(i).Points(j)
                        .x = .x - ppX
                        .y = .y - ppY
                    End With
                Next
            Next
            
            ' If the shape is taller than our laser, rotate it 90 degrees
            If ppH - ppY > 17.5 Then ' Inches
                For i = 1 To UBound(pData)
                    For j = 1 To UBound(pData(i).Points)
                        With pData(i).Points(j)
                            Swap pData(i).Points(j).x, pData(i).Points(j).y
                            ' Invert the Y
                            pData(i).Points(j).y = (ppW - ppX) - pData(i).Points(j).y
                        End With
                    Next
                Next
            End If
            
            
            
            ' Show it on screen so you can see what is being saved.
            zoomToFit
            
            fileSavePath = Replace(fName, "NN.ngc", Format(pCount, "000") & ".ngc", , , vbTextCompare)
            exportGCODE fileSavePath, Val(cFeedRate.Text), TB1.ButtonChecked("zplunge")
            'MsgBox "Page " & pCount & " - Shape " & p
            
            DoEvents
            Sleep 100
            
            ' Restore the data
            pData = pBackup
            'If pCount = 3 Then Exit Function

        End If
    Next
    
    
    zoomToFit


   On Error GoTo 0
   Exit Function

doPages_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure doPages of Form frmInterface", vbCritical, App.ProductName & " ERROR"
    addToLog "[ERROR]", "In doPages of Form frmInterface", Err.Number, Err.Description

    
End Function


Function doPagesOld()
    ' Split the document up into a series of pages.
    
    Dim pageCountX As Long, pageCountY As Long
    Dim maxX As Double, maxY As Double
    Dim pX As Long, pY As Long
    Dim thePage() As typLine
    
    Dim ppX As Double, ppY As Double, ppW As Double, ppH As Double
    Dim i As Long
    Dim j As Long
    Dim n As Long, n2 As Long
    Dim lastPointInside As Boolean
    Dim intersect As pointD
    Dim testPoint As pointD
    Dim testPoint2 As pointD
    Dim pCount As Long
    Dim pBackup() As typLine
    
    Dim lastLineLeft As String
    Dim thisLineLeft As String
    
    
    
    ' Configure for 11x17 paper
    Const pageW = 99
    Const pageH = 250
    
    ' How many pages will we need?
    
    getExtents maxX, maxY
    
    pageCountX = -Int(-maxX / pageW)
    pageCountY = -Int(-maxY / pageH)
    
    If MsgBox("With " & pageW & " by " & pageH & " paper, this document will require " & pageCountX & " by " & pageCountY & " pages, or " & (pageCountX * pageCountY) & " sheets.", vbYesNo) = vbYes Then
        
        
        For pY = 1 To pageCountY
            For pX = 1 To pageCountX
        
                ' Calculate the bounds of this page.
                ppX = (pX - 1) * pageW
                ppY = (pY - 1) * pageH
                ppW = ppX + pageW
                ppH = ppY + pageH
                
                pCount = pCount + 1
                
                ' Clear the array
                ReDim thePage(0)
                
                ' Loop through each polygon and copy it onto our page.
                For i = 1 To UBound(pData)
                    
                    n = 0
                    lastPointInside = False
                    lastLineLeft = ""
                    
                    ' Loop through each line segment of the polygon. If the line segment fits onto the page, copy it into our new polygon.
                    For j = 1 To UBound(pData(i).Points)
                    
                        With pData(i).Points(j)
                            
                            
                            
                            If .x >= ppX And .x <= ppW And .y >= ppY And .y <= ppH Then
                                
                                If n = 0 Then
                                    ' Create a new poly
                                    n = UBound(thePage) + 1
                                    ReDim Preserve thePage(n)
                                    ReDim thePage(n).Points(0)
                                End If
                                
                                ' Was the last point NOT inside?
                                If lastPointInside = False And j > 1 Then ' Also, this can't be the first point.
                                
                                    ' Test 1: Top segment
                                    testPoint.x = ppX:  testPoint.y = ppY: testPoint2.x = ppW: testPoint2.y = ppY
                                    intersect = lineIntersectLine(pData(i).Points(j - 1), pData(i).Points(j), testPoint, testPoint2)
                                    thisLineLeft = "T"
                                    If intersect.x = -6666 Then
                                        
                                        ' Right side
                                        testPoint.x = ppW:  testPoint.y = ppY: testPoint2.x = ppW: testPoint2.y = ppH
                                        intersect = lineIntersectLine(pData(i).Points(j - 1), pData(i).Points(j), testPoint, testPoint2)
                                        thisLineLeft = "R"
                                        If intersect.x = -6666 Then
                                            ' Bottom
                                            testPoint.x = ppX:  testPoint.y = ppH: testPoint2.x = ppW: testPoint2.y = ppH
                                            intersect = lineIntersectLine(pData(i).Points(j - 1), pData(i).Points(j), testPoint, testPoint2)
                                            thisLineLeft = "B"
                                            If intersect.x = -6666 Then
                                                ' Left
                                                testPoint.x = ppX:  testPoint.y = ppY: testPoint2.x = ppX: testPoint2.y = ppH
                                                intersect = lineIntersectLine(pData(i).Points(j - 1), pData(i).Points(j), testPoint, testPoint2)
                                                thisLineLeft = "L"
                                            End If
                                        End If
                                    End If
                                    
                                    If intersect.x <> -6666 Then ' Did intersect.
                                    
                                        
                                        If (lastLineLeft = "T" And thisLineLeft = "L") Or (lastLineLeft = "L" And thisLineLeft = "T") Then addPoint3 thePage(n), ppX, ppY
                                        If (lastLineLeft = "L" And thisLineLeft = "B") Or (lastLineLeft = "B" And thisLineLeft = "L") Then addPoint3 thePage(n), ppX, ppH
                                        If (lastLineLeft = "B" And thisLineLeft = "R") Or (lastLineLeft = "R" And thisLineLeft = "B") Then addPoint3 thePage(n), ppW, ppH
                                        If (lastLineLeft = "T" And thisLineLeft = "R") Or (lastLineLeft = "R" And thisLineLeft = "T") Then addPoint3 thePage(n), ppW, ppY
                                        
                                        addPoint2 thePage(n), intersect
                                    End If
                                
                                End If
                                
                                
                                addPoint2 thePage(n), pData(i).Points(j)
                                
                                lastPointInside = True
                            Else
                                
                                ' Was the point previous to this one inside?
                                If lastPointInside Then
                                    ' Figure out where the line between this point and the last point intersects the borders.
                                    
                                    ' Which segment did it intersect?
                                    
                                    ' Test 1: Top segment
                                    testPoint.x = ppX:  testPoint.y = ppY: testPoint2.x = ppW: testPoint2.y = ppY
                                    intersect = lineIntersectLine(pData(i).Points(j - 1), pData(i).Points(j), testPoint, testPoint2)
                                    lastLineLeft = "T"
                                    If intersect.x = -6666 Then
                                        
                                        ' Right side
                                        testPoint.x = ppW:  testPoint.y = ppY: testPoint2.x = ppW: testPoint2.y = ppH
                                        intersect = lineIntersectLine(pData(i).Points(j - 1), pData(i).Points(j), testPoint, testPoint2)
                                        lastLineLeft = "R"
                                        If intersect.x = -6666 Then
                                            ' Bottom
                                            testPoint.x = ppX:  testPoint.y = ppH: testPoint2.x = ppW: testPoint2.y = ppH
                                            intersect = lineIntersectLine(pData(i).Points(j - 1), pData(i).Points(j), testPoint, testPoint2)
                                            lastLineLeft = "B"
                                            If intersect.x = -6666 Then
                                                ' Left
                                                testPoint.x = ppX:  testPoint.y = ppY: testPoint2.x = ppX: testPoint2.y = ppH
                                                intersect = lineIntersectLine(pData(i).Points(j - 1), pData(i).Points(j), testPoint, testPoint2)
                                                lastLineLeft = "L"
                                            End If
                                        End If
                                    End If
                                    
                                    If intersect.x <> -6666 Then ' Did intersect.
                                        addPoint2 thePage(n), intersect
                                    End If
                                
                                End If
                                lastPointInside = False
                                
                            End If
                            
                        End With
                    Next
                Next
                
                ' Backup the shapes
                pBackup = pData
                
                
                ' save out just this page
                pData = thePage
                
                ' Subtract the X and Y
                For i = 1 To UBound(pData)
                    For j = 1 To UBound(pData(i).Points)
                        With pData(i).Points(j)
                            .x = .x - ppX
                            .y = .y - ppY
                        End With
                    Next
                Next
                
                zoomToFit
                
                'MsgBox "This is page " & pCount
                
                exportGCODE "e:\temp\sign\page" & Format(pCount, "00") & ".ngc", Val(cFeedRate.Text), TB1.ButtonChecked("zplunge")
                
                DoEvents
                Sleep 100
                
                
                
                
                
                ' Restore the data
                pData = pBackup
                'If pCount = 3 Then Exit Function
                
            Next
        Next
    End If
    

End Function

Private Function addPoint2(theLine As typLine, thePoint As pointD)
    Dim n2 As Long
    
    n2 = UBound(theLine.Points) + 1
    ReDim Preserve theLine.Points(n2)
    
    ' Copy this point
    theLine.Points(n2) = thePoint


End Function

Private Function addPoint3(theLine As typLine, pX As Double, pY As Double)
    Dim n2 As Long
    
    n2 = UBound(theLine.Points) + 1
    ReDim Preserve theLine.Points(n2)
    
    ' Copy this point
    theLine.Points(n2).x = pX
    theLine.Points(n2).y = pY


End Function
'
'Function doBox()
'
'    Dim bWid As Double, bHig As Double, bDep As Double
'    Dim matThick As Double
'    Dim i As Long
'    Dim nubSize As Double
'    Dim shortestSide As Double
'    Dim xOff As Double, yOff As Double
'    Dim X As Double, Y As Double, n As Double
'    Dim nubCount As Double
'    Dim realNubSize As Double
'
'    ReDim pData(0)
'
'    ' Generate the pieces to make a box
'    ' 0.125
'
'    bWid = Val(InputBox("Box Width?", "Box", 4))
'    bHig = Val(InputBox("Box Height?", "Box", 4))
'    bDep = Val(InputBox("Box Depth?", "Box", 4))
'    matThick = Val(InputBox("Material Thickness?", "Box", 0.125))
'
'    ' What will be the size of the "nubbles"?
'    shortestSide = Min(bWid, Min(bHig, bDep))
'
'    ' Make it 5 nubbles for the shortest side.
'    nubSize = shortestSide / 10
'
'    ' Top side
'    newLine
'    xOff = 1
'    yOff = 1
'
'    ' How many nubbles will fit on this side?
'    nubCount = Int(CDbl(bHig / nubSize)) ' round down
'    If nubCount Mod 2 = 0 Then nubCount = nubCount + 1
'
'
'    ' Expand the nubbles to a whole number.
'    realNubSize = (bHig / nubCount)
'
'    addPoint xOff, yOff
'
'    ' Side 1 (Height)
'    doNubs matThick, nubCount, realNubSize, xOff, yOff, 0
'
'    ' Bottom (Width)
'
'
'
'
'
'
'
'    ' Generate four side pieces.
'    For i = 1 To 4
'
'
'
'
'
'    Next
'
'
'
'    zoomToFit
'
'    updateList
'
'
'End Function
'
'Function doNubs(matThick As Double, nubCount As Double, realNubSize As Double, _
'                xOff As Double, yOff As Double, _
'                direction As Long)
'
'    Dim i As Long
'    For n = 1 To nubCount
'        If n Mod 2 = 0 Then ' Down point
'            addPoint xOff + (n * realNubSize), yOff + matThick
'            addPoint xOff + (n * realNubSize), yOff
'        Else ' Up point
'            addPoint xOff + (n * realNubSize), yOff
'            addPoint xOff + (n * realNubSize), yOff + matThick
'        End If
'    Next
'
'
'
'End Function


Public Function goTile(nRows As Long, nCols As Long, wOff As Double, hOff As Double, rowDiff As Double, colDiff As Double)
    
    ' Tile the shape
    
    Dim maxX As Double
    Dim maxY As Double
    Dim upTo As Long
    Dim count As Long
    
    upTo = UBound(pData)
    
    getExtents maxX, maxY
    
    
    
    Dim x As Long, y As Long
    For y = 1 To nRows
        For x = 1 To nCols
            count = count + 1
            
            If count > 1 Then ' skpi the first one
                ' Copy the shapes.
                duplicateShapes upTo, ((maxX + wOff) * (x - 1)) + (((y - 1) Mod 2) * colDiff), _
                    ((maxY + hOff) * (y - 1)) + (((x - 1) Mod 2) * rowDiff)
            End If
        Next
    Next
    
    optimizePolys
    zoomToFit
    
    updateList
End Function

Function duplicateShapes(endAt As Long, Xadd As Double, Yadd As Double)

    Dim i As Long
    Dim j As Long
    
    Dim n As Long
    For i = 1 To endAt
        
        n = UBound(pData) + 1
        ReDim Preserve pData(n)
        
        pData(n) = pData(i)
        
        For j = 1 To UBound(pData(n).Points)
            pData(n).Points(j).x = pData(n).Points(j).x + Xadd
            pData(n).Points(j).y = pData(n).Points(j).y + Yadd
        Next
    Next

End Function
