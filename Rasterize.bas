Attribute VB_Name = "Rasterize"
Option Explicit

Public Const GREYLEVELS = 8


' Modules to control rasterization

Function rasterFile(inFile As String)
    ' Raster this file with beta greyscale support
    
    Dim p As PictureBox
    Dim pW As Long, pH As Long
    Dim scalar As Double
    Dim X As Long, Y As Long
    Dim RY As Double
    
    Dim c As Long
    Dim crr As typRGB
    Dim grey As Long
    
    ' How many levels of greyscale are there? Let's do GREYLEVELS.
    Dim lastColor As Long
    Dim lastX As Long
    
    
    Set p = frmInterface.Picture2
    
    Set p.Picture = LoadPicture(inFile)
    
    ' Get the width and height and calculate a scaler
    pW = p.Width
    pH = p.Height
    
    ' Desired width: 5 inches
    scalar = 5 / pW
    
    ' Now scan the picture from left to right and build the lines
    For RY = 0 To pH Step 0.5
        Y = CLng(RY)
        newLine
        addPoint 0, RY * scalar
        lastColor = -1
                
        For X = 0 To pW
            ' Get the color of this point
            c = p.Point(X, Y)
            
            ' Convert to greyscale
            crr = convertVBtoRGB(c)
            grey = ((crr.R + crr.G + crr.b) / 3) * (GREYLEVELS / 255) ' Convert to 0 to GREYLEVELS
                        
            frmInterface.Picture1.PSet (X, Y), RGB(grey * (255 / GREYLEVELS), grey * (255 / GREYLEVELS), grey * (255 / GREYLEVELS))
            
                        
            ' Draw a line to this point
            If grey <> lastColor Then
                If lastColor <> -1 Then
                    addPoint X * scalar, RY * scalar
                    pData(currentLine).greyLevel = lastColor
                
                    newLine
                    addPoint X * scalar, RY * scalar
                End If
                
                lastColor = grey
            End If
            
        Next
    
        ' Finish last line
        If lastColor <> -1 Then
            addPoint (X - 1) * scalar, RY * scalar
            pData(currentLine).greyLevel = lastColor
        End If
    Next
    
    Debug.Print "dpone"
    

End Function
