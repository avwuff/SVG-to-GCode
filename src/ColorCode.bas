Attribute VB_Name = "ColorCode"
Option Explicit
' Routines for converting colours between HSB and RGB, etc.

Public Type typRGB
    R As Integer
    G As Integer
    b As Integer
End Type

Public Type typXYZ
    X As Double
    Y As Double
    z As Double
End Type

Public Type typLAB
    L As Double
    A As Double
    b As Double
End Type

Public Type typHSL
    H As Double
    S As Double
    L As Double
End Type

Public Type typCMYK
    c As Long
    M As Long
    Y As Long
    K As Long
End Type

Public Type typPantone
    PantoneName As String
    Color As Long
End Type

Public pantoneList() As typPantone

Function makeHTMLColor(c As Long) As String
    ' Turns the colour into HTML HEX code.
    Dim R As Long, G As Long, b As Long
    
    R = c Mod 256
    G = (c \ 256) Mod 256
    b = c \ 256 \ 256
    
    makeHTMLColor = fixZeros(Hex(R)) & fixZeros(Hex(G)) & fixZeros(Hex(b))
End Function

Function fixZeros(inSt As String) As String
    ' Adds a 0 to the front if needed.
    fixZeros = inSt
    If Len(fixZeros) = 1 Then fixZeros = "0" & fixZeros
End Function

Function convertVBtoRGB(c As Long) As typRGB
    ' Convert vb color to RGB
    
    convertVBtoRGB.R = c Mod 256
    convertVBtoRGB.G = (c \ 256) Mod 256
    convertVBtoRGB.b = c \ 256 \ 256

End Function


Function RGBtoHSL(inRGB As typRGB) As typHSL

    Dim nTemp As Double
    Dim lMin As Double
    Dim lMax As Double
    Dim lDelta As Double
    
    Dim H As Double, S As Double, L As Double
    Dim R As Double, G As Double, b As Double
     
    R = inRGB.R
    G = inRGB.G
    b = inRGB.b
    
     
    If R > G Then
        If R > b Then
            lMax = R
        Else
            lMax = b
        End If
    Else
        If G > b Then
            lMax = G
        Else
            lMax = b
        End If
    End If
    
    If R < G Then
        If R < b Then
            lMin = R
        Else
            lMin = b
        End If
    Else
        If G < b Then
            lMin = G
        Else
            lMin = b
        End If
    End If
    
    lDelta = lMax - lMin
      
    L = (lMax * 100) / 255
      
    If lMax > 0 Then
        S = (lDelta / lMax) * 100
        If lDelta > 0 Then
            If lMax = R Then
                nTemp = (G - b) / lDelta
            ElseIf lMax = G Then
                nTemp = 2 + (b - R) / lDelta
            Else
                nTemp = 4 + (R - G) / lDelta
            End If
            
            H = nTemp * 60
            If H < 0 Then H = H + 360
        End If
    End If

    RGBtoHSL.H = Int(H)
    RGBtoHSL.S = Int(S)
    RGBtoHSL.L = Int(L)
    
End Function


Function HSLtoRGB(inHSL As typHSL) As typRGB


    Dim R As Double, G As Double, b As Double
    Dim nH As Double, nS As Double, nB As Double
    Dim lH As Double
    Dim nF As Double, nP As Double, nQ As Double, nT As Double
    
    nH = IIf(inHSL.H = 360, 0, inHSL.H) / 60
    nS = inHSL.S / 100
    nB = inHSL.L / 100
    
    lH = Int(nH)
    nF = nH - lH
    nP = nB * (1 - nS)
    nQ = nB * (1 - nS * nF)
    nT = nB * (1 - nS * (1 - nF))
    
    If lH = 0 Then
        R = nB * 255
        G = nT * 255
        b = nP * 255
    ElseIf lH = 1 Then
        R = nQ * 255
        G = nB * 255
        b = nP * 255
    ElseIf lH = 2 Then
        R = nP * 255
        G = nB * 255
        b = nT * 255
    ElseIf lH = 3 Then
        R = nP * 255
        G = nQ * 255
        b = nB * 255
    ElseIf lH = 4 Then
        R = nT * 255
        G = nP * 255
        b = nB * 255
    ElseIf lH = 5 Then
        R = nB * 255
        G = nP * 255
        b = nQ * 255
    Else
        R = (nB * 255) / 100
        G = R
        b = R
    End If
    
    HSLtoRGB.R = CByte(Min(Max(R, 0), 255))
    HSLtoRGB.G = CByte(Min(Max(G, 0), 255))
    HSLtoRGB.b = CByte(Min(Max(b, 0), 255))
End Function

Function convertRGBToVB(inRGB As typRGB) As Long
    convertRGBToVB = RGB(inRGB.R, inRGB.G, inRGB.b)
End Function

Function RGBtoXYZ(inRGB As typRGB) As typXYZ
    ' Convert RGB color space to XYZ
    Dim var_R As Double
    Dim var_G As Double
    Dim var_B As Double
    
    var_R = (inRGB.R / 255)          'Where R = 0 ÷ 255
    var_G = (inRGB.G / 255)          'Where G = 0 ÷ 255
    var_B = (inRGB.b / 255)          'Where B = 0 ÷ 255
    
    If var_R > 0.04045 Then
        var_R = ((var_R + 0.055) / 1.055) ^ 2.4
    Else
        var_R = var_R / 12.92
    End If
    If var_G > 0.04045 Then
        var_G = ((var_G + 0.055) / 1.055) ^ 2.4
    Else
        var_G = var_G / 12.92
    End If
    If var_B > 0.04045 Then
        var_B = ((var_B + 0.055) / 1.055) ^ 2.4
    Else
        var_B = var_B / 12.92
    End If
    
    var_R = var_R * 100
    var_G = var_G * 100
    var_B = var_B * 100
    
    'Observer. = 2°, Illuminant = D65
    RGBtoXYZ.X = var_R * 0.4124 + var_G * 0.3576 + var_B * 0.1805
    RGBtoXYZ.Y = var_R * 0.2126 + var_G * 0.7152 + var_B * 0.0722
    RGBtoXYZ.z = var_R * 0.0193 + var_G * 0.1192 + var_B * 0.9505

End Function

Function XYZtoLAB(inXYZ As typXYZ) As typLAB
    ' Convert XYZ color space to Hunter-Lab
    XYZtoLAB.L = 10 * Sqr(inXYZ.Y)
    
    If inXYZ.Y <> 0 Then XYZtoLAB.A = 17.5 * (((1.02 * inXYZ.X) - inXYZ.Y) / Sqr(inXYZ.Y))
    If inXYZ.Y <> 0 Then XYZtoLAB.b = 7 * ((inXYZ.Y - (0.847 * inXYZ.z)) / Sqr(inXYZ.Y))

End Function

Function loadPantone(inFile As String)
    Dim G As String
    Dim f As Long
    Dim i As Long
    Dim n As Long
    Dim X, X2
    Dim c As String
    
    ReDim pantoneList(0)
    
    If myDir(inFile) <> "" Then
        
        
        G = Space(FileLen(inFile))
        f = FreeFile
        
        Open inFile For Binary As f
            Get #f, , G
        Close f
        
        ' Split and translate
        X = Split(G, vbCrLf)
        
        For i = 0 To UBound(X)
        
            If X(i) <> "" Then
                X2 = Split(X(i), Chr(9)) 'split by tabs
                
                n = UBound(pantoneList) + 1
                ReDim Preserve pantoneList(n)
                
                With pantoneList(n)
                    c = Trim(X2(1))
                    If Len(c) > 6 Then c = Right(c, 6)
                    
                    .PantoneName = Trim(X2(0))
                    .Color = makeVBColor(c)
                End With
            End If
        Next
    End If

    addToLog UBound(pantoneList), " pantone colors loaded"
    

End Function

Function RGBtoCMYK(inRGB As typRGB) As typCMYK
    
    Dim c As Double, M As Double, Y As Double, K As Double
    
    c = 1 - (inRGB.R / 255)
    M = 1 - (inRGB.G / 255)
    Y = 1 - (inRGB.b / 255)
    K = Min(c, Min(M, Y))
    
    c = Min(1, Max(0, c - K))
    M = Min(1, Max(0, M - K))
    Y = Min(1, Max(0, Y - K))
    K = Min(1, Max(0, K))
    
    RGBtoCMYK.c = c * 100
    RGBtoCMYK.M = M * 100
    RGBtoCMYK.Y = Y * 100
    RGBtoCMYK.K = K * 100
    
End Function

Function CMYKtoRGB(inCMYK As typCMYK) As typRGB
    CMYKtoRGB.R = ((1 - (inCMYK.K / 100)) * (1 - (inCMYK.c / 100))) * 255
    CMYKtoRGB.G = ((1 - (inCMYK.K / 100)) * (1 - (inCMYK.M / 100))) * 255
    CMYKtoRGB.b = ((1 - (inCMYK.K / 100)) * (1 - (inCMYK.Y / 100))) * 255
End Function

Function textColorBasedOnBackground(inC As Long) As Long
    ' Should the text color be white or black based on the brightness of the background
    Dim myL As typLAB
    Dim myH As typHSL
    Dim myR As typRGB
    myR = convertVBtoRGB(inC)
    myL = XYZtoLAB(RGBtoXYZ(myR))
    myH = RGBtoHSL(myR)
    textColorBasedOnBackground = IIf(myH.L < 60 Or myL.L < 62, vbWhite, vbBlack)
End Function

Function CMYKtoHex(inCMYK As typCMYK) As String
    CMYKtoHex = fixZeros(Hex(inCMYK.c)) & fixZeros(Hex(inCMYK.M)) & fixZeros(Hex(inCMYK.Y)) & fixZeros(Hex(inCMYK.K))
End Function

Function HextoCMYK(inHex As String) As typCMYK
    ' Turns 64646464 into the CMYK color
    
    If Len(inHex) < 8 Then
        inHex = Replace(Space(8 - Len(inHex)), " ", "0") & inHex
    End If
    
    If Len(inHex) = 8 Then
        HextoCMYK.c = Val("&H" & Mid(inHex, 1, 2))
        HextoCMYK.M = Val("&H" & Mid(inHex, 3, 2))
        HextoCMYK.Y = Val("&H" & Mid(inHex, 5, 2))
        HextoCMYK.K = Val("&H" & Mid(inHex, 7, 2))
    End If

End Function

