Attribute VB_Name = "Bezier"
Option Explicit




' Parametric functions for drawing a degree 3 Bezier curve.
Private Function bezX(ByVal t As Double, ByVal x0 As Double, _
    ByVal x1 As Double, ByVal X2 As Double, ByVal x3 As _
    Double) As Double
    bezX = CSng( _
        x0 * (1 - t) ^ 3 + _
        x1 * 3 * t * (1 - t) ^ 2 + _
        X2 * 3 * t ^ 2 * (1 - t) + _
        x3 * t ^ 3 _
        )
End Function

Private Function bezY(ByVal t As Double, ByVal y0 As Double, _
    ByVal y1 As Double, ByVal y2 As Double, ByVal y3 As _
    Double) As Double
    bezY = CSng( _
        y0 * (1 - t) ^ 3 + _
        y1 * 3 * t * (1 - t) ^ 2 + _
        y2 * 3 * t ^ 2 * (1 - t) + _
        y3 * t ^ 3 _
        )
End Function

' Draw the Bezier curve.
Public Sub DrawBezier(gr As PictureBox, dt As Double, pt0 As pointD, _
    pt1 As pointD, pt2 As pointD, pt3 As pointD)
    ' Debugging code commented out.
    ' Draw the control lines.
    gr.ForeColor = vbRed
    gr.Line (pt0.x, pt0.y)-(pt1.x, pt1.y)
    gr.Line (pt1.x, pt1.y)-(pt2.x, pt2.y)
    gr.Line (pt2.x, pt2.y)-(pt3.x, pt3.y)
    
    gr.ForeColor = vbBlack
    
    
    ' Draw the curve.
    Dim t, x0, y0, x1, y1 As Double

    t = 0
    x1 = bezX(t, pt0.x, pt1.x, pt2.x, pt3.x)
    y1 = bezY(t, pt0.y, pt1.y, pt2.y, pt3.y)
    t = t + dt
    Do While t < 1#
        x0 = x1
        y0 = y1
        x1 = bezX(t, pt0.x, pt1.x, pt2.x, pt3.x)
        y1 = bezY(t, pt0.y, pt1.y, pt2.y, pt3.y)
        gr.Line (x0, y0)-(x1, y1)
        t = t + dt
    Loop

    ' Connect to the final point.
    t = 1#
    x0 = x1
    y0 = y1
    x1 = bezX(t, pt0.x, pt1.x, pt2.x, pt3.x)
    y1 = bezY(t, pt0.y, pt1.y, pt2.y, pt3.y)
    gr.Line (x0, y0)-(x1, y1)
End Sub

Public Sub AddBezier(dt As Double, pt0 As pointD, _
    pt1 As pointD, pt2 As pointD, pt3 As pointD)
    ' Draw the curve.
    Dim t As Double, x0 As Double, y0 As Double, x1 As Double, y1 As Double

    t = 0
    x1 = bezX(t, pt0.x, pt1.x, pt2.x, pt3.x)
    y1 = bezY(t, pt0.y, pt1.y, pt2.y, pt3.y)
    t = t + dt
    
    addPoint x1, y1
    Do While t < 1#
        x0 = x1
        y0 = y1
        x1 = bezX(t, pt0.x, pt1.x, pt2.x, pt3.x)
        y1 = bezY(t, pt0.y, pt1.y, pt2.y, pt3.y)
        'gr.Line (x0, y0)-(x1, y1)
        
        addPoint x1, y1
        
        
        t = t + dt
    Loop

    ' Connect to the final point.
    t = 1#
    x0 = x1
    y0 = y1
    x1 = bezX(t, pt0.x, pt1.x, pt2.x, pt3.x)
    y1 = bezY(t, pt0.y, pt1.y, pt2.y, pt3.y)
    'gr.Line (x0, y0)-(x1, y1)
    
    addPoint x1, y1
    
End Sub

Public Sub AddQuadBezier(dt As Double, pt0 As pointD, pt1 As pointD, pt2 As pointD)


    'Protected i
    'Protected.f t, t1, A, b, c, D
    Dim i As Long
    Dim t1 As Double
    Dim A As Double, b As Double, c As Double, D As Double
    Dim t As Double
    
    t = 0
    Do While t < 1#
        t1 = 1# - t
        A = t1 ^ 2
        b = 2# * t * t1
        c = t ^ 2
        
        addPoint A * pt0.x + b * pt1.x + c * pt2.x, A * pt0.y + b * pt1.y + c * pt2.y
        t = t + dt
    Loop
    
    ' One more at t = 1
    t = 1
    t1 = 1# - t
    A = t1 ^ 2
    b = 2# * t * t1
    c = t ^ 2
    
    addPoint A * pt0.x + b * pt1.x + c * pt2.x, A * pt0.y + b * pt1.y + c * pt2.y
    
    
End Sub



