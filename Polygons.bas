Attribute VB_Name = "Polygons"
Option Explicit
Public Type POINTAPI
  X As Long
  Y As Long
End Type


Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long


Function lineIntersectPoly(a As pointD, B As pointD, polyID As Long) As pointD()
    
    Dim An As Long, Bn As Long
    
    Dim cX As Double, cY As Double
    
    Dim K As Long, j As Long
    Dim result() As pointD
    ReDim result(0)
    
    An = 1
    Bn = 1
    
    'Dim pointList As New Scripting.Dictionary
    

    'pa.push(pa[0]); ' add itself to the end?
    
    'result.intersects = false;
    'result.intersections=[];
    'result.start_inside=false;
    'result.end_inside=false;
    
                
    doPoly polyID, a, B, result
    
    lineIntersectPoly = result
    
    
End Function

Private Function doPoly(polyID, a As pointD, B As pointD, result() As pointD)
    Dim c As pointD, D As pointD, i As pointD
    Dim n As Long, n2 As Long
    Dim cl As Collection
    Dim K As Long
    
    With pData(polyID)
        n = UBound(.Points) ' Set n to the last item
        Do While n > 0
            c = .Points(n)
                If n = 1 Then
                    D = .Points(UBound(.Points))
                Else
                    D = .Points(n - 1)
                End If
                i = lineIntersectLine(a, B, c, D)
                If i.X <> -6666 Then
                    
                    n2 = UBound(result) + 1
                    ReDim Preserve result(n2)
                    result(n2) = i
                End If
                
                'If lineIntersectLine(A, newPoint(C.X + D.X, A.Y), C, D).X <> -6666 Then
                '    An = An + 1
                'End If
                'If lineIntersectLine(b, newPoint(C.X + D.X, b.Y), C, D).X <> -6666 Then
                '    Bn = Bn + 1
                'End If
            n = n - 1
        Loop
        
        'If An Mod 2 = 0 Then
        '    'result.start_inside=true;
        'End If
        'If Bn Mod 2 = 0 Then
        '    'result.end_inside=true;
        'End If
        'result.centroid=new Point(cx/(pa.length-1),cy/(pa.length-1));
        'result.intersects = result.intersections.length > 0;
        'return result;
        
        ' Do my kids
        If containList.Exists(polyID) Then Set cl = containList(polyID) ' A list of polygons that I contain
        If Not cl Is Nothing Then
            For K = 1 To cl.Count
                doPoly cl(K), a, B, result
            Next
        End If
    End With
End Function

Function lineIntersectLine(a As pointD, B As pointD, e As pointD, f As pointD, Optional as_seg As Boolean = True) As pointD
    Dim ip As pointD
    Dim a1 As Double, a2 As Double, b1 As Double, b2 As Double, c1 As Double, c2 As Double
    Dim denom As Double
 
    lineIntersectLine.X = -6666 ' Instead of returning null, we return this to indicate no intersection
    
    ' This is a hack, but it does the job. If the line falls on one of my vertices, move it slightly, since unpredictable results occur.
    
    If e.Y = a.Y Then a.Y = a.Y + 0.000001
    If f.Y = a.Y Then a.Y = a.Y + 0.000001
 
    a1 = B.Y - a.Y
    b1 = a.X - B.X
    c1 = B.X * a.Y - a.X * B.Y
    a2 = f.Y - e.Y
    b2 = e.X - f.X
    c2 = f.X * e.Y - e.X * f.Y
 
    denom = a1 * b2 - a2 * b1
    If denom = 0 Then
        Exit Function
    End If
    
    ip.X = (b1 * c2 - b2 * c1) / denom
    ip.Y = (a2 * c1 - a1 * c2) / denom
 
    'If E.Y = A.Y Then Exit Function
    'If F.Y = A.Y Then Exit Function ' If the line goes through the end vertex, skip it, since we'll let it get caught by the start vertex
 
    '---------------------------------------------------
    'Do checks to see if intersection to endpoints
    'distance is longer than actual Segments.
    'Return null if it is with any.
    '---------------------------------------------------
    If as_seg Then
        If pointDistance(ip, B) > pointDistance(a, B) Then
            Exit Function
        End If
        If pointDistance(ip, a) > pointDistance(a, B) Then
            Exit Function
        End If
 
        If pointDistance(ip, f) > pointDistance(e, f) Then
            Exit Function
        End If
        If pointDistance(ip, e) > pointDistance(e, f) Then
            Exit Function
        End If
    End If
    
    lineIntersectLine = ip

End Function

Function pointDistance(a As pointD, B As pointD) As Double
    ' Return the distance between these two points
    pointDistance = Sqr((a.Y - B.Y) ^ 2 + (a.X - B.X) ^ 2)
End Function

Function newPoint(X As Double, Y As Double) As pointD
    newPoint.X = X
    newPoint.Y = Y
End Function


Function removeDupes(pointList() As pointD)
    ' remove duplicate points from an array of points
    'Dim pointList As New Scripting.Dictionary
    'Dim i As Long
    
    
    
    

End Function

Function calcPolyCenter(polyID As Long, ByRef X As Double, ByRef Y As Double)
    ' Calculate the centerpoint of the polygon
    
    Dim i As Long
    Dim cX As Double, cY As Double
    With pData(polyID)
        For i = 1 To UBound(.Points)
            cX = cX + .Points(i).X
            cY = cY + .Points(i).Y
        Next
        
        X = cX / UBound(.Points)
        Y = cY / UBound(.Points)
        
    End With

End Function

Function flipPolyStartEnd(polyID As Long)
    ' Flip the points around.
    Dim pTemp() As pointD
    Dim i As Long
    
    With pData(polyID)
        ' Store a copy of the array
        pTemp = .Points
        
        For i = 1 To UBound(pTemp)
            .Points(UBound(pTemp) - i + 1) = pTemp(i)
        Next
    End With

End Function

Function addFillPolies(polyFills() As POINTAPI, polyID As Long)

    Dim i As Long
    Dim j As Long
    Dim n As Long
    
    For i = 1 To UBound(pData)
        With pData(i)
            If .ContainedBy = polyID And .Fillable Then
                
                n = UBound(polyFills)
                ReDim Preserve polyFills(n + UBound(.Points))
                
                For j = 1 To UBound(.Points)
                    polyFills(j + n).X = (.Points(j).X + frmInterface.panX) * frmInterface.Zoom
                    polyFills(j + n).Y = (.Points(j).Y + frmInterface.panY) * frmInterface.Zoom
                Next
                
                addFillPolies polyFills, i
            End If
        End With
    Next

End Function
