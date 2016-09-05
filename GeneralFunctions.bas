Attribute VB_Name = "GeneralFunctions"
'General functions that could be used in any project.
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

Public Declare Function MakeSureDirectoryPathExists Lib "IMAGEHLP.DLL" (ByVal DirPath As String) As Long

Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type typGenSizeInfo
    numFiles As Long
    totalSize As Long
End Type

Public debug_noLog As Boolean

Public Type SHFILEOPSTRUCT
        hWnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAnyOperationsAborted As Long
        hNameMappings As Long
        lpszProgressTitle As String
End Type

Public Declare Function SHFileOperation Lib "shell32.dll" _
Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Public Const FO_DELETE = &H3
Public Const FO_MOVE = &H1
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_CONFIRMMOUSE = &H2
Public Const FOF_FILESONLY = &H80                  '  on *.*, do only files
Public Const FOF_MULTIDESTFILES = &H1
Public Const FOF_NOCONFIRMATION = &H10             '  Don't prompt the user.
Public Const FOF_NOCONFIRMMKDIR = &H200            '  don't confirm making any needed dirs
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_SILENT = &H4                      '  don't create progress/report
Public Const FOF_SIMPLEPROGRESS = &H100            '  means don't show names of files
Public Const FOF_WANTMAPPINGHANDLE = &H20          '  Fill in SHFILEOPSTRUCT.hNameMappings
Public Const FOF_NOERRORUI = &H400


' FILE AND DISK

' Get file extension.
Function getFileExten(inName As String) As String
    Dim e As Long
    e = InStrRev(inName, ".")
    If e > 0 Then getFileExten = LCase(Right(inName, Len(inName) - e))
End Function


Function getFileNameFromPath(inPth As String) As String
    ' Return the file name from the path
    
    Dim e As Long
    e = InStrRev(inPth, "\")
    If e = 0 Then e = InStrRev(inPth, "/")
    
    If e > 0 Then
        getFileNameFromPath = Right(inPth, Len(inPth) - e)
    Else
        getFileNameFromPath = inPth
    End If

End Function

Function getFileNameNoExten(inName As String) As String
    Dim e As Long
    e = InStrRev(inName, ".")
    If e > 0 Then getFileNameNoExten = Left(inName, e - 1)
End Function

Function getFolderNameFromPath(inPth As String) As String
    ' Return the path from the path and file
    
    Dim e As Long
    e = InStrRev(inPth, "\")
    If e = 0 Then e = InStrRev(inPth, "/")
    
    If e > 0 Then
        getFolderNameFromPath = Left(inPth, e - 1)
    End If
End Function

' DOES NOT RECURSE
' TODO: Make it take a file mask.
Function getSizeOfFolder(pth As String) As typGenSizeInfo
    Dim a As String
    
    a = myDir(pth & "\*.*")
    Do Until a = ""
        Select Case LCase(getFileExten(a))
            Case "jpg", "jpeg", "png", "gif", "wav"
                getSizeOfFolder.totalSize = getSizeOfFolder.totalSize + FileLen(pth & "\" & a)
                getSizeOfFolder.numFiles = getSizeOfFolder.numFiles + 1
        End Select
        a = Dir
    Loop
End Function




' GENERAL
Public Function Swap(ByRef in1, ByRef in2)
    Dim B
    B = in1
    in1 = in2
    in2 = B
End Function

' COLOR

Function makeVBColor(inHex As String) As Long
    ' Turns FF0033 into the VB color.
    Dim R As Long, g As Long, B As Long
    
    If Len(inHex) < 6 Then
        inHex = Replace(Space(6 - Len(inHex)), " ", "0") & inHex
    End If
    
    If Len(inHex) = 6 Then
        R = Val("&H" & Mid(inHex, 1, 2))
        g = Val("&H" & Mid(inHex, 3, 2))
        B = Val("&H" & Mid(inHex, 5, 2))
        
        makeVBColor = (B * 256 * 256) + (g * 256) + R
    Else
        makeVBColor = 0
    End If
    

End Function

Function myTrim(inSt As String, Optional trimChar As String) As String
    ' Trim with a specified character.
    Dim i As Long
    Dim j As Long
    
    If trimChar = "" Then trimChar = " "
    
    For i = 1 To Len(inSt)
        If Mid(inSt, i, 1) <> trimChar Then Exit For
    Next
    
    For j = Len(inSt) To i Step -1
        If Mid(inSt, j, 1) <> trimChar Then Exit For
    Next
    
    myTrim = Mid(inSt, i, j - i + 1)

End Function

Function Escape(inTxt As String) As String
    
    'Escape = vbcurl_string_escape(inTxt, Len(inTxt))
    
    'Exit Function
       
    
    ' Escape the text.
    Dim i As Long
    Dim outText As String
    Dim B As String
    
    Escape = inTxt
    
    Escape = Replace(Escape, "%", "%25")
    For i = 1 To 255
        If i = 37 Then
            ' skip %
        ElseIf i >= 65 And i <= 90 Then
            ' A-Z
        ElseIf i >= 97 And i <= 122 Then
            ' a-z
        ElseIf i >= 48 And i <= 57 Then
            ' 0-9
        Else
            Escape = Replace(Escape, Chr(i), "%" & fixZeros(Hex(i)))
        End If
    Next
    
    
    
    
    ' Old code
    'For i = 1 To Len(inTxt)
    '    b = Mid(inTxt, i, 1)
    '    If (b >= "A" And b <= "Z") Or (b >= "a" And b <= "z") Or (b >= "0" And b <= "9") Then
    '        outText = outText & b
    '    Else
    '        outText = outText & "%" & fixZeros(Hex(Asc(b)))
    '    End If
    'Next
    
    'Escape = outText

End Function

Function Unescape(ByVal inSt As String) As String
    
    ' Unescape %20 and stuff like that.
    Dim e As Long
    Dim nTwo As String
    Dim nVal As Long
    On Error Resume Next
    
    Do
        e = InStr(e + 1, inSt, "%")
        If e > 0 Then
            
            ' get next two characters.
            If e + 2 <= Len(inSt) Then
                nTwo = Mid(inSt, e + 1, 2)
                
                ' Convert hex to number
                nVal = 0
                nVal = Val("&H" & nTwo)
                
                If nVal > 0 Then
                    inSt = Left(inSt, e - 1) & Chr(nVal) & Right(inSt, Len(inSt) - e - 2)
                End If
            End If
        End If
    Loop Until e = 0
    
    Unescape = inSt
    
End Function

Function fixZeros(inSt As String) As String
    ' Adds a 0 to the front if needed.
    fixZeros = inSt
    If Len(fixZeros) = 1 Then fixZeros = "0" & fixZeros
End Function


Function Max(n1, n2)
    If n1 > n2 Then
        Max = n1
    Else
        Max = n2
    End If
End Function

Function Min(n1, n2)
    If n1 < n2 Then
        Min = n1
    Else
        Min = n2
    End If
End Function

Function isIn(checkFor, ParamArray checkIn()) As Boolean
    ' See if the value is in one of the items
    Dim i As Long
    For i = 0 To UBound(checkIn)
        If checkFor = checkIn(i) Then isIn = True: Exit Function
    Next
    
    
End Function

Function isInArray(checkFor, checkIn) As Boolean
    ' See if the value is in one of the items
    Dim i As Long
    For i = 0 To UBound(checkIn)
        If checkFor = checkIn(i) Then isInArray = True: Exit Function
    Next
End Function

Function separateURL(inURL As String, ByRef Host As String, ByRef path As String, Optional ByRef port As Long, Optional ByRef UserN As String, Optional ByRef PassW As String, Optional protocol As String) As Boolean
    ' Turn the URL:
    ' http://www.site.com/path/to/file.pat?text
    ' ftp://user:pass@www.site.com:port/path/to/file
    
    ' into host and path
    Dim URL As String
    Dim userpass As String
    URL = inURL
    
    Dim e As Long
    Dim f As Long
    
    e = InStr(URL, "://")
    If e > 0 Then
        
        ' remove http://
        protocol = Left(URL, e + 2)
        URL = Right(URL, Len(URL) - e - 2)
        
        
        e = InStr(URL, "/")
        If e > 0 Then
            Host = Left(URL, e - 1)
            path = Right(URL, Len(URL) - e + 1)
        Else
            Host = URL
        End If
        
        e = InStrRev(Host, ":")
        f = InStrRev(Host, "@")
        If e > 0 And e > f Then
            port = Val(Right(Host, Len(Host) - e))
            Host = Left(Host, e - 1)
        End If
        
        e = InStr(Host, "@")
        If e > 0 Then
            userpass = Left(Host, e - 1)
            Host = Right(Host, Len(Host) - e)
            
            e = InStr(userpass, ":")
            If e > 0 Then
                ' user and pass
                UserN = Left(userpass, e - 1)
                PassW = Right(userpass, Len(userpass) - e)
            Else
                UserN = userpass
            End If
        End If
        
        separateURL = True
    End If
End Function

Function setComboBoxToTextItem(checkComboBox As ComboBox, matchString As String) As Boolean
    ' Set a combobox to the item
    
    Dim i As Long
    For i = 0 To checkComboBox.ListCount - 1
        If matchString = checkComboBox.List(i) Then checkComboBox.ListIndex = i: setComboBoxToTextItem = True: Exit Function
    Next
    checkComboBox.ListIndex = -1
End Function

Function myDir(Optional Pathname As String, Optional Attributes As VbFileAttribute = vbNormal) As String

    ' Dir to avoid errors
    On Error Resume Next
    If Pathname = "" Then
        myDir = ""
    Else
        myDir = Dir(Pathname, Attributes)
    End If

End Function


Public Function addToLog(ParamArray entries())
    
    If debug_noLog Then Exit Function
    
    On Error Resume Next
    
    Dim t As String
    Dim x
    t = "[" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "." & Format(Int(getDecimal(Timer) * 1000), "000") & "] "
    
    For Each x In entries
        t = t & x & Chr(9)
    Next
    
    Dim f As Long
    f = FreeFile
    
    'Open debugLogPath For Append As f
    '    Print #f, t
    'Close f
    
    Debug.Print t
    
    
    


End Function

Function secsToTime(inSecs As Double) As String
    ' Convert 3232.587329 to a nice number like 4:56:22.31
    Dim s As Double
    s = inSecs
    Dim h As Long, M As Long
    
    h = Int(s / 3600)
    s = s - (h * 3600)
    
    M = Int(s / 60)
    s = s - (M * 60)
    
    secsToTime = h & ":" & Format(M, "00") & ":" & Format(Round(s, 2), "00.00")
End Function

Function secsToTime2(inSecs As Double) As String
    ' Convert 3232.587329 to a nice number like 4:56:22.31
    Dim s As Double
    s = inSecs
    Dim h As Long, M As Long
    
    h = Int(s / 3600)
    s = s - (h * 3600)
    
    M = Int(s / 60)
    s = s - (M * 60)
    
    secsToTime2 = h & ":" & Format(M, "00") & ":" & Format(Round(s, 2), "00")
End Function
Function secsToTime3(inSecs As Double) As String
    ' Convert 3232.587329 to a nice number like 4:56:22.31
    Dim s As Double
    s = inSecs
    Dim h As Long, M As Long
    
    h = Int(s / 3600)
    s = s - (h * 3600)
    
    M = Int(s / 60)
    s = s - (M * 60)
    
    secsToTime3 = h & " hours, " & M & " minutes"
End Function

Function Filerize(inSt As String) As String

    ' Turn any string of text into a valid filename.
    Dim i As Long
    Dim B As Long
    For i = 1 To Len(inSt)
        B = Asc(Mid(inSt, i, 1))
        If (B >= 65 And B <= 90) _
            Or (B >= 48 And B <= 57) _
            Or (B >= 97 And B <= 122) _
            Or (B >= 35 And B <= 41) _
            Or (B >= 44 And B <= 46) _
            Or B = 32 Or B = 95 Then
            
                Filerize = Filerize & Chr(B)
        Else
            Filerize = Filerize & "_"
        End If
    Next

End Function


Public Function GCD_Of(ByVal First_Int As Double, ByVal Second_Int As Double, ByRef Numerator As Long, ByRef Denominator As Long) As Boolean
    Dim Q
    Dim R
    Dim x
    Dim y
    Dim i As Long
    
    On Error Resume Next
    
         
    Q = CDec(1) ' Initialize quotient as DECIMAL variable type
    R = Q       ' Initialize remainder
        

      
    ' Convert input arguments into DECIMAL variable type
    If Int(First_Int) <> First_Int Or Int(Second_Int) <> Second_Int Then First_Int = First_Int * 100: Second_Int = Second_Int * 100
    
    First_Int = CDec(First_Int)
    Second_Int = CDec(Second_Int)
       
    ' Read the input argument values
    x = First_Int
    y = Second_Int
       
    ' Make sure both arguments are absolute values
    x = Abs(x)
    y = Abs(y)
       
    ' Report error if either argument is zero
    If x = 0 Or y = 0 Then Exit Function
      
    ' Swap argument values, if necessary, so that X > Y
    If x < y Then Q = x: x = y: y = Q
      
    ' Perform Euclid's algorithm to find GCD of X and Y
    While R <> 0
        Q = x / y
        'i = InStr(Q, ".")
        'If i > 0 Then Q = Left(Q, i - 1)
        
        ' Truncate decimal.
        
        Q = Int(Q)
        
        
        R = x - Q * y
        x = y
        y = R
    Wend
      
    ' Return the result
    GCD_Of = True
    'Debug.Print "result: ", X
    
    Numerator = First_Int / x
    Denominator = Second_Int / x
    
    
      
End Function


Public Sub MakeTopMost(hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Function ceiling(Number As Double) As Long
    ceiling = -Int(-Number)
End Function

Function makeHTMLCodes(ByVal ST As String)

    ST = Replace(ST, "&", "&amp;")

    ST = Replace(ST, Chr(34), "&quot;")
    ST = Replace(ST, "<", "&lt;")
    ST = Replace(ST, ">", "&gt;")
    ST = Replace(ST, "'", "&apos;")
    
      
    makeHTMLCodes = ST
End Function


Function toJSONString(inSt As String) As String

    toJSONString = inSt
    toJSONString = Replace(toJSONString, "\", "\\")
    toJSONString = Replace(toJSONString, Chr(8), "\b")
    toJSONString = Replace(toJSONString, Chr(34), "\""")
    toJSONString = Replace(toJSONString, Chr(12), "\f")
    toJSONString = Replace(toJSONString, Chr(10), "\n")
    toJSONString = Replace(toJSONString, Chr(13), "\r")
    toJSONString = Replace(toJSONString, Chr(9), "\t")

End Function


Function isKeyDown(KeyCode As Long) As Boolean

    isKeyDown = Abs(GetKeyState(KeyCode)) > 1

End Function

Function MyIsNumeric(Expression) As Boolean
    ' Deals with bugs in the real IsNumeric
    
    If VarType(Expression) = vbString Then
        If safeRight(Trim(Expression), 1) = "+" Then Exit Function
        If safeRight(Trim(Expression), 1) = "-" Then Exit Function
    End If
        
    MyIsNumeric = IsNumeric(Expression)
    
    
    
End Function

Function safeLeft(inSt As String, numChr As Long) As String
    If Len(inSt) >= numChr Then
        safeLeft = Left(inSt, numChr)
    Else
        safeLeft = inSt
    End If
End Function


Function safeRight(inSt As String, numChr As Long) As String
    If Len(inSt) >= numChr Then
        safeRight = Right(inSt, numChr)
    Else
        safeRight = inSt
    End If
End Function

Function fileIntoMemory(tPath As String) As String
    ' Quick load a file into memory
    
    On Error Resume Next
    
    Dim f As Long
    Dim g As String
    
    g = Space(FileLen(tPath))
    f = FreeFile
    Open tPath For Binary As f
        Get #f, , g
    Close f
    
    fileIntoMemory = g

End Function

Function getDecimal(inNum As Double) As Double
    ' Return the number after the decimal point
    getDecimal = inNum - Int(inNum)
End Function

Function GetFile(inPath As String) As String
    ' Load this file into memory.
    
    
    Dim f As Long, g As String
   On Error GoTo GetFile_Error

    f = FreeFile
    g = Space(FileLen(inPath))
    
    Open inPath For Binary As f
        Get #f, , g
    Close f
    
    GetFile = g

   On Error GoTo 0
   Exit Function

GetFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetFile of Module GeneralFunctions", vbCritical, "LIVELAYOUT ERROR"
    addToLog "[ERROR]", "In GetFile of Module GeneralFunctions", Err.Number, Err.Description

End Function


Function MoveFile(origPath As String, newPath As String)
    
    ' Move the files
    Dim WinType_SFO As SHFILEOPSTRUCT
    Dim lRet As Long
    
    With WinType_SFO
        .wFunc = FO_MOVE
        .pFrom = origPath & Chr(0)
        .pTo = newPath & Chr(0)
        .fFlags = FOF_MULTIDESTFILES Or FOF_NOCONFIRMMKDIR Or FOF_NOERRORUI Or FOF_SILENT
    End With
    
    lRet = SHFileOperation(WinType_SFO)

End Function

Function addCredentialsToPath(inPath As String, sUser As String, sPass As String) As String
    ' Add these credentials to the path.
    Dim aHost As String, aPath As String, aPort As Long, aProtocol As String
    
    separateURL inPath, aHost, aPath, aPort, , , aProtocol
    
    addCredentialsToPath = aProtocol & sUser & ":" & sPass & "@" & aHost & IIf(aPort > 0, ":" & aPort, "") & aPath
    
End Function


Public Function HandleError(ErrLine, ErrLocation, ErrNum, ErrDesc)
    addToLog "[ERROR]", ErrLocation, ErrNum, ErrDesc, "Line " & ErrLine
    MsgBox "Error " & ErrNum & " (" & ErrDesc & ") " & ErrLocation & " (Line " & ErrLine & ")", vbCritical, App.ProductName & " ERROR"
End Function
