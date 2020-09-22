Attribute VB_Name = "modUUcode"

Private Function Decode(szData As String) As String

    Dim szOut   As String
    Dim nChar   As Integer
    Dim I       As Integer
    Dim Pos     As Integer
    
    Do
        Pos = InStr(szData, Chr$(96))
        If Pos Then
            szData = Left$(szData, Pos - 1) & " " & Mid$(szData, Pos + 1)
        Else
            Exit Do
        End If
    Loop
    
    For I = 1 To Len(szData) Step 4
        szOut = szOut & Chr((Asc(Mid(szData, I, 1)) - 32) * 4 + (Asc(Mid(szData, I + 1, 1)) - 32) \ 16)
        szOut = szOut & Chr((Asc(Mid(szData, I + 1, 1)) Mod 16) * 16 + (Asc(Mid(szData, I + 2, 1)) - 32) \ 4)
        szOut = szOut & Chr((Asc(Mid(szData, I + 2, 1)) Mod 4) * 64 + Asc(Mid(szData, I + 3, 1)) - 32)
    Next I
    
    Decode = szOut
    
    Exit Function
    
End Function

Private Function Encode(szData As String) As String

    Dim szOut   As String
    Dim nChar   As Integer
    Dim I       As Integer
    Dim ThisChar As String
    Dim Here As Boolean
    
    '   pad to 3 byte multiple
    If Len(szData) Mod 3 <> 0 Then
        szData = szData & String(3 - Len(szData) Mod 3, Chr$(0))
    End If
    
    For I = 1 To Len(szData) Step 3
        
        ThisChar = Chr(Asc(Mid(szData, I, 1)) \ 4 + 32)
        If Asc(ThisChar) = 32 Then
            szOut = szOut & Chr$(96)
        Else
            szOut = szOut & ThisChar
        End If
        
        ThisChar = Chr((Asc(Mid(szData, I, 1)) Mod 4) * 16 + Asc(Mid(szData, I + 1, 1)) \ 16 + 32)
        If Asc(ThisChar) = 32 Then
            szOut = szOut & Chr$(96)
        Else
            szOut = szOut & ThisChar
        End If
        
        ThisChar = Chr((Asc(Mid(szData, I + 1, 1)) Mod 16) * 4 + Asc(Mid(szData, I + 2, 1)) \ 64 + 32)
        If Asc(ThisChar) = 32 Then
            szOut = szOut & Chr$(96)
        Else
            szOut = szOut & ThisChar
        End If
        
        ThisChar = Chr(Asc(Mid(szData, I + 2, 1)) Mod 64 + 32)
        If Asc(ThisChar) = 32 Then
            szOut = szOut & Chr$(96)
        Else
            szOut = szOut & ThisChar
        End If
    Next I
    
    Encode = szOut
    
End Function


Public Function UUDecode(szFileIn As String, szFileOut As String) As Integer

    Dim nFileIn     As Integer
    Dim nFileOut    As Integer
    Dim szData      As String
    Dim szOut       As String
    Dim lBytesIn    As Long
    Dim lFullLines  As Long
    
    On Error GoTo ERR_UUDecode
    
    '   open the ascii input file
    nFileIn = FreeFile
    Open szFileIn For Input As nFileIn
    
    '   find the header in the input file
    While LCase(Left(Trim(szData), 6)) <> "begin "
        Line Input #nFileIn, szData
        Wend
    
    '   open the binary output file
    nFileOut = FreeFile
    
    '   if an output file wasn't given, take the name from the input file
    If szFileOut = "" Then
        szData = Trim(szData)
        szData = Trim(Mid(szData, InStr(szData, " ")))
        szFileOut = Trim(Mid(szData, InStr(szData, " ")))
        End If
        
    Open szFileOut For Binary As nFileOut
    
    Do While Not EOF(nFileIn)
        
        '   get a 45 bytes chunk, encode it and put it in the output file
        Line Input #nFileIn, szData
        
        If Trim$(LCase$(szData)) = "end" Then
            Exit Do
        ElseIf Trim$(szData) <> "" Then
            '   decode the input line and put it into the output file
            szOut = Left(Decode(Mid(szData, 2, Len(szData) - 1)), Asc(Left(szData, 1)) - 32)
            Put #nFileOut, , szOut
        End If
        
    Loop
        
    '   close the files
    Close nFileIn
    Close nFileOut
    
    '   if we got this far, then it must have worked!
    '       return of 0 means there were no errors
    UUDecode = 0
    
    Exit Function

ERR_UUDecode:
    '   argghhh!, something went wrong, return the error code
    UUDecode = Err
    
    Close nFileIn
    Close nFileOut
    
    Exit Function
    
End Function

Public Function UUEncode(ByVal szFileIn As String, ByVal szFileOut As String, nAppend As Integer, Optional szFileInShort As String = "") As Long

    Dim nFileIn     As Integer
    Dim nFileOut    As Integer
    Dim nIndex      As Integer
    Dim szData      As String
    Dim szOutData   As String
    Dim lBytesIn    As Long
    Dim lFullLines  As Long
    
    UUEncode = 0
    On Error GoTo ERR_UUEncode
    
    '   open the binary input file
    nFileIn = FreeFile
    Open szFileIn For Binary As nFileIn
    lBytesIn = LOF(nFileIn)
    
    '   open the ascii output file
    nFileOut = FreeFile
    If nAppend Then
        Open szFileOut For Append As nFileOut
    Else
        Open szFileOut For Output As nFileOut
    End If
    
    '-- Return just the filename portion of the outfile
    For nIndex = Len(szFileOut) - 1 To 1 Step -1
        If Mid$(szFileOut, nIndex, 1) = "\" Then
            szFileOut = Mid$(szFileOut, nIndex + 1)
            Exit For
        End If
    Next
    
    '-- Return just the filename portion of the infile
    If (szFileInShort = "") Then
        For nIndex = Len(szFileIn) - 1 To 1 Step -1
            If Mid$(szFileIn, nIndex, 1) = "\" Then
                szFileInShort = Mid$(szFileIn, nIndex + 1)
                Exit For
            End If
        Next
    End If
    '-- Start with a CRLF if this is an append job.
    If nAppend Then
        Print #nFileOut, ""
    End If
    
    '-- Put the header in the output file
    Print #nFileOut, "begin 600 " & szFileInShort
    
    '-- Determine how many full lines we get, 45 bytes gets
    '   expanded to 60 bytes
    lFullLines = lBytesIn \ 45
    szData = Space(45)
    
    While lFullLines > 0
        
        '   get a 45 bytes chunk, encode it and put it in the output file
        Get nFileIn, , szData
        
        szOutData = "M" & Encode(szData)
        
        If szOutData = "M" Then Stop
        
        Print #nFileOut, szOutData
        
        '   another one "bytes" the dust
        lFullLines = lFullLines - 1
        
        Wend
        
    '   determine the leftover portion
    szData = Space(lBytesIn Mod 45)
    
    '   get the partial chunk of bytes that are left
    Get nFileIn, , szData
    
    '   put them in the output file
    Print #nFileOut, Chr(Len(szData) + 32) & Encode(szData)
    
    '   add on the file trailer
    Print #nFileOut, Chr$(96) & vbCrLf & "end"
    
    '   close the files
    Close nFileIn
    Close nFileOut
    
    '   if we got this far, then it must have worked!
    '       return of 0 means there were no errors
    UUEncode = lBytesIn \ 45 + 2
    
    Exit Function

ERR_UUEncode:
    '   argghhh!, something went wrong, return the error code
    UUEncode = Err
    
    Close nFileIn
    Close nFileOut
    
    Exit Function

End Function


Public Function UUDecodeString(ByVal InData As String, FileName As String) As String
    
    Dim ThisLine    As String
    Dim OutData     As String
    Dim lBytesIn    As Long
    Dim lFullLines  As Long
    
    On Error GoTo ERR_UUDecodeString
    
    '   find the header in the input file
    Do
        Pos = InStr(LCase(InData), vbLf & "begin ")
        If Pos Then
            InData = Mid$(InData, Pos + 7)
            Pos = InStr(InData, " ")
            If Pos Then
                InData = Mid$(InData, Pos + 1)
                Pos = InStr(InData, " ")
                If Pos Then
                    FileName = Left$(InData, Pos - 1)
                End If
            End If
        Else
            Exit Do
        End If
    Loop
    
    If Len(FileName) Then
        '-- Start at the next line
        Pos = InStr(InData, vbLf)
        If Pos Then
            InData = Mid$(InData, Pos + 1)
        End If
    
        Do While Len(InData)
            '   get a 45 bytes chunk, encode it and add it to the output string
            Pos = InStr(InData, vbLf)
            If Pos Then
                ThisLine = Left$(InData, Pos - 1)
                InData = Mid$(InData, Pos + 1)
                If Right$(ThisLine, 1) = vbCr Then
                    ThisLine = Left$(ThisLine, Len(ThisLine) - 1)
                End If
            End If
            
            If Trim$(LCase$(ThisLine)) = "end" Then
                Exit Do
            ElseIf Trim$(ThisLine) <> "" Then
                If Trim$(ThisLine) = Chr$(96) Then
                    Exit Do
                ElseIf Len(ThisLine) > 61 Then
                    ThisLine = Left$(ThisLine, 61)
                Else
                    ThisLine = Left$(ThisLine, Len(ThisLine) - 2)
                End If
                '   decode the input line and append it to the output string
                'szOut = Left(Decode(Mid(szData, 2, Len(szData) - 1)), Asc(Left(szData, 1)) - 32)
                OutData = OutData & Left$(Decode(Mid$(ThisLine, 2, Len(ThisLine) - 1)), Asc(Left$(ThisLine, 1)) - 32)
            End If
            DoEvents
        Loop
            
        UUDecodeString = OutData
    End If
    
    Exit Function

ERR_UUDecodeString:
    
    On Error Resume Next
    Exit Function

End Function
