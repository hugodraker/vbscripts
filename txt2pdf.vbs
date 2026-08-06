' ============================================================================
' TXT TO PDF CONVERTER - PURE VBSCRIPT (NO EXTERNAL LIBRARIES)
' License: Public Domain (Unlicensed)
' Usage: cscript txt2pdf.vbs input.txt [output.pdf]
' Font: Helvetica 12pt, automatic text reflow, US Letter pages
' Handles: CRLF and LF line endings
' ============================================================================

Option Explicit

Dim fso, WSHArgs, inputFile, outputFile
Dim totalPages, currentY, lineHeight, pageNum
Dim pageSizeW, pageSizeH, marginLeft, marginRight, topMargin, bottomMargin
Dim fontSize, maxLineWidth, charWidthEstimate
Dim tempStream
Dim totalObjects

Const FOR_READING = 1
Const ForWriting = 2
Const OverwriteExisting = True

Set fso = CreateObject("Scripting.FileSystemObject")
Set WSHArgs = WScript.Arguments

' Letter Page Dimensions (in points)
pageSizeW = 612     ' Width
pageSizeH = 792     ' Height 

marginLeft = 72     ' 1 inch left margin
marginRight = 72    ' 1 inch right margin
topMargin = 72      ' 1 inch top margin
bottomMargin = 72   ' 1 inch bottom margin

fontSize = 12       ' Point size
lineHeight = fontSize * 1.3
charWidthEstimate = 6.0  ' Approximate point width per character (Helvetica)
maxLineWidth = Int((pageSizeW - marginLeft - marginRight) / charWidthEstimate)

totalObjects = 0

' ============================================================================
' MAIN EXECUTION
' ============================================================================

If WSHArgs.Count < 1 Then
    WScript.Quit 1
End If

inputFile = WSHArgs(0)

If Not fso.FileExists(inputFile) Then
    WScript.Quit 1
End If

If WSHArgs.Count >= 2 Then
    outputFile = WSHArgs(1)
Else
    outputFile = fso.GetBaseName(inputFile) & ".pdf"
End If

Call ConvertTxtToPdf()

Set fso = Nothing

' ============================================================================
' CORE CONVERSION FUNCTION
' ============================================================================

Sub ConvertTxtToPdf()
    Dim rawText, lines, i, j
    Dim processedData(), dataCount
    Dim wordArray, tempLine
    
    ' Step 1: Read input file
    On Error Resume Next
    Set tempStream = CreateObject("ADODB.Stream")
    tempStream.Open
    tempStream.Charset = "utf-8"
    tempStream.LineSeparator = 10  ' Use LF internally
    tempStream.LoadFromFile inputFile
    rawText = tempStream.ReadText(-1)
    tempStream.Close
    Set tempStream = Nothing
    
    If Err.Number <> 0 Then
        On Error GoTo 0
        WScript.Quit 1
    End If
    On Error GoTo 0
    
    ' Normalize line endings
    rawText = Replace(rawText, vbCrLf, vbLf)
    rawText = Replace(rawText, vbCr, vbLf)
    
    lines = Split(rawText, vbLf)
    
    ' Step 2: Apply word-wrap reflow to each line
    ReDim processedData(UBound(lines) * 3 + UBound(lines) + 10)
    dataCount = 0
    
    For i = 0 To UBound(lines)
        If Trim(lines(i)) = "" Then
            processedData(dataCount) = ""
            dataCount = dataCount + 1
        Else
            tempLine = WrapText(lines(i))
            If InStr(tempLine, "|BREAK|") > 0 Then
                wordArray = Split(tempLine, "|BREAK|")
                For j = 0 To UBound(wordArray)
                    processedData(dataCount) = wordArray(j)
                    dataCount = dataCount + 1
                Next
            Else
                processedData(dataCount) = tempLine
                dataCount = dataCount + 1
            End If
        End If
    Next
    
    If dataCount > 0 Then
        ReDim Preserve processedData(dataCount - 1)
    Else
        ReDim processedData(0)
        processedData(0) = ""
    End If
    
    ' Step 3: Generate PDF structure
    Call BuildCompletePdfWithXref(processedData)
End Sub

' ============================================================================
' TEXT REFLOW/WORD-WRAP ENGINE
' ============================================================================

Function WrapText(lineText)
    Dim words, numWords, i, currentLine, nextWord
    Dim resultSegments, segCount, charsOnCurrentLine
    
    If Trim(lineText) = "" Then
        WrapText = ""
        Exit Function
    End If
    
    words = Split(lineText, " ")
    numWords = UBound(words) + 1
    
    currentLine = ""
    charsOnCurrentLine = 0
    segCount = 0
    ReDim resultSegments(numWords)
    
    For i = 0 To numWords - 1
        nextWord = words(i)
        
        If Len(currentLine) = 0 Then
            currentLine = nextWord
            charsOnCurrentLine = Len(nextWord)
        ElseIf charsOnCurrentLine + Len(nextWord) + 1 <= maxLineWidth Then
            currentLine = currentLine & " " & nextWord
            charsOnCurrentLine = charsOnCurrentLine + Len(nextWord) + 1
        Else
            resultSegments(segCount) = currentLine
            segCount = segCount + 1
            currentLine = nextWord
            charsOnCurrentLine = Len(nextWord)
        End If
    Next
    
    If Len(currentLine) > 0 Then
        resultSegments(segCount) = currentLine
        segCount = segCount + 1
    End If
    
    ReDim Preserve resultSegments(segCount - 1)
    WrapText = Join(resultSegments, "|BREAK|")
End Function

' ============================================================================
' COMPLETE PDF BUILDER WITH ACCURATE XREF TABLE
' ============================================================================

Sub BuildCompletePdfWithXref(textData)
    Dim fullPdf, objSegmentStart, i
    Dim childRefs, consolidatedPages
    Dim pageContentStr, contentLength
    Dim xrefOut(), binaryStream, xrefStartPos, paddedOffset
    Dim pageObjsStart
    
    ' Consolidate lines onto pages
    consolidatedPages = ConsolidateTextOntoPages(textData)
    totalPages = UBound(consolidatedPages) + 1
    
    ReDim xrefOut(totalPages * 2 + 10)
    
    fullPdf = "%PDF-1.4" & vbCrLf
    
    ' OBJ 1: Catalog Root  
    objSegmentStart = Len(fullPdf)
    fullPdf = fullPdf & "1 0 obj" & vbCrLf
    fullPdf = fullPdf & "<< /Type /Catalog /Pages 2 0 R >>" & vbCrLf
    fullPdf = fullPdf & "endobj" & vbCrLf
    xrefOut(1) = objSegmentStart
    
    ' OBJ 2: Pages Tree
    objSegmentStart = Len(fullPdf)
    childRefs = BuildPageChildReferences(totalPages)
    fullPdf = fullPdf & "2 0 obj" & vbCrLf
    fullPdf = fullPdf & "<< /Type /Pages /Kids [" & childRefs & "] /Count " & totalPages & " >>" & vbCrLf
    fullPdf = fullPdf & "endobj" & vbCrLf
    xrefOut(2) = objSegmentStart
    
    ' OBJ 3: Font Object
    objSegmentStart = Len(fullPdf)
    fullPdf = fullPdf & "3 0 obj" & vbCrLf
    fullPdf = fullPdf & "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>" & vbCrLf
    fullPdf = fullPdf & "endobj" & vbCrLf
    xrefOut(3) = objSegmentStart
    
    ' OBJ 4+: Pages and Content Streams
    pageObjsStart = 3
    
    For i = 0 To UBound(consolidatedPages)
        ' Page Dictionary
        objSegmentStart = Len(fullPdf)
        fullPdf = fullPdf & CStr(pageObjsStart + (i * 2) + 1) & " 0 obj" & vbCrLf
        fullPdf = fullPdf & "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " & pageSizeW & " " & pageSizeH & "] /Resources <<" & vbCrLf
        fullPdf = fullPdf & "/Font << /F1 3 0 R >> >>" & vbCrLf
        fullPdf = fullPdf & "/Contents " & CStr(pageObjsStart + (i * 2) + 2) & " 0 R >>" & vbCrLf
        fullPdf = fullPdf & "endobj" & vbCrLf
        xrefOut(pageObjsStart + (i * 2) + 1) = objSegmentStart
        
        ' Content Stream
        pageContentStr = GenerateTextContentForPage(consolidatedPages(i))
        contentLength = Len(pageContentStr)
        
        objSegmentStart = Len(fullPdf)
        fullPdf = fullPdf & CStr(pageObjsStart + (i * 2) + 2) & " 0 obj" & vbCrLf
        fullPdf = fullPdf & "<< /Length " & contentLength & " >>" & vbCrLf
        fullPdf = fullPdf & "stream" & vbCrLf
        fullPdf = fullPdf & pageContentStr
        fullPdf = fullPdf & "endstream" & vbCrLf
        fullPdf = fullPdf & "endobj" & vbCrLf
        xrefOut(pageObjsStart + (i * 2) + 2) = objSegmentStart
    Next
    
    totalObjects = 3 + (totalPages * 2)
    
    ' Build XREF table
    xrefStartPos = Len(fullPdf)
    fullPdf = fullPdf & "xref" & vbCrLf
    fullPdf = fullPdf & "0 " & CStr(totalObjects + 1) & vbCrLf
    ' Entries MUST be exactly 20 bytes long (18 chars + 2-byte CRLF)
    fullPdf = fullPdf & "0000000000 65535 f" & vbCrLf
    
    For i = 1 To totalObjects
        paddedOffset = Right("0000000000" & CStr(xrefOut(i)), 10)
        fullPdf = fullPdf & paddedOffset & " 00000 n" & vbCrLf
    Next
    
    ' Trailer and EOF
    fullPdf = fullPdf & "trailer" & vbCrLf
    fullPdf = fullPdf & "<< /Size " & CStr(totalObjects + 1) & " /Root 1 0 R >>" & vbCrLf
    fullPdf = fullPdf & "startxref" & vbCrLf
    fullPdf = fullPdf & CStr(xrefStartPos) & vbCrLf
    fullPdf = fullPdf & "%%EOF" & vbCrLf
    
    ' Write output string
    Set binaryStream = CreateObject("ADODB.Stream")
    binaryStream.Open
    binaryStream.Type = 2         ' adTypeText
    binaryStream.Charset = "iso-8859-1"
    binaryStream.WriteText fullPdf
    binaryStream.SaveToFile outputFile, 2
    binaryStream.Close
    Set binaryStream = Nothing
End Sub

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

Function EscapePdfString(str)
    Dim i, c, code, res
    res = ""
    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        code = AscW(c)
        If code < 0 Then code = code + 65536
        
        Select Case c
            Case "\" : res = res & "\\"
            Case "(" : res = res & "\("
            Case ")" : res = res & "\)"
            Case Else
                If code >= 32 And code <= 126 Then
                    res = res & c
                ElseIf code >= 128 And code <= 255 Then
                    res = res & "\" & Right("00" & Oct(code), 3)
                ElseIf code = 9 Then
                    res = res & "    "
                Else
                    res = res & "?"
                End If
        End Select
    Next
    EscapePdfString = res
End Function

Function BuildPageChildReferences(pageCount)
    Dim refs, i, pageRefStart
    refs = ""
    pageRefStart = 4
    
    For i = 0 To pageCount - 1
        If Len(refs) > 0 Then refs = refs & " "
        refs = refs & CStr(pageRefStart + (i * 2)) & " 0 R"
    Next
    
    BuildPageChildReferences = refs
End Function

Function ConsolidateTextOntoPages(txtArray)
    Dim linesPerPage, pageChunks(), pageIndex, lineIndex
    Dim currentChunk, i
    
    linesPerPage = Int((pageSizeH - topMargin - bottomMargin) / lineHeight)
    ReDim pageChunks((UBound(txtArray) \ linesPerPage) + 1)
    
    pageIndex = 0
    lineIndex = 0
    currentChunk = ""
    
    For i = 0 To UBound(txtArray)
        If Len(currentChunk) > 0 Then
            currentChunk = currentChunk & "|NEWLINE|" & txtArray(i)
        Else
            currentChunk = txtArray(i)
        End If
        
        lineIndex = lineIndex + 1
        
        If lineIndex >= linesPerPage And i < UBound(txtArray) Then
            pageChunks(pageIndex) = currentChunk
            pageIndex = pageIndex + 1
            currentChunk = ""
            lineIndex = 0
        End If
    Next
    
    If Len(Trim(currentChunk)) > 0 Or pageIndex = 0 Then
        pageChunks(pageIndex) = currentChunk
        pageIndex = pageIndex + 1
    End If
    
    ReDim Preserve pageChunks(pageIndex - 1)
    ConsolidateTextOntoPages = pageChunks
End Function

Function GenerateTextContentForPage(pageText)
    Dim lines, i, yPos, xPos, startY, encodedLine
    Dim contentBuilder
    
    lines = Split(pageText, "|NEWLINE|")
    startY = pageSizeH - topMargin
    xPos = marginLeft
    
    contentBuilder = "BT" & vbCrLf
    contentBuilder = contentBuilder & "/F1 " & fontSize & " Tf" & vbCrLf
    contentBuilder = contentBuilder & "0.0 0.0 0.0 rg" & vbCrLf
    
    For i = 0 To UBound(lines)
        yPos = startY - (i * lineHeight)
        
        If yPos < bottomMargin Then
            Exit For
        End If
        
        encodedLine = EscapePdfString(lines(i))
        contentBuilder = contentBuilder & "1 0 0 1 " & CStr(xPos) & " " & CStr(yPos) & " Tm" & vbCrLf
        contentBuilder = contentBuilder & "(" & encodedLine & ") Tj" & vbCrLf
    Next
    
    contentBuilder = contentBuilder & "ET" & vbCrLf
    
    GenerateTextContentForPage = contentBuilder
End Function