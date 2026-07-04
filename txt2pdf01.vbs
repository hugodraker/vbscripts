' ============================================================================
' TXT TO PDF CONVERTER - PURE VBSCRIPT (NO EXTERNAL LIBRARIES)
' Author: Lumo Assistant  
' Usage: cscript txt2pdf.vbs input.txt [output.pdf]
' Font: Helvetica 12pt, automatic text reflow, A4 pages
' Handles: CRLF and LF line endings
' ============================================================================

Option Explicit

Dim fso, WSHArgs, inputFile, outputFile
Dim pdfContent, byteOffset, objOffsets()
Dim totalPages, currentY, lineHeight, pageNum
Dim pageSizeW, pageSizeH, marginLeft, marginRight, topMargin, bottomMargin
Dim fontSize, maxLineWidth, charWidthEstimate
Dim streamObj, tempStream

Const FOR_READING = 1
Const ForWriting = 2
Const OverwriteExisting = True

Set fso = CreateObject("Scripting.FileSystemObject")
Set WSHArgs = WScript.Arguments

' Page layout constants (A4 size in points)
pageSizeW = 595     ' Width
pageSizeH = 842     ' Height  
marginLeft = 72     ' 1 inch left margin
marginRight = 72    ' 1 inch right margin
topMargin = 72      ' 1 inch top margin
bottomMargin = 72   ' 1 inch bottom margin

fontSize = 12       ' Point size
lineHeight = fontSize * 1.3
charWidthEstimate = 6.0  ' Approximate point width per character (Helvetica)
maxLineWidth = Int((pageSizeW - marginLeft - marginRight) / charWidthEstimate)

' Global tracking arrays
Dim totalObjects : totalObjects = 0
ReDim objOffsets(200) ' Dynamic resizing handled below

' ============================================================================
' MAIN EXECUTION
' ============================================================================

If WSHArgs.Count < 1 Then
    ShowUsage
    WScript.Quit 1
End If

inputFile = WSHArgs(0)

If Not fso.FileExists(inputFile) Then
    WScript.Echo "[ERROR] File not found: " & inputFile
    WScript.Quit 1
End If

outputFile = IIf(WSHArgs.Count >= 2, WSHArgs(1), _
                 fso.GetBaseName(inputFile) & ".pdf")

WScript.Echo "[INFO] Converting text to PDF..."
WScript.Echo "[INPUT ] " & inputFile
WScript.Echo "[OUTPUT] " & outputFile
WScript.Echo "[FONT]   Helvetica " & fontSize & "pt"
WScript.Echo "[SIZE]   A4 (" & pageSizeW & "x" & pageSizeH & " pts)"

Call ConvertTxtToPdf()

WScript.Echo "[SUCCESS] Conversion completed!"
WScript.Echo "[PAGES]  " & totalPages & " page(s) created"

Set fso = Nothing

' ============================================================================
' CORE CONVERSION FUNCTION
' ============================================================================

Sub ConvertTxtToPdf()
    Dim rawText, lines, wrappedLines, i, j
    Dim processedData(), dataCount
    Dim pageContentStream, wordArray, tempLine
    
    ' Step 1: Read input file
    On Error Resume Next
    Set tempStream = CreateObject("ADODB.Stream")
    tempStream.Open
    tempStream.Charset = "utf-8"
    tempStream.LineSeparator = 10  ' Use LF internally
    tempStream.LoadFromFile inputFile
    rawText = tempStream.ReadText(-1) ' -1 = read until end
    tempStream.Close
    Set tempStream = Nothing
    On Error GoTo 0
    
    If Err.Number <> 0 Then
        WScript.Echo "[ERROR] Cannot read file: " & Err.Description
        WScript.Quit 1
    End If
    
    ' Normalize line endings (CRLF, CR, or LF)
    rawText = Replace(rawText, vbCrLf, vbLf)
    rawText = Replace(rawText, vbCr, vbLf)
    
    ' Split into initial lines
    lines = Split(rawText, vbLf)
    
    ' Step 2: Apply word-wrap reflow to each line
    ReDim processedData(UBound(lines) * 3 + UBound(lines))
    dataCount = 0
    
    For i = 0 To UBound(lines)
        If Trim(lines(i)) = "" Then
            ' Empty line becomes a single blank text line
            processedData(dataCount) = ""
            dataCount = dataCount + 1
        Else
            ' Apply word-wrapping algorithm
            tempLine = WrapText(lines(i))
            If InStr(tempLine, "|BREAK|") > 0 Then
                ' Multiple wrapped segments
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
    
    ' Resize to actual count
    ReDim Preserve processedData(dataCount - 1)
    
    ' Step 3: Calculate required pages (text flows vertically)
    totalPages = CalculateRequiredPages(processedData)
    WScript.Echo "[LINES]  " & dataCount & " logical text segments"
    
    ' Step 4: Generate PDF structure in phases
    Call BuildTwoPassPDFStructure(processedData)
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
    
    ' Split on spaces for word boundary detection
    words = Split(lineText, " ")
    numWords = UBound(words) + 1
    
    currentLine = ""
    charsOnCurrentLine = 0
    segCount = 0
    ReDim resultSegments(numWords)  ' Safe overestimate
    
    For i = 0 To numWords - 1
        nextWord = words(i)
        
        If Len(currentLine) = 0 Then
            ' First word on new line
            currentLine = nextWord
            charsOnCurrentLine = Len(nextWord)
        ElseIf charsOnCurrentLine + Len(nextWord) + 1 <= maxLineWidth Then
            ' Fits on same line (adds space separator)
            currentLine = currentLine & " " & nextWord
            charsOnCurrentLine = charsOnCurrentLine + Len(nextWord) + 1
        Else
            ' Exceeds line length - flush current and start new
            resultSegments(segCount) = currentLine
            segCount = segCount + 1
            currentLine = nextWord
            charsOnCurrentLine = Len(nextWord)
        End If
    Next
    
    ' Flush remaining accumulated text
    If Len(currentLine) > 0 Then
        resultSegments(segCount) = currentLine
        segCount = segCount + 1
    End If
    
    ' Resize and join output
    ReDim Preserve resultSegments(segCount - 1)
    WrapText = Join(resultSegments, "|BREAK|")
End Function

' ============================================================================
' PAGE CALCULATION
' ============================================================================

Function CalculateRequiredPages(dataArr)
    Dim i, availableSpacePerPage, spaceUsed
    Dim calcPages, yPosition
    
    availableSpacePerPage = pageSizeH - topMargin - bottomMargin
    spaceUsed = 0
    calcPages = 1
    
    For i = 0 To UBound(dataArr)
        If Len(Trim(dataArr(i))) > 0 Then
            spaceUsed = spaceUsed + lineHeight
            
            ' Check if exceeds current page bounds
            If spaceUsed > availableSpacePerPage Then
                calcPages = calcPages + 1
                spaceUsed = lineHeight  ' Restart counting on new page
            End If
        End If
    Next
    
    CalculateRequiredPages = calcPages
End Function

' ============================================================================
' TWO-PASS PDF GENERATION
' Phase 1: Build all objects, track byte offsets
' Phase 2: Write cross-reference table with accurate positions
' ============================================================================

Sub BuildTwoPassPDFStructure(textData)
    Dim pass1Buffer, pass2Buffer, i
    Dim startPos, currentPos, objNum, fontObjNum, catalogObjNum
    Dim pagesObjNum, pageObjNum, contentsObjNum
    Dim objectCounter, xrefEntries()
    
    ' Initialize global counter
    objectCounter = 0
   pageNum = 1
    currentY = pageSizeH - topMargin
    
    ' Pass 1: Build complete PDF content with placeholders
    ' We'll construct it incrementally and measure byte positions
    
    pass1Buffer = "%PDF-1.4" & vbCrLf
    AddBytePosition pass1Buffer  ' Record position of first object (will be object 0 entry)
    
    ' Object 1: Catalog root
    AddObjectDefinition pass1Buffer, "/Type /Catalog /Pages 2 0 R"
    
    ' Object 2: Pages tree (placeholder for kids array)
    Dim kidsPlaceholder : kidsPlaceholder = ""
    AddObjectDefinition pass1Buffer, "/Type /Pages /Kids [" & kidsPlaceholder & "] /Count " & totalPages
    
    ' Objects will continue: 3=catalog-page ref, 4=media box, 5+=font
    ' We need dynamic allocation as pages are added
    
    ' Calculate how many objects we'll need:
    ' Base objects = 5 (catalog, pages tree, font)
    ' Per page = 3 (page dict, media box already done, contents stream)
    ' But pages share parent tree and media box...
    ' Simplified: 5 base + (pages × 2) for page dict + contents
    objectCounter = 5 + (totalPages * 2)
    
    ReDim xrefEntries(totalPages * 3 + 10)
    
    ' Reset to build phase 2 properly
    Call BuildCompletePdfWithXref(textData, xrefEntries)
End Sub

Sub AddBytePosition(ByRef buf)
    Dim currentLen
    currentLen = Len(buf)
    ReDim Preserve objOffsets(totalObjects)
    objOffsets(totalObjects) = currentLen
    totalObjects = totalObjects + 1
End Sub

Sub AddObjectDefinition(ByRef buf, content)
    Dim objNum
    objNum = totalObjects
    ReDim Preserve objOffsets(totalObjects)
    objOffsets(totalObjects) = Len(buf)
    totalObjects = totalObjects + 1
    
    buf = buf & CStr(objNum) & " 0 obj" & vbCrLf
    buf = buf & "<< " & content & " >>" & vbCrLf
    buf = buf & "endobj" & vbCrLf
End Sub

' ============================================================================
' FINAL COMPLETE PDF BUILDER WITH ACCURATE XREF TABLE
' ============================================================================

Sub BuildCompletePdfWithXref(textData, xrefOut)
    Dim fullPdf, segmentStart, posTracker, i, j
    Dim currentObjNum, pageDict, contentStream
    Dim fontString, catalogString, pagesString
    Dim pageString, childrenStrings, childRefs
    Dim actualBytes, binaryStream
    
    ' Initialize tracking variables
    totalObjects = 0
    ReDim xrefOut(1000)  ' Preallocate for safety
    pageNum = 1
    currentPageY = pageSizeH - topMargin
    
    ' PHASE 1: Construct entire PDF as text string first
    fullPdf = "%PDF-1.4" & vbCrLf
    
    ' Store starting positions before each object
    SegmentTracking(0) = Len(fullPdf)  ' Position 0 (free object marker)
    
    ' OBJ 1: Catalog Root  
    objSegmentStart = Len(fullPdf)
    fullPdf = fullPdf & "1 0 obj" & vbCrLf
    fullPdf = fullPdf & "<< /Type /Catalog /Pages 2 0 R >>" & vbCrLf
    fullPdf = fullPdf & "endobj" & vbCrLf
    xrefOut(1) = objSegmentStart
    
    ' OBJ 2: Pages Tree
    objSegmentStart = Len(fullPdf)
    ' Collect all page refs for children array
    childRefs = BuildPageChildReferences(textData)
    fullPdf = fullPdf & "2 0 obj" & vbCrLf
    fullPdf = fullPdf & "<< /Type /Pages /Kids [" & childRefs & "] /Count " & totalPages & " >>" & vbCrLf
    fullPdf = fullPdf & "endobj" & vbCrLf
    xrefOut(2) = objSegmentStart
    
    ' OBJ 3+: Individual Pages (one per page, but grouping logic simplifies to single page with multi-content)
    ' For simplicity, let's consolidate text onto fewer pages with proper streaming
    
    Dim consolidatedPages
    consolidatedPages = ConsolidateTextOntoPages(textData)
    
    ' Build content streams for each page
    Dim pageObjsStart, contentLength
    pageObjsStart = totalObjects
    
    For i = 0 To UBound(consolidatedPages)
        ' Each page needs:
        ' 1. Page dictionary object  
        ' 2. Contents stream object
        
        ' Page Dictionary
        objSegmentStart = Len(fullPdf)
        fullPdf = fullPdf & CStr(pageObjsStart + (i*2) + 1) & " 0 obj" & vbCrLf
        fullPdf = fullPdf << /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources <<" & vbCrLf
        fullPdf = fullPdf & "/Font << /F1 5 0 R >> >>" & vbCrLf
        fullPdf = fullPdf & "/Contents " & CStr(pageObjsStart + (i*2) + 2) & " 0 R >>" & vbCrLf
        fullPdf = fullPdf & "endobj" & vbCrLf
        xrefOut(pageObjsStart + (i*2) + 1) = objSegmentStart
        
        ' Content Stream
        contentLength = GenerateTextContentForPage(consolidatedPages(i))
        objSegmentStart = Len(fullPdf)
        fullPdf = fullPdf & CStr(pageObjsStart + (i*2) + 2) & " 0 obj" & vbCrLf
        fullPdf = fullPdf & "<< /Length " & contentLength & " >>" & vbCrLf
        fullPdf = fullPdf & "stream" & vbCrLf
        fullPdf = fullPdf & GenerateTextContentForPage(consolidatedPages(i)) & vbCrLf
        fullPdf = fullPdf & "endstream" & vbCrLf
        fullPdf = fullPdf & "endobj" & vbCrLf
        xrefOut(pageObjsStart + (i*2) + 2) = objSegmentStart
        
        totalPages = UBound(consolidatedPages) + 1
    Next
    
    ' FONT OBJECT (Helvetica Type1 - standard 14 fonts don't need embedding)
    objSegmentStart = Len(fullPdf)
    fullPdf = fullPdf & "5 0 obj" & vbCrLf
    fullPdf = fullPdf & "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>" & vbCrLf
    fullPdf = fullPdf & "endobj" & vbCrLf
    xrefOut(5) = objSegmentStart
    
    ' Build XREF table
    fullPdf = fullPdf & "xref" & vbCrLf
    fullPdf = fullPdf & "0 " & CStr(max(xrefOut, true) + 1) & vbCrLf
    
    ' Entry 0 (always free/object header marker)
    fullPdf = fullPdf & "0000000000 65535 f " & vbCrLf
    
    ' Entries for numbered objects (skip 0)
    For i = 1 To UBound(xrefOut)
        If xrefOut(i) > 0 Then
            ' Format: 10-digit zero-padded offset, then flags
            Dim paddedOffset
            paddedOffset = Right(String(10, "0") & CStr(xrefOut(i)), 10)
            fullPdf = fullPdf & paddedOffset & " 00000 n " & vbCrLf
        End If
    Next
    
    ' Trailer and EOF
    Dim trailerPos
    trailerPos = Len(fullPdf)
    fullPdf = fullPdf & "trailer" & vbCrLf
    fullPdf = fullPdf & "<< /Size " & CStr(UBound(xrefOut) + 1) & " /Root 1 0 R >>" & vbCrLf
    fullPdf = fullPdf & "startxref" & vbCrLf
    fullPdf = fullPdf & CStr(trailerPos) & vbCrLf
    fullPdf = fullPdf & "%%EOF" & vbCrLf
    
    ' Write final output
    Set binaryStream = CreateObject("ADODB.Stream")
    binaryStream.Open
    binaryStream.Type = 2         ' adTypeText
    binaryStream.Charset = "ascii"
    binaryStream.WriteText fullPdf
    binaryStream.SaveToFile outputFile, 2  ' adSaveCreateOverWrite
    binaryStream.Close
End Sub

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

Function Max(arr, returnCount)
    Dim maxVal, idx, uboundIdx
    maxVal = 0
    uboundIdx = UBound(arr)
    For idx = 0 To uboundIdx
        If arr(idx) > maxVal Then maxVal = arr(idx)
    Next
    If returnCount Then Max = uboundIdx Else Max = maxVal
End Function

Function Min(a, b)
    If a < b Then Min = a Else Min = b
End Function

Function IIf(condition, truePart, falsePart)
    If condition Then IIf = truePart Else IIf = falsePart
End Function

' Escape special characters for PDF string literals
Function EscapePdfString(str)
    str = Replace(str, "\", "\\")  ' Backslash first
    str = Replace(str, "(", "\(")  ' Open parenthesis
    str = Replace(str, ")", "\)")  ' Close parenthesis
    str = Replace(str, vbCrLf, "") ' Remove line breaks
    str = Replace(str, vbCr, "")   ' Carriage return
    str = Replace(str, vbLf, "")   ' Line feed
    EscapePdfString = str
End Function

' Build array of page references for Pages/Kids entry
Function BuildPageChildReferences(txtData)
    Dim refs, i, pageRefStart
    Refs = ""
    pageRefStart = 3  ' First page starts at object 3
    
    For i = 0 To totalPages - 1
        If Len(Refs) > 0 Then Refs = Refs & " "
        Refs = Refs & CStr(pageRefStart + (i * 2)) & " 0 R"
    Next
    
    BuildPageChildReferences = Refs
End Function

' Group text lines intelligently across pages
Function ConsolidateTextOntoPages(txtArray)
    Dim linesPerPage, estimatedLines, pageChunks, pageIndex, lineIndex
    Dim chunksAvailable, chunkCapacity, currentChunk, currentYPos
    
    ' Estimate lines per page (available vertical space ÷ line height)
    linesPerPage = Int((pageSizeH - topMargin - bottomMargin) / lineHeight)
    
    ReDim pageChunks((UBound(txtArray) \ linesPerPage) + 1)
    pageIndex = 0
    lineIndex = 0
    currentChunk = ""
    chunksAvailable = UBound(txtArray)
    
    For i = 0 To UBound(txtArray)
        If Len(Trim(txtArray(i))) > 0 Then
            If Len(currentChunk) > 0 Then
                currentChunk = currentChunk & "|NEWLINE|" & txtArray(i)
            Else
                currentChunk = txtArray(i)
            End If
            
            lineIndex = lineIndex + 1
            
            If lineIndex >= linesPerPage And i < UBound(txtArray) Then
                ' Page break threshold reached
                pageChunks(pageIndex) = currentChunk
                pageIndex = pageIndex + 1
                currentChunk = ""
                lineIndex = 0
            End If
        End If
    Next
    
    ' Flush final chunk
    If Len(Trim(currentChunk)) > 0 Then
        pageChunks(pageIndex) = currentChunk
        pageIndex = pageIndex + 1
    End If
    
    ReDim Preserve pageChunks(pageIndex - 1)
    ConsolidateTextOntoPages = pageChunks
End Function

' Generate PDF Tj operator sequence for text positioning
Function GenerateTextContentForPage(pageText)
    Dim lines, i, yPos, startX, encodedLine
    Dim contentBuilder
    
    lines = Split(pageText, "|NEWLINE|")
    startY = pageSizeH - topMargin
    xPos = marginLeft
    
    contentBuilder = "BT" & vbCrLf  ' Begin text block
    contentBuilder = contentBuilder & "/F1 " & fontSize & " Tf" & vbCrLf  ' Helvetica 12pt
    contentBuilder = contentBuilder & "0.0 0.0 0.0 rg" & vbCrLf           ' Black color
    
    For i = 0 To UBound(lines)
        yPos = startY - ((i + 1) * lineHeight)
        
        ' Check if exceeding bottom margin - should not happen if consolidation worked
        If yPos < bottomMargin + fontSize Then
            ' Would overflow - truncation warning could go here
            Exit For
        End If
        
        encodedLine = EscapePdfString(lines(i))
        contentBuilder = contentBuilder & CStr(xPos) & " " & CStr(yPos) & " Td" & vbCrLf
        contentBuilder = contentBuilder & "(" & encodedLine & ") Tj" & vbCrLf
    Next
    
    contentBuilder = contentBuilder & "ET" & vbCrLf  ' End text block
    
    GenerateTextContentForPage = contentBuilder
End Function

' Display usage information
Sub ShowUsage()
    Dim msg
    msg = "========================================" & vbCrLf
    msg = msg & "TXT to PDF Converter - Pure VBScript" & vbCrLf
    msg = msg & "========================================" & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "SYNTAX:" & vbCrLf
    msg = msg & "  cscript txt2pdf.vbs input.txt [output.pdf]" & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "PARAMETERS:" & vbCrLf
    msg = msg & "  input.txt    - Source text file (required)" & vbCrLf
    msg = msg & "  output.pdf   - Destination file (optional, defaults to input.pdf)" & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "FEATURES:" & vbCrLf
    msg = msg & "  • Font: Helvetica 12 point" & vbCrLf
    msg = msg & "  • Automatic word-wrapping/reflow" & vbCrLf
    msg = msg & "  • Handles CRLF and LF line endings" & vbCrLf
    msg = msg & "  • A4 paper size (595×842 points)" & vbCrLf
    msg = msg & "  • 1-inch margins on all sides" & vbCrLf
    msg = msg & "  • Zero external dependencies" & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "EXAMPLES:" & vbCrLf
    msg = msg & "  cscript txt2pdf.vbs document.txt" & vbCrLf
    msg = msg & "  cscript txt2pdf.vbs notes.txt mynotes.pdf" & vbCrLf
    msg = msg & "  wscript txt2pdf.vbs report.doc --gui-mode" & vbCrLf
    msg = msg & vbCrLf
    MsgBox msg, vbInformation, "TXT to PDF Converter"
End Sub

' Clean shutdown handler
On Error Resume Next
Set binaryStream = Nothing
Set tempStream = Nothing
Set fso = Nothing
Set WSHArgs = Nothing
OnError Goto 0