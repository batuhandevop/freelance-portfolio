Attribute VB_Name = "BatchFileProcessor"
'===============================================================================
' Module:      BatchFileProcessor
' Author:      Batuhan
' Description: Scans a user-selected folder for .xlsx and .csv files, opens
'              each one, extracts key metadata and summary statistics, then
'              consolidates everything into a professionally formatted master
'              summary workbook with a clickable table of contents.
'
' Usage:       Run the "ProcessFolder" macro from the Macros dialog (Alt+F8)
'              or assign it to a button / ribbon command.
'
' Requirements: Microsoft Excel (tested on 2016 / 2019 / 365)
'               Reference: Microsoft Scripting Runtime (Tools > References)
'===============================================================================
Option Explicit

' ---------------------------------------------------------------------------
' Constants
' ---------------------------------------------------------------------------
Private Const APP_NAME          As String = "Batch File Processor"
Private Const TOC_SHEET_NAME    As String = "Table of Contents"
Private Const MAX_PREVIEW_ROWS  As Long = 5       ' rows of sample data to show
Private Const HEADER_COLOR      As Long = 4474111  ' light blue  (RGB 255,200,68 stored as BGR)
Private Const ACCENT_COLOR      As Long = 16247773 ' soft grey-blue

' ---------------------------------------------------------------------------
' Custom type to hold per-file results
' ---------------------------------------------------------------------------
Private Type FileInfo
    FileName        As String
    FilePath        As String
    FileSize        As String
    SheetCount      As Long
    RowCount        As Long
    ColumnCount     As Long
    Headers()       As String
    NumericColCount As Long
    MinVals()       As Variant
    MaxVals()       As Variant
    AvgVals()       As Variant
    SummarySheet    As String   ' name of the sheet created in the master workbook
    ErrorMsg        As String   ' non-empty if the file could not be processed
End Type

' ===========================  PUBLIC ENTRY POINT  ===========================

Public Sub ProcessFolder()
'-------------------------------------------------------------------------------
' Main routine. Prompts the user for a folder, iterates over every .xlsx / .csv
' file found, extracts data, and builds the master summary workbook.
'-------------------------------------------------------------------------------
    Dim folderPath      As String
    Dim fso             As Object   ' Scripting.FileSystemObject
    Dim folder          As Object   ' Scripting.Folder
    Dim file            As Object   ' Scripting.File
    Dim fileList()      As FileInfo
    Dim fileCount       As Long
    Dim i               As Long
    Dim wbMaster        As Workbook
    Dim startTime       As Double

    ' --- Prompt user for folder ---
    folderPath = BrowseForFolder("Select the folder containing .xlsx / .csv files")
    If Len(folderPath) = 0 Then Exit Sub   ' user cancelled

    ' --- Enumerate matching files ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Folder not found: " & folderPath, vbExclamation, APP_NAME
        Exit Sub
    End If

    Set folder = fso.GetFolder(folderPath)
    fileCount = CountMatchingFiles(folder)

    If fileCount = 0 Then
        MsgBox "No .xlsx or .csv files found in:" & vbCrLf & folderPath, _
               vbInformation, APP_NAME
        Exit Sub
    End If

    ' --- Confirm before proceeding ---
    If MsgBox("Found " & fileCount & " file(s) to process." & vbCrLf & vbCrLf & _
              "This will create a new summary workbook. Continue?", _
              vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Sub

    ' --- Prepare environment ---
    startTime = Timer
    OptimizeExcel True
    ReDim fileList(1 To fileCount)

    ' --- Create the master workbook ---
    Set wbMaster = Workbooks.Add(xlWBATWorksheet)   ' single sheet
    wbMaster.Sheets(1).Name = TOC_SHEET_NAME

    ' --- Process each file ---
    i = 0
    For Each file In folder.Files
        If IsTargetFile(file.Name) Then
            i = i + 1
            UpdateProgress i, fileCount, file.Name
            fileList(i) = ExtractFileData(file, wbMaster, i)
        End If
    Next file

    ' --- Build the Table of Contents ---
    BuildTableOfContents wbMaster, fileList, fileCount

    ' --- Final formatting pass on the TOC sheet ---
    FormatTOCSheet wbMaster.Sheets(TOC_SHEET_NAME), fileCount

    ' --- Activate TOC sheet ---
    wbMaster.Sheets(TOC_SHEET_NAME).Activate

    ' --- Restore environment ---
    OptimizeExcel False
    Application.StatusBar = False

    ' --- Done ---
    MsgBox "Processing complete." & vbCrLf & vbCrLf & _
           "Files processed: " & fileCount & vbCrLf & _
           "Elapsed time: " & Format(Timer - startTime, "0.0") & " seconds", _
           vbInformation, APP_NAME
End Sub

' ==========================  PRIVATE HELPERS  ===============================

' ---------------------------------------------------------------------------
' BrowseForFolder  -  Shows a folder-picker dialog and returns the path.
' ---------------------------------------------------------------------------
Private Function BrowseForFolder(prompt As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = prompt
        .AllowMultiSelect = False
        If .Show = -1 Then
            BrowseForFolder = .SelectedItems(1)
            ' Ensure trailing separator
            If Right(BrowseForFolder, 1) <> Application.PathSeparator Then
                BrowseForFolder = BrowseForFolder & Application.PathSeparator
            End If
        End If
    End With
End Function

' ---------------------------------------------------------------------------
' CountMatchingFiles  -  Returns the number of .xlsx / .csv files in a folder.
' ---------------------------------------------------------------------------
Private Function CountMatchingFiles(folder As Object) As Long
    Dim f As Object
    Dim n As Long
    For Each f In folder.Files
        If IsTargetFile(f.Name) Then n = n + 1
    Next f
    CountMatchingFiles = n
End Function

' ---------------------------------------------------------------------------
' IsTargetFile  -  Returns True for .xlsx and .csv extensions.
' ---------------------------------------------------------------------------
Private Function IsTargetFile(fileName As String) As Boolean
    Dim ext As String
    ext = LCase(Right(fileName, InStr(StrReverse(fileName), ".") ))
    IsTargetFile = (ext = ".xlsx" Or ext = ".csv")
End Function

' ---------------------------------------------------------------------------
' ExtractFileData  -  Opens one file, reads its metadata and stats, writes a
'                     detail sheet into the master workbook, then closes it.
' ---------------------------------------------------------------------------
Private Function ExtractFileData(file As Object, _
                                  wbMaster As Workbook, _
                                  idx As Long) As FileInfo
    Dim fi          As FileInfo
    Dim wbSource    As Workbook
    Dim wsSource    As Worksheet
    Dim wsDetail    As Worksheet
    Dim usedRng     As Range
    Dim lastRow     As Long
    Dim lastCol     As Long
    Dim col         As Long
    Dim numericCols As Long
    Dim r           As Long

    fi.FileName = file.Name
    fi.FilePath = file.Path
    fi.FileSize = FormatFileSize(file.Size)

    ' --- Attempt to open the source file ---
    On Error GoTo FileOpenError
    Set wbSource = Workbooks.Open( _
        fileName:=file.Path, _
        UpdateLinks:=0, _
        ReadOnly:=True, _
        CorruptLoad:=xlRepairFile)
    On Error GoTo 0

    ' --- Gather basic metadata ---
    fi.SheetCount = wbSource.Sheets.Count
    Set wsSource = wbSource.Sheets(1)       ' analyse the first sheet
    Set usedRng = wsSource.UsedRange

    If usedRng Is Nothing Then
        fi.RowCount = 0
        fi.ColumnCount = 0
    Else
        lastRow = usedRng.Rows.Count + usedRng.Row - 1
        lastCol = usedRng.Columns.Count + usedRng.Column - 1
        fi.RowCount = lastRow
        fi.ColumnCount = lastCol
    End If

    ' --- Extract headers (first row) ---
    If fi.ColumnCount > 0 Then
        ReDim fi.Headers(1 To fi.ColumnCount)
        For col = 1 To fi.ColumnCount
            fi.Headers(col) = CStr(wsSource.Cells(1, col).Value)
        Next col
    End If

    ' --- Compute summary statistics for numeric columns ---
    ReDim fi.MinVals(1 To fi.ColumnCount)
    ReDim fi.MaxVals(1 To fi.ColumnCount)
    ReDim fi.AvgVals(1 To fi.ColumnCount)
    numericCols = 0

    If fi.RowCount > 1 Then
        For col = 1 To fi.ColumnCount
            If IsNumericColumn(wsSource, col, 2, fi.RowCount) Then
                numericCols = numericCols + 1
                On Error Resume Next
                fi.MinVals(col) = Application.WorksheetFunction.Min( _
                    wsSource.Range(wsSource.Cells(2, col), wsSource.Cells(fi.RowCount, col)))
                fi.MaxVals(col) = Application.WorksheetFunction.Max( _
                    wsSource.Range(wsSource.Cells(2, col), wsSource.Cells(fi.RowCount, col)))
                fi.AvgVals(col) = Application.WorksheetFunction.Average( _
                    wsSource.Range(wsSource.Cells(2, col), wsSource.Cells(fi.RowCount, col)))
                On Error GoTo 0
            End If
        Next col
    End If
    fi.NumericColCount = numericCols

    ' --- Create a detail sheet in the master workbook ---
    fi.SummarySheet = SanitizeSheetName(fi.FileName, idx)
    Set wsDetail = wbMaster.Sheets.Add(After:=wbMaster.Sheets(wbMaster.Sheets.Count))
    wsDetail.Name = fi.SummarySheet

    WriteDetailSheet wsDetail, fi, wsSource

    ' --- Close source workbook without saving ---
    wbSource.Close SaveChanges:=False

    ExtractFileData = fi
    Exit Function

FileOpenError:
    fi.ErrorMsg = "Could not open file: " & Err.Description
    Err.Clear
    On Error GoTo 0
    ' Still create a detail sheet noting the error
    fi.SummarySheet = SanitizeSheetName(fi.FileName, idx)
    Set wsDetail = wbMaster.Sheets.Add(After:=wbMaster.Sheets(wbMaster.Sheets.Count))
    wsDetail.Name = fi.SummarySheet
    wsDetail.Range("A1").Value = "ERROR"
    wsDetail.Range("A2").Value = fi.ErrorMsg
    wsDetail.Range("A1").Font.Bold = True
    wsDetail.Range("A1").Font.Color = vbRed
    ExtractFileData = fi
End Function

' ---------------------------------------------------------------------------
' WriteDetailSheet  -  Populates the per-file summary sheet in the master
'                      workbook with metadata, headers, stats, and a data
'                      preview.
' ---------------------------------------------------------------------------
Private Sub WriteDetailSheet(ws As Worksheet, fi As FileInfo, wsSource As Worksheet)
    Dim r As Long, col As Long

    ' --- Title ---
    r = 1
    ws.Cells(r, 1).Value = "File Summary: " & fi.FileName
    ws.Cells(r, 1).Font.Size = 14
    ws.Cells(r, 1).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Merge

    ' --- Metadata block ---
    r = 3
    WriteKVPair ws, r, "File Path:", fi.FilePath: r = r + 1
    WriteKVPair ws, r, "File Size:", fi.FileSize: r = r + 1
    WriteKVPair ws, r, "Sheet Count:", fi.SheetCount: r = r + 1
    WriteKVPair ws, r, "Total Rows:", fi.RowCount: r = r + 1
    WriteKVPair ws, r, "Total Columns:", fi.ColumnCount: r = r + 1
    WriteKVPair ws, r, "Numeric Columns:", fi.NumericColCount: r = r + 1

    ' --- Column Headers ---
    r = r + 1
    ws.Cells(r, 1).Value = "Column Headers"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 12
    r = r + 1

    If fi.ColumnCount > 0 Then
        ws.Cells(r, 1).Value = "#"
        ws.Cells(r, 2).Value = "Column Name"
        ws.Cells(r, 3).Value = "Min"
        ws.Cells(r, 4).Value = "Max"
        ws.Cells(r, 5).Value = "Average"
        FormatHeaderRow ws, r, 5
        r = r + 1

        For col = 1 To fi.ColumnCount
            ws.Cells(r, 1).Value = col
            ws.Cells(r, 2).Value = fi.Headers(col)
            If Not IsEmpty(fi.MinVals(col)) Then
                ws.Cells(r, 3).Value = fi.MinVals(col)
                ws.Cells(r, 3).NumberFormat = "#,##0.00"
                ws.Cells(r, 4).Value = fi.MaxVals(col)
                ws.Cells(r, 4).NumberFormat = "#,##0.00"
                ws.Cells(r, 5).Value = fi.AvgVals(col)
                ws.Cells(r, 5).NumberFormat = "#,##0.00"
            End If
            r = r + 1
        Next col
    End If

    ' --- Data Preview ---
    r = r + 1
    ws.Cells(r, 1).Value = "Data Preview (first " & MAX_PREVIEW_ROWS & " rows)"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 12
    r = r + 1

    If fi.ColumnCount > 0 And fi.RowCount > 1 Then
        ' Write header row
        For col = 1 To fi.ColumnCount
            ws.Cells(r, col).Value = fi.Headers(col)
        Next col
        FormatHeaderRow ws, r, fi.ColumnCount
        r = r + 1

        ' Write data rows
        Dim previewRows As Long
        previewRows = Application.WorksheetFunction.Min(MAX_PREVIEW_ROWS, fi.RowCount - 1)
        Dim dataRow As Long
        For dataRow = 1 To previewRows
            For col = 1 To fi.ColumnCount
                ws.Cells(r, col).Value = wsSource.Cells(dataRow + 1, col).Value
            Next col
            r = r + 1
        Next dataRow
    End If

    ' --- "Back to TOC" hyperlink ---
    r = r + 2
    ws.Hyperlinks.Add Anchor:=ws.Cells(r, 1), _
                       Address:="", _
                       SubAddress:="'" & TOC_SHEET_NAME & "'!A1", _
                       TextToDisplay:="<< Back to Table of Contents"
    ws.Cells(r, 1).Font.Size = 11

    ' --- Auto-fit ---
    ws.Columns("A:Z").AutoFit
End Sub

' ---------------------------------------------------------------------------
' BuildTableOfContents  -  Creates the TOC sheet with a summary row per file
'                          and hyperlinks to each detail sheet.
' ---------------------------------------------------------------------------
Private Sub BuildTableOfContents(wbMaster As Workbook, _
                                  fileList() As FileInfo, _
                                  fileCount As Long)
    Dim ws As Worksheet
    Dim r  As Long
    Dim i  As Long

    Set ws = wbMaster.Sheets(TOC_SHEET_NAME)

    ' --- Title ---
    r = 1
    ws.Cells(r, 1).Value = APP_NAME & " - Summary Report"
    ws.Cells(r, 1).Font.Size = 16
    ws.Cells(r, 1).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Merge

    r = 2
    ws.Cells(r, 1).Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(r, 1).Font.Italic = True
    ws.Cells(r, 1).Font.Color = RGB(100, 100, 100)

    r = 3
    ws.Cells(r, 1).Value = "Total files processed: " & fileCount
    ws.Cells(r, 1).Font.Italic = True
    ws.Cells(r, 1).Font.Color = RGB(100, 100, 100)

    ' --- Header row ---
    r = 5
    ws.Cells(r, 1).Value = "#"
    ws.Cells(r, 2).Value = "File Name"
    ws.Cells(r, 3).Value = "Size"
    ws.Cells(r, 4).Value = "Sheets"
    ws.Cells(r, 5).Value = "Rows"
    ws.Cells(r, 6).Value = "Columns"
    ws.Cells(r, 7).Value = "Status"
    FormatHeaderRow ws, r, 7

    ' --- Data rows ---
    For i = 1 To fileCount
        r = 5 + i
        ws.Cells(r, 1).Value = i

        ' Clickable file name linking to its detail sheet
        If Len(fileList(i).SummarySheet) > 0 Then
            ws.Hyperlinks.Add Anchor:=ws.Cells(r, 2), _
                               Address:="", _
                               SubAddress:="'" & fileList(i).SummarySheet & "'!A1", _
                               TextToDisplay:=fileList(i).FileName
        Else
            ws.Cells(r, 2).Value = fileList(i).FileName
        End If

        ws.Cells(r, 3).Value = fileList(i).FileSize
        ws.Cells(r, 4).Value = fileList(i).SheetCount
        ws.Cells(r, 5).Value = fileList(i).RowCount
        ws.Cells(r, 5).NumberFormat = "#,##0"
        ws.Cells(r, 6).Value = fileList(i).ColumnCount

        If Len(fileList(i).ErrorMsg) > 0 Then
            ws.Cells(r, 7).Value = "Error"
            ws.Cells(r, 7).Font.Color = vbRed
        Else
            ws.Cells(r, 7).Value = "OK"
            ws.Cells(r, 7).Font.Color = RGB(0, 128, 0)
        End If

        ' Alternate row shading
        If i Mod 2 = 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Interior.Color = RGB(245, 245, 250)
        End If
    Next i

    ' --- Totals row ---
    r = 6 + fileCount
    ws.Cells(r, 1).Value = ""
    ws.Cells(r, 2).Value = "TOTAL"
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 5).Formula = "=SUM(E6:E" & (5 + fileCount) & ")"
    ws.Cells(r, 5).Font.Bold = True
    ws.Cells(r, 5).NumberFormat = "#,##0"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Borders(xlEdgeTop).LineStyle = xlDouble
End Sub

' ---------------------------------------------------------------------------
' FormatTOCSheet  -  Applies final professional formatting to the TOC.
' ---------------------------------------------------------------------------
Private Sub FormatTOCSheet(ws As Worksheet, fileCount As Long)
    Dim lastRow As Long
    lastRow = 6 + fileCount   ' includes totals row

    ' Column widths
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 10

    ' Thin border around data area
    Dim dataRange As Range
    Set dataRange = ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, 7))
    With dataRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 180, 180)
    End With

    ' Freeze panes below header
    ws.Activate
    ws.Range("A6").Select
    ActiveWindow.FreezePanes = True
End Sub

' ---------------------------------------------------------------------------
' FormatHeaderRow  -  Styles a row as a table header (bold, coloured, borders).
' ---------------------------------------------------------------------------
Private Sub FormatHeaderRow(ws As Worksheet, rowNum As Long, colCount As Long)
    With ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, colCount))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(55, 86, 135)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .HorizontalAlignment = xlCenter
    End With
End Sub

' ---------------------------------------------------------------------------
' WriteKVPair  -  Writes a label in column A (bold) and value in column B.
' ---------------------------------------------------------------------------
Private Sub WriteKVPair(ws As Worksheet, r As Long, label As String, value As Variant)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Value = value
End Sub

' ---------------------------------------------------------------------------
' IsNumericColumn  -  Checks whether a column is predominantly numeric by
'                     sampling the first few data rows.
' ---------------------------------------------------------------------------
Private Function IsNumericColumn(ws As Worksheet, col As Long, _
                                  startRow As Long, endRow As Long) As Boolean
    Dim sampleSize As Long
    Dim numCount   As Long
    Dim r          As Long

    sampleSize = Application.WorksheetFunction.Min(10, endRow - startRow + 1)

    For r = startRow To startRow + sampleSize - 1
        If IsNumeric(ws.Cells(r, col).Value) And _
           Not IsEmpty(ws.Cells(r, col).Value) Then
            numCount = numCount + 1
        End If
    Next r

    ' Consider numeric if >= 70% of sampled values are numbers
    IsNumericColumn = (numCount / sampleSize >= 0.7)
End Function

' ---------------------------------------------------------------------------
' SanitizeSheetName  -  Creates a valid, unique Excel sheet name (max 31 chars,
'                       no special characters).
' ---------------------------------------------------------------------------
Private Function SanitizeSheetName(fileName As String, idx As Long) As String
    Dim baseName As String
    Dim i As Long

    ' Strip extension
    baseName = Left(fileName, InStrRev(fileName, ".") - 1)

    ' Remove illegal characters:  \ / * [ ] : ?
    Dim illegal As Variant
    illegal = Array("\", "/", "*", "[", "]", ":", "?")
    For i = LBound(illegal) To UBound(illegal)
        baseName = Replace(baseName, illegal(i), "_")
    Next i

    ' Prefix with index for uniqueness; truncate to 31 characters
    baseName = idx & "_" & baseName
    If Len(baseName) > 31 Then baseName = Left(baseName, 31)

    SanitizeSheetName = baseName
End Function

' ---------------------------------------------------------------------------
' FormatFileSize  -  Converts bytes to a human-readable string (KB / MB).
' ---------------------------------------------------------------------------
Private Function FormatFileSize(sizeInBytes As Variant) As String
    Dim sz As Double
    sz = CDbl(sizeInBytes)
    Select Case sz
        Case Is >= 1073741824
            FormatFileSize = Format(sz / 1073741824, "#,##0.0") & " GB"
        Case Is >= 1048576
            FormatFileSize = Format(sz / 1048576, "#,##0.0") & " MB"
        Case Is >= 1024
            FormatFileSize = Format(sz / 1024, "#,##0.0") & " KB"
        Case Else
            FormatFileSize = sz & " bytes"
    End Select
End Function

' ---------------------------------------------------------------------------
' UpdateProgress  -  Shows progress in the Excel status bar.
' ---------------------------------------------------------------------------
Private Sub UpdateProgress(current As Long, total As Long, fileName As String)
    Dim pct As Long
    pct = Int((current / total) * 100)
    Application.StatusBar = APP_NAME & ": Processing file " & current & " of " & _
                            total & " (" & pct & "%) - " & fileName
    DoEvents   ' allow the UI to repaint
End Sub

' ---------------------------------------------------------------------------
' OptimizeExcel  -  Toggles performance-critical Application settings.
' ---------------------------------------------------------------------------
Private Sub OptimizeExcel(turboOn As Boolean)
    With Application
        .ScreenUpdating = Not turboOn
        .EnableEvents = Not turboOn
        .Calculation = IIf(turboOn, xlCalculationManual, xlCalculationAutomatic)
        .DisplayAlerts = Not turboOn
    End With
End Sub
