Attribute VB_Name = "Module1"
Option Explicit

' Helper function to check if a sheet exists in a workbook
Function SheetExists(wb As Workbook, nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(nm)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' Helper function to get a unique sheet name (max 31 chars, valid chars only)
Function GetUniqueSheetName(wb As Workbook, baseName As String) As String
    Dim nm As String, suffix As Integer
    Dim invalidChars As Variant, c As Variant

    ' Replace invalid characters
    invalidChars = Array("\", "/", "?", "*", "[", "]", ":")
    nm = baseName
    For Each c In invalidChars
        nm = Replace(nm, c, "_")
    Next c

    ' Truncate to 25 chars to leave room for suffix
    If Len(nm) > 25 Then nm = Left(nm, 25)
    GetUniqueSheetName = nm
    suffix = 1

    Do While SheetExists(wb, GetUniqueSheetName)
        GetUniqueSheetName = nm & "_" & suffix
        suffix = suffix + 1
        If Len(GetUniqueSheetName) > 31 Then GetUniqueSheetName = Left(nm, 25 - Len(CStr(suffix)) - 1) & "_" & suffix
    Loop
End Function

Sub CombineExcelFilesWithAllEnhancements()
    Dim FSO As Object, FolderPath As String, fDialog As FileDialog, FileList As Collection
    Dim wbSource As Workbook, wsSource As Worksheet, wsDest As Worksheet, ThisWB As Workbook
    Dim Headers As Object, MasterHeaders As Variant, SourceHeaders As Variant
    Dim i As Long, j As Long, LastRowSource As Long, LastRowDest As Long, ColCount As Long
    Dim ColMap As Object, arrSource As Variant, arrRow As Variant
    Dim ErrorLog As Worksheet, ErrorCount As Long
    Dim Filename As Variant, FileExt As String, SheetName As String
    Dim BackupName As String, UserInput As String, Answer As VbMsgBoxResult
    Dim RowKey As String, DuplicateDict As Object
    Dim StatusMsg As String, FileCounter As Long, TotalRows As Long
    Dim wsCheck As Worksheet
    Dim wsList As Collection, wsItem As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Starting: Preparing to combine files..."

    Set ThisWB = ThisWorkbook

    ' --- 1. Pick folder ---
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    If fDialog.Show = -1 Then
        FolderPath = fDialog.SelectedItems(1) & "\"
    Else
        MsgBox "No folder selected. Exiting.", vbExclamation
        GoTo CleanUp
    End If

    ' --- 2. Gather files ---
    Set FileList = New Collection
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim file As Object
    For Each file In FSO.GetFolder(FolderPath).Files
        FileExt = LCase(FSO.GetExtensionName(file.Name))
        If FileExt = "xls" Or FileExt = "xlsx" Or FileExt = "xlsm" Or FileExt = "csv" Then
            If file.Name <> ThisWB.Name Then FileList.Add file.Name
        End If
    Next file
    If FileList.Count = 0 Then
        MsgBox "No Excel or CSV files found in folder.", vbExclamation
        GoTo CleanUp
    End If

    ' --- 3. Get sheet selection ---
    Answer = MsgBox("Combine ALL sheets in every file?" & vbCrLf & _
                    "Click YES to combine ALL sheets." & vbCrLf & _
                    "Click NO to combine a specific sheet name only.", vbYesNoCancel + vbQuestion, "Sheet Selection")
    If Answer = vbCancel Then GoTo CleanUp
    If Answer = vbNo Then
        UserInput = InputBox("Enter the name of the sheet to combine (case-sensitive):", "Sheet Name", "Sheet1")
        If UserInput = "" Then GoTo CleanUp
        SheetName = UserInput
    Else
        SheetName = ""
    End If

    ' --- 4. Handle existing CombinedData sheet with safe renaming ---
    Set wsDest = Nothing
    For Each wsCheck In ThisWB.Worksheets
        If wsCheck.Name = "CombinedData" Then
            BackupName = GetUniqueSheetName(ThisWB, "CombinedData_Backup_" & Format(Now, "yyyymmdd_hhnnss"))
            wsCheck.Name = BackupName
            Exit For
        End If
    Next wsCheck

    Set wsDest = ThisWB.Sheets.Add(After:=ThisWB.Sheets(ThisWB.Sheets.Count))
    wsDest.Name = "CombinedData"

    ' --- 5. Prepare error log ---
    Set ErrorLog = ThisWB.Sheets.Add(After:=wsDest)
    ErrorLog.Name = GetUniqueSheetName(ThisWB, "CombineErrors")
    ErrorLog.Cells(1, 1).Value = "File"
    ErrorLog.Cells(1, 2).Value = "Sheet"
    ErrorLog.Cells(1, 3).Value = "Error"
    ErrorCount = 2

    ' --- 6. Collect all unique headers (case-insensitive) ---
    Set Headers = CreateObject("Scripting.Dictionary")
    Headers.CompareMode = vbTextCompare
    FileCounter = 0
    For Each Filename In FileList
        FileCounter = FileCounter + 1
        Application.StatusBar = "Scanning headers (" & FileCounter & "/" & FileList.Count & "): " & Filename
        On Error Resume Next
        Set wbSource = Nothing
        FileExt = LCase(FSO.GetExtensionName(Filename))
        If FileExt = "csv" Then
            Set wbSource = Workbooks.Open(FolderPath & Filename, ReadOnly:=True, Format:=6)
        Else
            Set wbSource = Workbooks.Open(FolderPath & Filename, ReadOnly:=True)
        End If
        On Error GoTo 0
        If wbSource Is Nothing Then
            ErrorLog.Cells(ErrorCount, 1).Value = Filename
            ErrorLog.Cells(ErrorCount, 3).Value = "Could not open file"
            ErrorCount = ErrorCount + 1
            GoTo NextFileHeader
        End If

        If SheetName = "" Then
            For Each wsSource In wbSource.Worksheets
                SourceHeaders = wsSource.Rows(1).Value
                ColCount = wsSource.UsedRange.Columns.Count
                For i = 1 To ColCount
                    If Not Headers.exists(Trim(SourceHeaders(1, i))) And Trim(SourceHeaders(1, i)) <> "" Then
                        Headers.Add Trim(SourceHeaders(1, i)), Headers.Count + 1
                    End If
                Next i
            Next wsSource
        Else
            On Error Resume Next
            Set wsSource = wbSource.Sheets(SheetName)
            On Error GoTo 0
            If Not wsSource Is Nothing Then
                SourceHeaders = wsSource.Rows(1).Value
                ColCount = wsSource.UsedRange.Columns.Count
                For i = 1 To ColCount
                    If Not Headers.exists(Trim(SourceHeaders(1, i))) And Trim(SourceHeaders(1, i)) <> "" Then
                        Headers.Add Trim(SourceHeaders(1, i)), Headers.Count + 1
                    End If
                Next i
            Else
                ErrorLog.Cells(ErrorCount, 1).Value = Filename
                ErrorLog.Cells(ErrorCount, 2).Value = SheetName
                ErrorLog.Cells(ErrorCount, 3).Value = "Sheet not found"
                ErrorCount = ErrorCount + 1
            End If
        End If
        wbSource.Close SaveChanges:=False
NextFileHeader:
    Next Filename

    ' --- 7. (Optional) Custom column order ---
    If Headers.Count = 0 Then
        MsgBox "No headers found in files.", vbCritical
        GoTo CleanUp
    End If

    Answer = MsgBox("Do you want to specify a custom column order?", vbYesNo + vbQuestion)
    If Answer = vbYes Then
        Dim CustomOrder As String, OrderArr As Variant, ValidOrder As Boolean
        CustomOrder = InputBox("Enter comma-separated column names, as you want them in the output header row." & vbCrLf & _
            "Available columns: " & Join(Headers.Keys, ", "), "Custom Column Order", Join(Headers.Keys, ", "))
        OrderArr = Split(CustomOrder, ",")
        ValidOrder = True
        For i = LBound(OrderArr) To UBound(OrderArr)
            If Not Headers.exists(Trim(OrderArr(i))) Then
                ValidOrder = False
                Exit For
            End If
        Next i
        If ValidOrder Then
            ReDim MasterHeaders(0 To UBound(OrderArr))
            For i = LBound(OrderArr) To UBound(OrderArr)
                MasterHeaders(i) = Trim(OrderArr(i))
            Next i
        Else
            MsgBox "Invalid column order specified. Using default order.", vbExclamation
            MasterHeaders = Headers.Keys
        End If
    Else
        MasterHeaders = Headers.Keys
    End If

    ' --- 8. Write master headers ---
    For i = LBound(MasterHeaders) To UBound(MasterHeaders)
        wsDest.Cells(1, i + 1).Value = Trim(MasterHeaders(i))
    Next i

    ' --- 9. Prepare duplicate handling ---
    Answer = MsgBox("Skip duplicate rows (based on all columns)?", vbYesNo + vbQuestion)
    Set DuplicateDict = CreateObject("Scripting.Dictionary")
    DuplicateDict.CompareMode = vbTextCompare

    ' --- 10. Combine data ---
    LastRowDest = 2
    FileCounter = 0
    TotalRows = 0
    For Each Filename In FileList
        FileCounter = FileCounter + 1
        Application.StatusBar = "Combining (" & FileCounter & "/" & FileList.Count & "): " & Filename
        On Error Resume Next
        Set wbSource = Nothing
        FileExt = LCase(FSO.GetExtensionName(Filename))
        If FileExt = "csv" Then
            Set wbSource = Workbooks.Open(FolderPath & Filename, ReadOnly:=True, Format:=6)
        Else
            Set wbSource = Workbooks.Open(FolderPath & Filename, ReadOnly:=True)
        End If
        On Error GoTo 0
        If wbSource Is Nothing Then
            ErrorLog.Cells(ErrorCount, 1).Value = Filename
            ErrorLog.Cells(ErrorCount, 3).Value = "Could not open file"
            ErrorCount = ErrorCount + 1
            GoTo NextFileData
        End If

        ' Build sheet list for this file
        Set wsList = New Collection
        If SheetName = "" Then
            For Each wsSource In wbSource.Worksheets
                wsList.Add wsSource
            Next wsSource
        Else
            On Error Resume Next
            Set wsSource = wbSource.Sheets(SheetName)
            On Error GoTo 0
            If Not wsSource Is Nothing Then
                wsList.Add wsSource
            End If
        End If

        For Each wsItem In wsList
            Set wsSource = wsItem
            If wsSource.UsedRange.Rows.Count < 2 Then GoTo NextSheet

            ' Map columns
            SourceHeaders = wsSource.Rows(1).Value
            ColCount = wsSource.UsedRange.Columns.Count
            Set ColMap = CreateObject("Scripting.Dictionary")
            ColMap.CompareMode = vbTextCompare
            For i = 1 To ColCount
                If Headers.exists(Trim(SourceHeaders(1, i))) Then
                    For j = LBound(MasterHeaders) To UBound(MasterHeaders)
                        If Trim(SourceHeaders(1, i)) = Trim(MasterHeaders(j)) Then
                            ColMap.Add i, j + 1
                            Exit For
                        End If
                    Next j
                End If
            Next i

            LastRowSource = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
            If LastRowSource < 2 Then GoTo NextSheet

            arrSource = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(LastRowSource, ColCount)).Value

            For i = 1 To UBound(arrSource, 1)
                ' Skip hidden/filtered rows
                If wsSource.Rows(i + 1).Hidden Then GoTo NextRow

                ReDim arrRow(1 To UBound(MasterHeaders) + 1)
                For j = 1 To ColCount
                    If ColMap.exists(j) Then
                        arrRow(ColMap(j)) = arrSource(i, j)
                    End If
                Next j

                ' Duplicate check
                RowKey = Join(arrRow, Chr(30))
                If DuplicateDict.exists(RowKey) And Answer = vbYes Then GoTo NextRow
                DuplicateDict(RowKey) = True

                wsDest.Range(wsDest.Cells(LastRowDest, 1), wsDest.Cells(LastRowDest, UBound(MasterHeaders) + 1)).Value = arrRow
                LastRowDest = LastRowDest + 1
                TotalRows = TotalRows + 1
NextRow:
            Next i
NextSheet:
        Next wsItem

        wbSource.Close SaveChanges:=False
NextFileData:
    Next Filename

    ' --- 11. Finish up ---
    Application.StatusBar = False
    wsDest.UsedRange.Columns.AutoFit
    ErrorLog.UsedRange.Columns.AutoFit
    wsDest.Activate

    ' --- 12. Summary ---
    StatusMsg = "All files have been combined!" & vbCrLf & _
        "Total files processed: " & FileList.Count & vbCrLf & _
        "Total rows combined: " & TotalRows & vbCrLf & _
        "Errors logged: " & ErrorCount - 2 & vbCrLf & _
        "Check the """ & ErrorLog.Name & """ sheet for details if any errors occurred."
    MsgBox StatusMsg, vbInformation

CleanUp:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
End Sub

