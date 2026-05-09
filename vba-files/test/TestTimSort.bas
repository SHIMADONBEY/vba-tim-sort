Attribute VB_Name = "TestTimSort"
'/ ============================================================
'/ TestTimSort.bas
'/ Unit tests for VbaTimSort (SortArray / SortCollection)
'/
'/ NOTE:
'/   This module is for verification only.
'/
'/ How to use:
'/   1. Open the VBA editor (ALT+F11) and insert this module into the same workbook as VbaTimSort.
'/   2. Run the RunAll sub to execute all tests and write results to the "Results" sheet.
'/ ============================================================
'namespace=vba-files/test
Option Explicit

' -- Result sheet helpers -------------------------------------
Private Const RESULTS_SHEET As String = "Results"
Private mPassCount As Long
Private mFailCount As Long

Private Sub InitResults()
    mPassCount = 0
    mFailCount = 0
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(RESULTS_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = RESULTS_SHEET
    End If
    On Error GoTo 0
    ws.Cells.ClearContents
    ws.Range("A1").Value = "TestName"
    ws.Range("B1").Value = "Status"
    ws.Range("C1").Value = "Message"
    ws.Range("A1:C1").Font.Bold = True
End Sub

Private Sub RecordResult(ByVal testName As String, ByVal passed As Boolean, Optional ByVal message As String = "")
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(RESULTS_SHEET)
    Dim nextRow As Long: nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = testName
    If passed Then
        ws.Cells(nextRow, 2).Value = "PASS"
        ws.Cells(nextRow, 2).Font.Color = RGB(0, 128, 0)
        mPassCount = mPassCount + 1
    Else
        ws.Cells(nextRow, 2).Value = "FAIL"
        ws.Cells(nextRow, 2).Font.Color = RGB(200, 0, 0)
        ws.Cells(nextRow, 3).Value = message
        mFailCount = mFailCount + 1
    End If
End Sub

Private Sub WriteSummary()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(RESULTS_SHEET)
    Dim nextRow As Long: nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 2
    ws.Cells(nextRow, 1).Value = "Total: " & (mPassCount + mFailCount) & _
                                 "  PASS: " & mPassCount & _
                                 "  FAIL: " & mFailCount
    ws.Cells(nextRow, 1).Font.Bold = True
    ws.Columns("A:C").AutoFit
End Sub

' -- Assert helpers --------------------------------------------

'/ Checks that two arrays have the same length and equal elements at every index.
Private Function ArraysEqual(ByRef expected As Variant, ByRef actual As Variant) As Boolean
    If UBound(expected) - LBound(expected) <> UBound(actual) - LBound(actual) Then
        ArraysEqual = False
        Exit Function
    End If
    Dim i As Long
    For i = LBound(expected) To UBound(expected)
        Dim offset As Long: offset = LBound(actual) - LBound(expected)
        If expected(i) <> actual(i + offset) Then
            ArraysEqual = False
            Exit Function
        End If
    Next i
    ArraysEqual = True
End Function

'/ Builds a human-readable string from a Variant array (for failure messages).
Private Function ArrayToString(ByRef arr As Variant) As String
    Dim parts() As String
    ReDim parts(LBound(arr) To UBound(arr))
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        parts(i) = CStr(arr(i))
    Next i
    ArrayToString = "[" & Join(parts, ", ") & "]"
End Function

' -- Test cases -----------------------------------------------

' --- 1. Empty array / Single element / Multiple dimensions element ---

Private Sub Test_EmptyArray()
    Dim arr() As Variant
    ReDim arr(-1 To -1)
    Dim result As Variant
    result = VbaTimSort.SortArray(arr)
    Dim passed As Boolean: passed = (LBound(result) = -1 And UBound(result) = -1)
    RecordResult "Sort_EmptyArray", passed, IIf(passed, "", "Expected empty array bounds (-1 To -1)")
End Sub

Private Sub Test_SingleElement()
    Dim arr(0 To 0) As Variant
    arr(0) = 42
    Dim result As Variant
    result = VbaTimSort.SortArray(arr)
    Dim passed As Boolean: passed = (result(0) = 42)
    RecordResult "Sort_SingleElement", passed, IIf(passed, "", "Expected [42], got " & ArrayToString(result))
End Sub

Private Sub Test_MultiDimensionalElement()
    Dim raised As Boolean: raised = False
    Dim arr(0 To 1, 0 To 1) As Variant
    arr(0, 0) = 2: arr(0, 1) = 1
    arr(1, 0) = 1: arr(1, 1) = 2
    Dim result As Variant
    On Error Resume Next
    Err.Clear
    result = VbaTimSort.SortArray(arr)
    Dim errorNumber As Long: errorNumber = Err.Number
    If errorNumber = VbaTimSort.TIMSORT_ERR_ARG_INVALID_RANGE Then raised = True
    On Error GoTo 0
    RecordResult "Sort_MultiDimensionalElement_Error", raised, IIf(raised, "", "Expected error " & VbaTimSort.TIMSORT_ERR_ARG_INVALID_RANGE & " but got " & errorNumber)
End Sub

Private Sub Test_JaggedArray()
    Dim raised As Boolean: raised = False
    Dim arr(0 To 1) As Variant
    arr(0) = Array(3, 2, 1)
    arr(1) = Array(1, 2)
    Dim result As Variant
    On Error Resume Next
    Err.Clear
    result = VbaTimSort.SortArray(arr)
    Dim errorNumber As Long: errorNumber = Err.Number
    If errorNumber = VbaTimSort.TIMSORT_ERR_ARG_INVALID_RANGE Then raised = True
    On Error GoTo 0
    RecordResult "Sort_JaggedArray_Error", raised, IIf(raised, "", "Expected error " & VbaTimSort.TIMSORT_ERR_ARG_INVALID_RANGE & " but got " & errorNumber)
End Sub

' --- 2. Number Arrays (Long / Double) ---

Private Sub Test_Numbers_Ascending()
    Dim arr As Variant: arr = Array(5, 3, 8, 1, 2, 7, 4, 6)
    Dim expected As Variant: expected = Array(1, 2, 3, 4, 5, 6, 7, 8)
    Dim result As Variant: result = VbaTimSort.SortArray(arr)
    Dim passed As Boolean: passed = ArraysEqual(expected, result)
    RecordResult "Sort_Numbers_Ascending", passed, IIf(passed, "", "Expected " & ArrayToString(expected) & ", got " & ArrayToString(result))
End Sub

Private Sub Test_Numbers_Descending()
    Dim arr As Variant: arr = Array(5, 3, 8, 1, 2, 7, 4, 6)
    Dim expected As Variant: expected = Array(8, 7, 6, 5, 4, 3, 2, 1)
    Dim result As Variant: result = VbaTimSort.SortArray(arr, descending:=True)
    Dim passed As Boolean: passed = ArraysEqual(expected, result)
    RecordResult "Sort_Numbers_Descending", passed, IIf(passed, "", "Expected " & ArrayToString(expected) & ", got " & ArrayToString(result))
End Sub

Private Sub Test_Numbers_WithNegatives()
    Dim arr As Variant: arr = Array(0, -3, 5, -1, 2, -7)
    Dim expected As Variant: expected = Array(-7, -3, -1, 0, 2, 5)
    Dim result As Variant: result = VbaTimSort.SortArray(arr)
    Dim passed As Boolean: passed = ArraysEqual(expected, result)
    RecordResult "Sort_Numbers_WithNegatives", passed, IIf(passed, "", "Expected " & ArrayToString(expected) & ", got " & ArrayToString(result))
End Sub

Private Sub Test_Numbers_WithDecimals()
    Dim arr As Variant: arr = Array(1.5, 0.3, 2.7, 0.1, 1.1)
    Dim expected As Variant: expected = Array(0.1, 0.3, 1.1, 1.5, 2.7)
    Dim result As Variant: result = VbaTimSort.SortArray(arr)
    Dim passed As Boolean: passed = ArraysEqual(expected, result)
    RecordResult "Sort_Numbers_WithDecimals", passed, IIf(passed, "", "Expected " & ArrayToString(expected) & ", got " & ArrayToString(result))
End Sub

Private Sub Test_Numbers_Duplicates()
    Dim arr As Variant: arr = Array(3, 1, 2, 1, 3, 2)
    Dim expected As Variant: expected = Array(1, 1, 2, 2, 3, 3)
    Dim result As Variant: result = VbaTimSort.SortArray(arr)
    Dim passed As Boolean: passed = ArraysEqual(expected, result)
    RecordResult "Sort_Numbers_Duplicates", passed, IIf(passed, "", "Expected " & ArrayToString(expected) & ", got " & ArrayToString(result))
End Sub

Private Sub Test_Numbers_AllSame()
    Dim arr As Variant: arr = Array(5, 5, 5, 5, 5)
    Dim expected As Variant: expected = Array(5, 5, 5, 5, 5)
    Dim result As Variant: result = VbaTimSort.SortArray(arr)
    Dim passed As Boolean: passed = ArraysEqual(expected, result)
    RecordResult "Sort_Numbers_AllSame", passed, IIf(passed, "", "Expected " & ArrayToString(expected) & ", got " & ArrayToString(result))
End Sub

' --- 3. Boundary values: Already sorted / Reverse order ---

Private Sub Test_AlreadySorted()
    Dim arr As Variant: arr = Array(1, 2, 3, 4, 5, 6, 7, 8)
    Dim expected As Variant: expected = Array(1, 2, 3, 4, 5, 6, 7, 8)
    Dim result As Variant: result = VbaTimSort.SortArray(arr)
    Dim passed As Boolean: passed = ArraysEqual(expected, result)
    RecordResult "Sort_AlreadySorted", passed, IIf(passed, "", "Expected " & ArrayToString(expected) & ", got " & ArrayToString(result))
End Sub

Private Sub Test_ReverseSorted()
    Dim arr As Variant: arr = Array(8, 7, 6, 5, 4, 3, 2, 1)
    Dim expected As Variant: expected = Array(1, 2, 3, 4, 5, 6, 7, 8)
    Dim result As Variant: result = VbaTimSort.SortArray(arr)
    Dim passed As Boolean: passed = ArraysEqual(expected, result)
    RecordResult "Sort_ReverseSorted", passed, IIf(passed, "", "Expected " & ArrayToString(expected) & ", got " & ArrayToString(result))
End Sub

' --- 4. Boundary values: Size 32 / 64 (minRun boundary) / 1000 ---

Private Sub Test_Size32()
    Dim arr(1 To 32) As Variant
    Dim expected(1 To 32) As Variant
    Dim i As Long
    For i = 1 To 32
        arr(i) = 33 - i  ' Reverse order
        expected(i) = i
    Next i
    Dim result As Variant: result = VbaTimSort.SortArray(arr)
    Dim passed As Boolean: passed = ArraysEqual(expected, result)
    RecordResult "Sort_Size32", passed, IIf(passed, "", "Size-32 sort failed")
End Sub

Private Sub Test_Size64()
    Dim arr(1 To 64) As Variant
    Dim expected(1 To 64) As Variant
    Dim i As Long
    For i = 1 To 64
        arr(i) = 65 - i
        expected(i) = i
    Next i
    Dim result As Variant: result = VbaTimSort.SortArray(arr)
    Dim passed As Boolean: passed = ArraysEqual(expected, result)
    RecordResult "Sort_Size64", passed, IIf(passed, "", "Size-64 sort failed")
End Sub

Private Sub Test_Size1000()
    Dim arr(1 To 1000) As Variant
    Dim expected(1 To 1000) As Variant
    Dim i As Long
    For i = 1 To 1000
        arr(i) = 1001 - i
        expected(i) = i
    Next i
    Dim result As Variant: result = VbaTimSort.SortArray(arr)
    Dim passed As Boolean: passed = ArraysEqual(expected, result)
    RecordResult "Sort_Size1000", passed, IIf(passed, "", "Size-1000 sort failed")
End Sub

' --- 5. String arrays ---

Private Sub Test_Strings_Ascending()
    Dim arr As Variant: arr = Array("banana", "apple", "cherry", "date")
    Dim expected As Variant: expected = Array("apple", "banana", "cherry", "date")
    Dim result As Variant: result = VbaTimSort.SortArray(arr)
    Dim passed As Boolean: passed = ArraysEqual(expected, result)
    RecordResult "Sort_Strings_Ascending", passed, IIf(passed, "", "Expected " & ArrayToString(expected) & ", got " & ArrayToString(result))
End Sub

Private Sub Test_Strings_Descending()
    Dim arr As Variant: arr = Array("banana", "apple", "cherry", "date")
    Dim expected As Variant: expected = Array("date", "cherry", "banana", "apple")
    Dim result As Variant: result = VbaTimSort.SortArray(arr, descending:=True)
    Dim passed As Boolean: passed = ArraysEqual(expected, result)
    RecordResult "Sort_Strings_Descending", passed, IIf(passed, "", "Expected " & ArrayToString(expected) & ", got " & ArrayToString(result))
End Sub

' --- 6. Date arrays ---

Private Sub Test_Dates_Ascending()
    Dim arr(0 To 3) As Variant
    arr(0) = CDate("2024/03/15")
    arr(1) = CDate("2023/01/01")
    arr(2) = CDate("2024/12/31")
    arr(3) = CDate("2020/06/20")
    Dim expected(0 To 3) As Variant
    expected(0) = CDate("2020/06/20")
    expected(1) = CDate("2023/01/01")
    expected(2) = CDate("2024/03/15")
    expected(3) = CDate("2024/12/31")
    Dim result As Variant: result = VbaTimSort.SortArray(arr)
    Dim passed As Boolean: passed = ArraysEqual(expected, result)
    RecordResult "Sort_Dates_Ascending", passed, IIf(passed, "", "Date sort failed")
End Sub

' --- 7. Comparator (ComparatorNumber) ---

Private Sub Test_WithComparatorNumber()
    Dim cmp As New ComparatorNumber
    Dim arr As Variant: arr = Array(9, 1, 5, 3, 7)
    Dim expected As Variant: expected = Array(1, 3, 5, 7, 9)
    Dim result As Variant: result = VbaTimSort.SortArray(arr, cmp)
    Dim passed As Boolean: passed = ArraysEqual(expected, result)
    RecordResult "Sort_WithComparatorNumber", passed, IIf(passed, "", "Expected " & ArrayToString(expected) & ", got " & ArrayToString(result))
End Sub

' --- 8. Custom Object (PersonClass) ― Stability Test ---

Private Sub Test_Stability_CustomObject()
    '/ Verify that persons with the same Age retain their original order (stability test)
    Dim cmp As New ComparatorPerson
    Dim arr(0 To 5) As Variant

    Dim p0 As New PersonClass: p0.Name = "Alice":   p0.Age = 30: p0.OriginalIndex = 0
    Dim p1 As New PersonClass: p1.Name = "Bob":     p1.Age = 25: p1.OriginalIndex = 1
    Dim p2 As New PersonClass: p2.Name = "Charlie": p2.Age = 30: p2.OriginalIndex = 2
    Dim p3 As New PersonClass: p3.Name = "Diana":   p3.Age = 25: p3.OriginalIndex = 3
    Dim p4 As New PersonClass: p4.Name = "Eve":     p4.Age = 35: p4.OriginalIndex = 4
    Dim p5 As New PersonClass: p5.Name = "Frank":   p5.Age = 30: p5.OriginalIndex = 5

    Set arr(0) = p0: Set arr(1) = p1: Set arr(2) = p2
    Set arr(3) = p3: Set arr(4) = p4: Set arr(5) = p5

    Dim result As Variant: result = VbaTimSort.SortArray(arr, cmp)

    '/ Expected order by Age asc: Bob(25,1), Diana(25,3), Alice(30,0), Charlie(30,2), Frank(30,5), Eve(35,4)
    Dim expectedOrigIdx As Variant: expectedOrigIdx = Array(1, 3, 0, 2, 5, 4)
    Dim passed As Boolean: passed = True
    Dim i As Long
    For i = LBound(result) To UBound(result)
        Dim p As PersonClass: Set p = result(i)
        If p.OriginalIndex <> expectedOrigIdx(i - LBound(result)) Then
            passed = False
            Exit For
        End If
    Next i
    RecordResult "Sort_Stability_CustomObject", passed, IIf(passed, "", "Stability violated at index " & i)
End Sub

' --- 9. Collection sorting ---

Private Sub Test_SortCollection_Numbers()
    Dim coll As New Collection
    coll.Add 5: coll.Add 2: coll.Add 8: coll.Add 1: coll.Add 4
    Dim expected As Variant: expected = Array(1, 2, 4, 5, 8)
    Dim sortedColl As Collection
    Set sortedColl = VbaTimSort.SortCollection(coll)
    Dim passed As Boolean: passed = (sortedColl.Count = 5)
    Dim i As Long
    If passed Then
        For i = 1 To sortedColl.Count
            If sortedColl(i) <> expected(i - 1) Then
                passed = False
                Exit For
            End If
        Next i
    End If
    RecordResult "SortCollection_Numbers", passed, IIf(passed, "", "Collection sort failed at index " & i)
End Sub

Private Sub Test_SortCollection_Empty()
    Dim coll As New Collection
    Dim sortedColl As Collection
    Set sortedColl = VbaTimSort.SortCollection(coll)
    Dim passed As Boolean: passed = (sortedColl.Count = 0)
    RecordResult "SortCollection_Empty", passed, IIf(passed, "", "Expected empty collection")
End Sub

Private Sub Test_SortCollection_Nothing()
    Dim raised As Boolean: raised = False
    On Error Resume Next
    Err.Clear
    Dim sortedColl As Collection
    Set sortedColl = VbaTimSort.SortCollection(Nothing)
    Dim errorNumber As Long: errorNumber = Err.Number
    If errorNumber = VbaTimSort.TIMSORT_ERR_ARG_INVALID_RANGE Then raised = True
    On Error GoTo 0
    RecordResult "SortCollection_Nothing_Error", raised, IIf(raised, "", "Expected error " & VbaTimSort.TIMSORT_ERR_ARG_INVALID_RANGE & " but got " & errorNumber)
End Sub

' --- 10. Error handling ---

Private Sub Test_Error_NonArray()
    Dim raised As Boolean: raised = False
    On Error Resume Next
    Err.Clear
    Dim result As Variant: result = VbaTimSort.SortArray("not an array")
    Dim errorNumber As Long: errorNumber = Err.Number
    If errorNumber = VbaTimSort.TIMSORT_ERR_ARG_NOT_ARRAY Then raised = True
    On Error GoTo 0
    RecordResult "Error_NonArray_Error", raised, IIf(raised, "", "Expected error " & VbaTimSort.TIMSORT_ERR_ARG_NOT_ARRAY & " but got " & errorNumber)
End Sub

Private Sub Test_Error_ObjectWithoutComparator()
    Dim raised As Boolean: raised = False
    Dim arr(0 To 1) As Variant
    Set arr(0) = New PersonClass
    Set arr(1) = New PersonClass
    On Error Resume Next
    Err.Clear
    Dim result As Variant: result = VbaTimSort.SortArray(arr)
    Dim errorNumber As Long: errorNumber = Err.Number
    If errorNumber = VbaTimSort.TIMSORT_ERR_ARG_NO_COMPARATOR_FOR_OBJECTS Then raised = True
    On Error GoTo 0
    RecordResult "Error_ObjectWithoutComparator_Error", raised, IIf(raised, "", "Expected error " & VbaTimSort.TIMSORT_ERR_ARG_NO_COMPARATOR_FOR_OBJECTS & " but got " & errorNumber)
End Sub

' --- 11. Performance benchmarks ---

Private Sub Benchmark_Random(ByVal size As Long)
    Dim arr() As Variant
    ReDim arr(1 To size)
    Dim i As Long
    For i = 1 To size
        arr(i) = Int(Rnd() * 1000000)
    Next i
    Dim t As Double: t = Timer
    Dim result As Variant: result = VbaTimSort.SortArray(arr)
    Dim elapsed As Double: elapsed = Timer - t
    RecordResult "Benchmark_Random_" & size, True, Format(elapsed, "0.000") & " sec"
End Sub

Private Sub Benchmark_AlreadySorted(ByVal size As Long)
    Dim arr() As Variant
    ReDim arr(1 To size)
    Dim i As Long
    For i = 1 To size: arr(i) = i: Next i
    Dim t As Double: t = Timer
    Dim result As Variant: result = VbaTimSort.SortArray(arr)
    Dim elapsed As Double: elapsed = Timer - t
    RecordResult "Benchmark_AlreadySorted_" & size, True, Format(elapsed, "0.000") & " sec"
End Sub

Private Sub Benchmark_Reversed(ByVal size As Long)
    Dim arr() As Variant
    ReDim arr(1 To size)
    Dim i As Long
    For i = 1 To size: arr(i) = size + 1 - i: Next i
    Dim t As Double: t = Timer
    Dim result As Variant: result = VbaTimSort.SortArray(arr)
    Dim elapsed As Double: elapsed = Timer - t
    RecordResult "Benchmark_Reversed_" & size, True, Format(elapsed, "0.000") & " sec"
End Sub

' -- Public entry points --------------------------------------

'/ Runs all unit tests and writes results to the Results sheet.
Public Sub RunAll()
    Randomize
    InitResults

    ' Empty / Small size
    Test_EmptyArray
    Test_SingleElement
    Test_MultiDimensionalElement
    Test_JaggedArray

    ' Number arrays
    Test_Numbers_Ascending
    Test_Numbers_Descending
    Test_Numbers_WithNegatives
    Test_Numbers_WithDecimals
    Test_Numbers_Duplicates
    Test_Numbers_AllSame
    Test_AlreadySorted
    Test_ReverseSorted

    ' Boundary values
    Test_Size32
    Test_Size64
    Test_Size1000

    ' String arrays
    Test_Strings_Ascending
    Test_Strings_Descending

    ' Date arrays
    Test_Dates_Ascending

    ' Comparator
    Test_WithComparatorNumber

    ' Stability
    Test_Stability_CustomObject

    ' Collection
    Test_SortCollection_Numbers
    Test_SortCollection_Empty
    Test_SortCollection_Nothing

    ' Error handling
    Test_Error_NonArray
    Test_Error_ObjectWithoutComparator

    WriteSummary
    MsgBox "Tests complete.  PASS: " & mPassCount & "  FAIL: " & mFailCount, vbInformation, "VbaTimSort Test Results"
End Sub

'/ Runs only performance benchmarks. Results are appended to the Results sheet.
Public Sub RunBenchmarks()
    InitResults
    Randomize
    Dim sizes As Variant: sizes = Array(1000, 10000, 100000)
    Dim i As Long
    For i = LBound(sizes) To UBound(sizes)
        Benchmark_Random sizes(i)
        Benchmark_AlreadySorted sizes(i)
        Benchmark_Reversed sizes(i)
    Next i
    WriteSummary
    MsgBox "Benchmarks complete.", vbInformation, "VbaTimSort Benchmarks"
End Sub

' --- Headless runner (for automated test scripts) ---
' Adds RunAll_Headless(outPath) which runs the same tests and writes UTF-8 JSON.
Public Function RunAll_Headless(ByVal outPath As String) As Boolean
    On Error GoTo ErrHandler
    Randomize
    InitResults

    ' Run same tests as RunAll (call the same helpers)
    Test_EmptyArray
    Test_SingleElement
    Test_MultiDimensionalElement
    Test_JaggedArray
    Test_Numbers_Ascending
    Test_Numbers_Descending
    Test_Numbers_WithNegatives
    Test_Numbers_WithDecimals
    Test_Numbers_Duplicates
    Test_Numbers_AllSame
    Test_AlreadySorted
    Test_ReverseSorted
    Test_Size32
    Test_Size64
    Test_Size1000
    Test_Strings_Ascending
    Test_Strings_Descending
    Test_Dates_Ascending
    Test_WithComparatorNumber
    Test_Stability_CustomObject
    Test_SortCollection_Numbers
    Test_SortCollection_Empty
    Test_SortCollection_Nothing
    Test_Error_NonArray
    Test_Error_ObjectWithoutComparator

    ' Build JSON object before WriteSummary appends non-result rows
    Dim json As String
    json = BuildResultsJson()

    WriteSummary

    ' Write UTF-8 file using ADODB.Stream
    WriteUtf8File outPath, json

    ' Return overall pass/fail (true when no failures)
    RunAll_Headless = (mFailCount = 0)
    Exit Function

ErrHandler:
    RunAll_Headless = False
    On Error Resume Next
    Dim errObj As String
    errObj = "Error " & Err.Number & ": " & Err.Description
    WriteUtf8File outPath, "{""error"": true, ""message"": """ & JsonEscape(errObj) & """}"
End Function

' -- Helpers to produce JSON and write UTF-8 --
Private Function BuildResultsJson() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(RESULTS_SHEET)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    Dim items() As String
    ReDim items(0 To lastRow - 2) ' excluding header row
    Dim idx As Long: idx = 0
    For i = 2 To lastRow
        Dim t As String, s As String, m As String
        t = JsonEscape(CStr(ws.Cells(i, 1).Value))
        s = JsonEscape(CStr(ws.Cells(i, 2).Value))
        m = JsonEscape(CStr(ws.Cells(i, 3).Value))
        items(idx) = "{""testName"":""" & t & """, ""status"":""" & s & """, ""message"":""" & m & """}"
        idx = idx + 1
    Next i
    Dim body As String
    body = "[" & Join(items, ",") & "]"
    BuildResultsJson = "{""passed"":" & IIf(mFailCount = 0, "true", "false") & ", ""passCount"":" & CStr(mPassCount) & ", ""failCount"":" & CStr(mFailCount) & ", ""results"": " & body & "}"
End Function

' JSON-escape helper: escapes backslash, quotes and common control chars
Private Function JsonEscape(ByVal s As String) As String
    If Len(s) = 0 Then
        JsonEscape = ""
        Exit Function
    End If

    ' Escape existing backslashes first so original backslashes become \\
    s = Replace(s, "\", "\\")
    ' Escape double quotes (insert backslash + quote)
    s = Replace(s, """", "\""")
    ' Normalize CR/LF/CR/LF -> \n and tabs -> \t
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")

    JsonEscape = s
End Function

Private Sub WriteUtf8File(ByVal path As String, ByVal content As String)
    On Error GoTo ErrHandler
        ' tempPath is used to write the content first, then moved to the target path to ensure atomicity and avoid partial writes.
        ' tempPath is generated by appending random number and .tmp to the target path.
        ' prefer use guid for temp file name, but VBA doesn't have built-in support for that, so use random number with low collision probability instead.
        Randomize
        Dim randNum As Long: randNum = CLng(Rnd() * (2 ^ 30))
        Dim tmpPath As String: tmpPath = path & "." & CStr(randNum) & ".tmp"
        Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
        With stream
            .Type = 2 ' Text
            .Charset = "utf-8"
            .Open
            .WriteText content
            .SaveToFile tmpPath, 2 ' Overwrite
            .Close
        End With
        Set stream = Nothing

        ' Atomic move: first write to temp file, then move to target path to avoid partial writes.
        ' Keep normal error handling so delete/move failures are reported to the caller.
        Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FileExists(tmpPath) Then
            Err.Raise vbObjectError + 1000, "WriteUtf8File", "Temporary file was not created: " & tmpPath
        End If
        If fso.FileExists(path) Then fso.DeleteFile path
        fso.MoveFile tmpPath, path
        If Not fso.FileExists(path) Then
            Err.Raise vbObjectError + 1001, "WriteUtf8File", "Failed to move temporary file to target path: " & path
        End If
        Set fso = Nothing
    Exit Sub
ErrHandler:
        ' If writing fails, attempt to write error info to the target path
        On Error Resume Next
        If Not (stream Is Nothing) Then
            stream.Close
            Set stream = Nothing
        End If

        ' Best-effort cleanup of temp file.
        Dim fso1 As Object: Set fso1 = CreateObject("Scripting.FileSystemObject")
        If fso1.FileExists(tmpPath) Then fso1.DeleteFile tmpPath
        Set fso1 = Nothing
End Sub
