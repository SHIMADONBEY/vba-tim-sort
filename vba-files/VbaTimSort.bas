Attribute VB_Name = "VbaTimSort"
' VBA implementation of TimSort, a hybrid stable sorting algorithm derived from merge sort and insertion sort.
' TimSort was designed to perform well on many kinds of real-world data, and it is the default sorting algorithm in Python and Java.
Option Explicit

'/ <summary>
' Error codes for VbaTimSort.
' All error codes are based on vbObjectError to avoid conflicts with built-in VBA error codes.
'/ </summary>
Public Const TIMSORT_ERR_BASE As Long                               = vbObjectError

' General errors (ERR)
Public Const TIMSORT_ERR_GENERAL As Long                            = TIMSORT_ERR_BASE + 8192

' Argument-related errors (ARG)
Public Const TIMSORT_ERR_ARG_NOT_ARRAY As Long                      = TIMSORT_ERR_GENERAL + 1
Public Const TIMSORT_ERR_ARG_NOT_ONE_DIMENSIONAL_ARRAY As Long      = TIMSORT_ERR_GENERAL + 2
Public Const TIMSORT_ERR_ARG_COLLECTION_NOTHING As Long             = TIMSORT_ERR_GENERAL + 3
Public Const TIMSORT_ERR_ARG_INVALID_ITEMSCOUNT As Long             = TIMSORT_ERR_GENERAL + 4
Public Const TIMSORT_ERR_ARG_INVALID_RANGE As Long                  = TIMSORT_ERR_GENERAL + 5
Public Const TIMSORT_ERR_ARG_OUT_OF_BOUNDS As Long                  = TIMSORT_ERR_GENERAL + 6
Public Const TIMSORT_ERR_ARG_INVALID_START As Long                  = TIMSORT_ERR_GENERAL + 7
Public Const TIMSORT_ERR_ARG_LENGTH_NOT_POSITIVE As Long            = TIMSORT_ERR_GENERAL + 8
Public Const TIMSORT_ERR_ARG_INDEX_OUT_OF_RANGE As Long             = TIMSORT_ERR_GENERAL + 9
Public Const TIMSORT_ERR_ARG_RUN_LENGTHS_POSITIVE As Long           = TIMSORT_ERR_GENERAL + 10
Public Const TIMSORT_ERR_ARG_RUN_BASE_CONSISTENCY As Long           = TIMSORT_ERR_GENERAL + 11
Public Const TIMSORT_ERR_ARG_UNSUPPORTED_TYPE As Long               = TIMSORT_ERR_GENERAL + 12
Public Const TIMSORT_ERR_ARG_NO_COMPARATOR_FOR_OBJECTS As Long      = TIMSORT_ERR_GENERAL + 13

' State/stack inconsistency errors (STATE)
Public Const TIMSORT_ERR_STATE_REQUIRED_SIZE_NEGATIVE As Long       = TIMSORT_ERR_GENERAL + 65   ' requiredSize < 0
Public Const TIMSORT_ERR_STATE_RUN_STACK_MISMATCH As Long           = TIMSORT_ERR_GENERAL + 66   ' runBase/runLen init mismatch
Public Const TIMSORT_ERR_STATE_STACK_SIZE_NEGATIVE As Long          = TIMSORT_ERR_GENERAL + 67   ' stackSize < 0

' <summary>
' Upper bound used when computing TimSort's minimum run length.
' The minimum run length returned by GetMinRunLength is derived from the input size and this threshold.
' Runs shorter than the computed minimum may be extended and sorted using binary insertion sort before merging.
' Setting this value too high can lead to larger minimum runs, while setting it too low can lead to too many runs and excessive merging.
' 64 is a commonly used value that provides a good balance for many sorting tasks.
'</summary>
Private Const MAX_RUN_LENGTH As Long = 64

' <summary>
' Initial size of the run stack.
' The run stack will grow dynamically if needed, but this is the initial allocation size.
' The maximum stack size is determined by the maximum number of runs that can be created,
' which is related to the size of the input array and the minimum run length.
' In practice, this initial size should be sufficient for most sorting tasks,
' but it can be increased if you expect to sort very large arrays with many runs.
'</summary>
Private Const INITIAL_RUN_STACK_SIZE As Long = 16

' / <summary>
' / Sorts a one-dimensional array using the TimSort algorithm.
' / The original array is not modified; a new sorted array is returned.
' / </summary>
' / <param name="arr">The array to be sorted. Must be a one-dimensional array.</param>
' / <param name="comparator">
' / An optional IComparator implementation for custom object comparison.
' / This is required if the array contains objects that do not have a natural ordering or if you want to sort in a custom order.
' / If not provided, natural ordering is used for supported data types.
' / </param>
' / <param name="descending">If True, sorts in descending order. Default is False (ascending order).</param>
' / <returns>The sorted array. The original array is not modified in place.</returns>
Public Function SortArray(ByRef arr As Variant, Optional ByVal comparator As IComparator = Nothing, Optional descending As Boolean = False) As Variant
    Dim vNewArray As Variant
    If Not IsArray(arr) Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_NOT_ARRAY, "VbaTimSort.SortArray", "Input must be an array."
    ElseIf IsEmptyArray(arr) Then
        ' An unallocated array is considered empty, so we can return a new empty array.
        ReDim vNewArray(-1 To -1)
        SortArray = vNewArray
        Exit Function
    ElseIf IsMultiDimensionalArray(arr) Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_NOT_ONE_DIMENSIONAL_ARRAY, "VbaTimSort.SortArray", "Input array must be one-dimensional."
    End If

    Dim vUb As Long: vUb = UBound(arr)
    Dim vLb As Long: vLb = LBound(arr)

    ReDim vNewArray(0 To vUb - vLb)
    Dim i As Long
    For i = vLb To vUb
        If IsArray(arr(i)) Then
            ' TimSort is not designed to sort arrays that contain other arrays as elements.
            ' This is a limitation of this implementation, and we will raise an error if we encounter this case.
            Err.Raise VbaTimSort.TIMSORT_ERR_ARG_NOT_ONE_DIMENSIONAL_ARRAY, "VbaTimSort.SortArray", "Input array must be one-dimensional and cannot contain arrays as elements."
        End If

        AssignVariant vNewArray(i - vLb), arr(i)
    Next i

    SortArray = TimSortCore(vNewArray, comparator, descending)
End Function

' / <summary>
' / Sorts a Collection using the TimSort algorithm. The original Collection is not modified; a new sorted Collection is returned.
' / </summary>
' / <param name="coll">The Collection to be sorted. Cannot be Nothing.</param>
' / <param name="comparator">
' / An optional IComparator implementation for custom object comparison.
' / This is required if the Collection contains objects that do not have a natural ordering or if you want to sort in a custom order.
' / If not provided, natural ordering is used for supported data types.
' / </param>
' / <param name="descending">If True, sorts in descending order. Default is False (ascending order).</param>
' / <returns>A new Collection containing the sorted elements.</returns>
Public Function SortCollection(ByRef coll As Collection, Optional ByVal comparator As IComparator = Nothing, Optional descending As Boolean = False) As Collection
    If coll Is Nothing Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_COLLECTION_NOTHING, "VbaTimSort.SortCollection", "Input Collection cannot be Nothing."
    ElseIf coll.Count = 0 Then
        ' An empty collection is already sorted, so we can return a new empty collection.
        Set SortCollection = New Collection
        Exit Function
    End If

    Dim arr() As Variant
    Dim i As Long
    ReDim arr(1 To coll.Count)
    For i = 1 To coll.Count
        AssignVariant arr(i), coll(i)
    Next i
    Dim sortedArr As Variant
    sortedArr = TimSortCore(arr, comparator, descending)
    Dim sortedColl As New Collection
    For i = LBound(sortedArr) To UBound(sortedArr)
        sortedColl.Add sortedArr(i)
    Next i
    Set SortCollection = sortedColl

End Function

' / <summary>
' / The core TimSort algorithm that sorts an array in place. This function is called by the public sorting functions after preparing the input array.
' / </summary>
Private Function TimSortCore(ByRef arr As Variant, ByVal comparator As IComparator, ByVal descending As Boolean) As Variant
    If Not IsArray(arr) Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_NOT_ARRAY, "VbaTimSort.TimSortCore", "Input must be an array."
    End If

    Dim n As Long: n = UBound(arr) - LBound(arr) + 1
    If n <= 1 Then
        ' An array of length 0 or 1 is already sorted, so we can return it as is.
        TimSortCore = arr
        Exit Function
    End If

    TimSortCoreMainLoop arr, LBound(arr), UBound(arr), comparator, descending
    TimSortCore = arr
End Function

'/ <summary>
'/ The main loop of the TimSort algorithm. It identifies runs, sorts them if necessary, and merges them together.
'/ </summary>
'/ <param name="arr">The array to be sorted. This array is modified in place.</param>
'/ <param name="lo">The lower index of the portion of the array to be sorted.</param>
'/ <param name="hi">The upper index of the portion of the array to be sorted.</param>
'/ <param name="comparator">An optional IComparator implementation for custom object comparison.</param>
'/ <param name="descending">If True, sorts in descending order. Default is False (ascending order).</param>
Private Sub TimSortCoreMainLoop( _
    ByRef arr As Variant, _
    ByVal lo As Long, _
    ByVal hi As Long, _
    ByVal comparator As IComparator, _
    ByVal descending As Boolean _
)
    Dim vMinRun As Long: vMinRun = GetMinRunLength(hi - lo + 1)
    Dim vRunBase() As Long
    Dim vRunLen() As Long

    Dim vStackSize As Long: vStackSize = 0
    Dim vIndex As Long: vIndex = lo
    Do While vIndex <= hi
        Dim vRunLength As Long: vRunLength = CountRunAndMakeAscending(arr, vIndex, hi, comparator, descending)

        If vRunLength < vMinRun Then
            Dim vForceLen As Long: vForceLen = IIf(vMinRun < hi - vIndex + 1, vMinRun, hi - vIndex + 1)
            BinaryInsertionSort arr, vIndex, vIndex + vForceLen - 1, vIndex + vRunLength, comparator, descending
            vRunLength = vForceLen
        End If

        PushRun vRunBase, vRunLen, vStackSize, vIndex, vRunLength
        MergeCollapse arr, vRunBase, vRunLen, vStackSize, comparator, descending

        vIndex = vIndex + vRunLength
    Loop

    MergeForceCollapse arr, vRunBase, vRunLen, vStackSize, comparator, descending
End Sub

'/ <summary>
'/ Compares two elements using the provided comparator or natural ordering. Returns a negative number if a < b, zero if a = b, and a positive number if a > b.
'/ The comparison is adjusted for ascending or descending order based on the 'descending' parameter.
'/ </summary>
'/ <param name="a">The first element to compare.</param>
'/ <param name="b">The second element to compare.</param>
'/ <param name="comparator">An optional IComparator implementation for custom object comparison. If not provided, natural ordering is used.</param>
'/ <param name="descending">If True, the comparison is reversed for descending order. Default is False (ascending order).</param>
'/ <returns>A negative number if a < b, zero if a = b, and a positive number if a > b, adjusted for ascending or descending order.</returns>
Private Function InternalCompare(a As Variant, b As Variant, comparator As IComparator, descending As Boolean) As Long
    Dim vCompared As Long

    If (IsObject(a) Or IsObject(b)) Then
        If comparator Is Nothing Then
            Err.Raise VbaTimSort.TIMSORT_ERR_ARG_NO_COMPARATOR_FOR_OBJECTS, "VbaTimSort.InternalCompare", "Object comparison requires IComparator."
        End If
        vCompared = comparator.Compare(a, b)
    ElseIf (VarType(a) = vbString And VarType(b) = vbString) Then
        vCompared = StrComp(a, b, vbTextCompare)
    ElseIf (VarType(a) = vbDate And VarType(b) = vbDate) Then
        If a < b Then
            vCompared = -1
        ElseIf a > b Then
            vCompared = 1
        Else
            vCompared = 0
        End If
    ElseIf (IsNumeric(a) And IsNumeric(b)) Then
        If a < b Then
            vCompared = -1
        ElseIf a > b Then
            vCompared = 1
        Else
            vCompared = 0
        End If
    Else
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_UNSUPPORTED_TYPE, "VbaTimSort.InternalCompare", "Unsupported data types for comparison."
    End If

    InternalCompare = Sgn(vCompared * IIf(descending, -1, 1))
End Function

'/ <summary>
'/ Calculates the minimum run length for a given number of items.
'/ This is used to determine how to break the array into runs for the TimSort algorithm.
'/ </summary>
'/ <param name="ItemsCount">The total number of items to be sorted. Must be a positive integer.</param>
'/ <returns>The minimum run length to be used in the TimSort algorithm.</returns>
Private Function GetMinRunLength(ByVal ItemsCount As Long) As Long
    If ItemsCount <= 0 Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_LENGTH_NOT_POSITIVE, "VbaTimSort.GetMinRunLength", "ItemsCount must be positive."
    ElseIf ItemsCount < MAX_RUN_LENGTH Then
        GetMinRunLength = ItemsCount
        Exit Function
    End If

    Dim n As Long: n = ItemsCount
    Dim r As Long: r = 0
    Do While n >= MAX_RUN_LENGTH
        r = r Or (n And 1)
        n = n \ 2
    Loop

    GetMinRunLength = n + r
End Function

'/ <summary>
'/ Identifies a run in the array starting at index 'lo' and makes it ascending if it is not already. Returns the length of the run.
'/ </summary>
'/ <param name="arr">The array containing the run. This array is modified in place if the run is not ascending.</param>
'/ <param name="lo">The starting index of the run.</param>
'/ <param name="hi">The upper index of the array. The run cannot extend beyond this index.</param>
'/ <param name="comparator">An optional IComparator implementation for custom object comparison.</param>
'/ <param name="descending">If True, the run is considered descending and will be reversed to become ascending. Default is False (ascending order).</param>
'/ <returns>The length of the run starting at index 'lo'. The run is guaranteed to be ascending after this function returns.</returns>
Private Function CountRunAndMakeAscending( _
    ByRef arr As Variant, _
    ByVal lo As Long, _
    ByVal hi As Long, _
    ByVal comparator As IComparator, _
    ByVal descending As Boolean _
) As Long
    If lo > hi Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_INVALID_RANGE, "VbaTimSort.CountRunAndMakeAscending", "lo must be less than or equal to hi."
    ElseIf Not IsArray(arr) Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_NOT_ARRAY, "VbaTimSort.CountRunAndMakeAscending", "Input must be an array."
    ElseIf LBound(arr) > hi Or UBound(arr) < lo Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_OUT_OF_BOUNDS, "VbaTimSort.CountRunAndMakeAscending", "lo and hi must be within the bounds of the array."
    End If

    If (lo = hi) Then
        ' A single-element run is already considered ascending, so the run length is 1
        CountRunAndMakeAscending = 1
        Exit Function
    End If

    Dim vDirection As Long: vDirection = IIf(InternalCompare(arr(lo), arr(lo + 1), comparator, descending) <= 0, -1, 1)
    Dim vRunLength As Long: vRunLength = 2

    Do While (lo + vRunLength <= hi)
        If vDirection < 0 Then
            ' Ascending run
            If InternalCompare(arr(lo + vRunLength - 1), arr(lo + vRunLength), comparator, descending) > 0 Then
                Exit Do
            End If
        Else
            ' Descending run (treat equal values as ascending)
            If InternalCompare(arr(lo + vRunLength - 1), arr(lo + vRunLength), comparator, descending) <= 0 Then
                Exit Do
            End If
        End If
        vRunLength = vRunLength + 1
    Loop

    ' If the run is descending, reverse it to make it ascending
    If vDirection > 0 Then
        ReverseRange arr, lo, lo + vRunLength - 1
    End If

    CountRunAndMakeAscending = vRunLength
End Function

'/ <summary>
'/ Reverses a portion of the array in place from index 'lo' to index 'hi'. This is used to convert descending runs into ascending runs in the TimSort algorithm.
'/ </summary>
'/ <param name="arr">The array containing the range to reverse. This array is modified in place.</param>
'/ <param name="lo">The lower index of the range to reverse.</param>
'/ <param name="hi">The upper index of the range to reverse. Must be greater than or equal to 'lo'.</param>
Private Sub ReverseRange(ByRef arr As Variant, ByVal lo As Long, ByVal hi As Long)
    If lo > hi Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_INVALID_RANGE, "VbaTimSort.ReverseRange", "lo must be less than or equal to hi."
    ElseIf Not IsArray(arr) Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_NOT_ARRAY, "VbaTimSort.ReverseRange", "Input must be an array."
    ElseIf LBound(arr) > hi Or UBound(arr) < lo Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_OUT_OF_BOUNDS, "VbaTimSort.ReverseRange", "lo and hi must be within the bounds of the array."
    End If

    Dim i As Long: i = lo
    Dim j As Long: j = hi
    Do While i < j
        Dim temp As Variant
        AssignVariant temp, arr(i)
        AssignVariant arr(i), arr(j)
        AssignVariant arr(j), temp
        i = i + 1
        j = j - 1
    Loop
End Sub

'/ <summary>
'/ Sorts a portion of the array using binary insertion sort. This is used to sort small runs in the TimSort algorithm.
'/ </summary>
'/ <param name="arr">The array to be sorted. This array is modified in place.</param>
'/ <param name="lo">The lower index of the portion of the array to be sorted.</param>
'/ <param name="hi">The upper index of the portion of the array to be sorted.</param>
'/ <param name="start">The index at which to start the insertion sort. Must be between 'lo' and 'hi + 1'.</param>
'/ <param name="comparator">An optional IComparator implementation for custom object comparison.</param>
'/ <param name="descending">If True, sorts in descending order. Default is False (ascending order).</param>
Private Sub BinaryInsertionSort( _
    ByRef arr As Variant, _
    ByVal lo As Long, _
    ByVal hi As Long, _
    ByVal start As Long, _
    ByVal comparator As IComparator, _
    ByVal descending As Boolean _
)
    If IsArray(arr) = False Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_NOT_ARRAY, "VbaTimSort.BinaryInsertionSort", "Input must be an array."
    ElseIf lo < LBound(arr) Or hi > UBound(arr) Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_OUT_OF_BOUNDS, "VbaTimSort.BinaryInsertionSort", "lo and hi must be within the bounds of the array."
    ElseIf start < lo Or start > hi + 1 Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_INVALID_RANGE, "VbaTimSort.BinaryInsertionSort", "start must be greater than or equal to lo and less than or equal to hi + 1."
    End If

    Dim i As Long
    For i = start To hi
        Dim vPivot As Variant
        AssignVariant vPivot, arr(i)
        Dim vLeft As Long: vLeft = lo
        Dim vRight As Long: vRight = i

        ' Find the position where vPivot should be inserted in the sorted subarray arr(lo...i-1) using binary search.
        Do While vLeft < vRight
            Dim vMiddle As Long: vMiddle = (vLeft + vRight) \ 2
            If InternalCompare(vPivot, arr(vMiddle), comparator, descending) < 0 Then
                vRight = vMiddle
            Else
                vLeft = vMiddle + 1
            End If
        Loop

        ' Shift elements to make room for the pivot.
        Dim j As Long
        For j = i To vLeft + 1 Step -1
            AssignVariant arr(j), arr(j - 1)
        Next j
        AssignVariant arr(vLeft), vPivot
    Next i

End Sub

'/ <summary>
'/ Pushes a run onto the stack of runs to be merged. This function is called when a new run is identified in the main loop of the TimSort algorithm.
'/ </summary>
'/ <param name="runBase">An array storing the starting indices of runs.</param>
'/ <param name="runLen">An array storing the lengths of runs.</param>
'/ <param name="stackSize">The current size of the run stack.</param>
'/ <param name="base">The starting index of the new run.</param>
'/ <param name="length">The length of the new run.</param>
Private Sub PushRun( _
    ByRef runBase() As Long, _
    ByRef runLen() As Long, _
    ByRef stackSize As Long, _
    ByVal base As Long, _
    ByVal length As Long _
)
    If base < 0 Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_OUT_OF_BOUNDS, "VbaTimSort.PushRun", "base must be non-negative."
    ElseIf length <= 0 Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_LENGTH_NOT_POSITIVE, "VbaTimSort.PushRun", "length must be positive."
    End If

    EnsureStackCapacity runBase, runLen, stackSize + 1
    runBase(stackSize) = base
    runLen(stackSize) = length
    stackSize = stackSize + 1
End Sub

'/ <summary>
'/ Merges runs on the stack until the invariants of the TimSort algorithm are maintained.
'/ This function is called after pushing a new run onto the stack to ensure that the runs are merged in the correct order.
'/ </summary>
'/ <param name="arr">The array being sorted.</param>
'/ <param name="runBase">An array storing the starting indices of runs.</param>
'/ <param name="runLen">An array storing the lengths of runs.</param>
'/ <param name="stackSize">The current size of the run stack.</param>
'/ <param name="comparator">An object implementing the IComparator interface for custom comparison.</param>
'/ <param name="descending">If True, sorts in descending order. Default is False (ascending order).</param>
Private Sub MergeCollapse( _
    ByRef arr As Variant, _
    ByRef runBase() As Long, _
    ByRef runLen() As Long, _
    ByRef stackSize As Long, _
    ByVal comparator As IComparator, _
    ByVal descending As Boolean _
)
    If IsArray(arr) = False Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_NOT_ARRAY, "VbaTimSort.MergeCollapse", "Input must be an array."
    ElseIf stackSize < 0 Then
        Err.Raise VbaTimSort.TIMSORT_ERR_STATE_STACK_SIZE_NEGATIVE, "VbaTimSort.MergeCollapse", "stackSize must be non-negative."
    End If

    If stackSize < 2 Then
        ' no need to merge if there is only one run
        Exit Sub
    End If

    Dim vCursor As Long: vCursor = stackSize - 2
    Do While stackSize > 1
        Dim vMerged As Boolean: vMerged = False
        Dim vLenB As Long: vLenB = runLen(vCursor)
        Dim vLenC As Long: vLenC = runLen(vCursor + 1)

        If stackSize > 2 Then
            Dim vLenA As Long: vLenA = runLen(vCursor - 1)
            Dim vCheckLeft As Boolean: vCheckLeft = vCursor > 1
            Dim vLenD As Long
            If vCheckLeft Then
                vLenD = runLen(vCursor - 2)
            End If

            If (vLenA <= vLenB + vLenC) Or (vCheckLeft And vLenD <= vLenA + vLenB) Then
                If vLenA < vLenC Then
                    MergeAt arr, runBase, runLen, stackSize, vCursor - 1, comparator, descending
                    vMerged = True
                Else
                    MergeAt arr, runBase, runLen, stackSize, vCursor, comparator, descending
                    vMerged = True
                End If
            ElseIf vLenB <= vLenC Then
                MergeAt arr, runBase, runLen, stackSize, vCursor, comparator, descending
                vMerged = True
            End If
        Else
            If vLenB <= vLenC Then
                MergeAt arr, runBase, runLen, stackSize, vCursor, comparator, descending
                vMerged = True
            End If
        End If

        If vMerged Then vCursor = stackSize - 2 Else Exit Do
    Loop

End Sub

'/ <summary>
'/ Merges all runs on the stack until only one run remains. This is called at the end of the main loop of the TimSort algorithm to ensure that all runs are merged together.
'/ </summary>
'/ <param name="arr">The array being sorted.</param>
'/ <param name="runBase">An array storing the starting indices of runs.</param>
'/ <param name="runLen">An array storing the lengths of runs.</param>
'/ <param name="stackSize">The current size of the run stack.</param>
'/ <param name="comparator">An object implementing the IComparator interface for custom comparison.</param>
'/ <param name="descending">If True, sorts in descending order. Default is False (ascending order).</param>
Private Sub MergeForceCollapse( _
    ByRef arr As Variant, _
    ByRef runBase() As Long, _
    ByRef runLen() As Long, _
    ByRef stackSize As Long, _
    ByVal comparator As IComparator, _
    ByVal descending As Boolean _
)
    If IsArray(arr) = False Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_NOT_ARRAY, "VbaTimSort.MergeForceCollapse", "Input must be an array."
    ElseIf stackSize < 0 Then
        Err.Raise VbaTimSort.TIMSORT_ERR_STATE_STACK_SIZE_NEGATIVE, "VbaTimSort.MergeForceCollapse", "stackSize must be non-negative."
    End If

    Do While stackSize > 1
        Dim vCursor As Long: vCursor = stackSize - 2
        If vCursor > 0 And runLen(vCursor - 1) < runLen(vCursor + 1) Then
            vCursor = vCursor - 1
        End If

        MergeAt arr, runBase, runLen, stackSize, vCursor, comparator, descending
    Loop
End Sub

'/ <summary>
'/ Merges the runs at index 'i' and 'i + 1' on the stack. This function is called by both MergeCollapse and MergeForceCollapse to merge runs together.
'/ </summary>
'/ <param name="arr">The array being sorted.</param>
'/ <param name="runBase">An array storing the starting indices of runs.</param>
'/ <param name="runLen">An array storing the lengths of runs.</param>
'/ <param name="stackSize">The current size of the run stack.</param>
'/ <param name="i">The index of the first run to merge. Must be between 0 and stackSize - 2 inclusive.</param>
'/ <param name="comparator">An object implementing the IComparator interface for custom comparison.</param>
'/ <param name="descending">If True, sorts in descending order. Default is False (ascending order).</param>
Private Sub MergeAt( _
    ByRef arr As Variant, _
    ByRef runBase() As Long, _
    ByRef runLen() As Long, _
    ByRef stackSize As Long, _
    ByVal i As Long, _
    ByVal comparator As IComparator, _
    ByVal descending As Boolean _
)
    If i < 0 Or i >= stackSize - 1 Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_OUT_OF_BOUNDS, "VbaTimSort.MergeAt", "i must be between 0 and stackSize - 2 inclusive."
    End If

    Dim vBase1 As Long: vBase1 = runBase(i)
    Dim vLen1 As Long: vLen1 = runLen(i)
    Dim vBase2 As Long: vBase2 = runBase(i + 1)
    Dim vLen2 As Long: vLen2 = runLen(i + 1)

    If vLen1 <= 0 Or vLen2 <= 0 Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_RUN_LENGTHS_POSITIVE, "VbaTimSort.MergeAt", "run lengths must be positive."
    ElseIf (vBase1 + vLen1 <> vBase2) Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_RUN_BASE_CONSISTENCY, "VbaTimSort.MergeAt", "runBase[i] + runLen[i] must be equal to runBase[i + 1]."
    End If

    If (vLen1 <= vLen2) Then
        MergeLow arr, vBase1, vLen1, vBase2, vLen2, comparator, descending
    Else
        MergeHigh arr, vBase1, vLen1, vBase2, vLen2, comparator, descending
    End If

    runLen(i) = vLen1 + vLen2
    Dim j As Long
    For j = i + 1 To stackSize - 2
        ' Shift runBase and runLen down by one position
        runBase(j) = runBase(j + 1)
        runLen(j) = runLen(j + 1)
    Next j

    stackSize = stackSize - 1
End Sub

'/ <summary>
'/ Merges two adjacent runs where the first run is shorter than or equal to the second run. This function is optimized for this case to minimize the number of comparisons and movements.
'/ </summary>
'/ <param name="arr">The array being sorted.</param>
'/ <param name="base1">The starting index of the first run.</param>
'/ <param name="len1">The length of the first run.</param>
'/ <param name="base2">The starting index of the second run. Must be equal to base1 + len1.</param>
'/ <param name="len2">The length of the second run.</param>
'/ <param name="comparator">An object implementing the IComparator interface for custom comparison.</param>
'/ <param name="descending">If True, sorts in descending order. Default is False (ascending order).</param>
Private Sub MergeLow( _
    ByRef arr As Variant, _
    ByVal base1 As Long, _
    ByVal len1 As Long, _
    ByVal base2 As Long, _
    ByVal len2 As Long, _
    ByVal comparator As IComparator, _
    ByVal descending As Boolean _
)
    If len1 <= 0 Or len2 <= 0 Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_RUN_LENGTHS_POSITIVE, "VbaTimSort.MergeLow", "run lengths must be positive."
    ElseIf len1 > len2 Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_INVALID_RANGE, "VbaTimSort.MergeLow", "len1 must be less than or equal to len2."
    ElseIf (base1 + len1 <> base2) Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_RUN_BASE_CONSISTENCY, "VbaTimSort.MergeLow", "base1 + len1 must be equal to base2."
    End If

    Dim vTemp() As Variant
    ReDim vTemp(0 To len1 - 1)

    Dim i As Long
    For i = 0 To len1 - 1
        AssignVariant vTemp(i), arr(base1 + i)
    Next i

    i = 0
    Dim j As Long: j = base2
    Dim k As Long: k = base1

    Do While i < len1 And j < base2 + len2
        If InternalCompare(vTemp(i), arr(j), comparator, descending) <= 0 Then
            AssignVariant arr(k), vTemp(i)
            i = i + 1
        Else
            AssignVariant arr(k), arr(j)
            j = j + 1
        End If
        k = k + 1
    Loop

    ' Copy remaining elements of vTemp if any
    Do While i < len1
        AssignVariant arr(k), vTemp(i)
        i = i + 1
        k = k + 1
    Loop

    ' No need to copy the second run (arr) because they are already in place
End Sub

'/ <summary>
'/ Merges two adjacent runs where the first run is longer than or equal to the second run. This function is optimized for this case to minimize the number of comparisons and movements.
'/ </summary>
'/ <param name="arr">The array being sorted.</param>
'/ <param name="base1">The starting index of the first run.</param>
'/ <param name="len1">The length of the first run.</param>
'/ <param name="base2">The starting index of the second run. Must be equal to base1 + len1.</param>
'/ <param name="len2">The length of the second run.</param>
'/ <param name="comparator">An object implementing the IComparator interface for custom comparison.</param>
'/ <param name="descending">If True, sorts in descending order. Default is False (ascending order).</param>
Private Sub MergeHigh( _
    ByRef arr As Variant, _
    ByVal base1 As Long, _
    ByVal len1 As Long, _
    ByVal base2 As Long, _
    ByVal len2 As Long, _
    ByVal comparator As IComparator, _
    ByVal descending As Boolean _
)
    If len1 <= 0 Or len2 <= 0 Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_RUN_LENGTHS_POSITIVE, "VbaTimSort.MergeHigh", "run lengths must be positive."
    ElseIf len1 < len2 Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_INVALID_RANGE, "VbaTimSort.MergeHigh", "len1 must be greater than or equal to len2."
    ElseIf (base1 + len1 <> base2) Then
        Err.Raise VbaTimSort.TIMSORT_ERR_ARG_RUN_BASE_CONSISTENCY, "VbaTimSort.MergeHigh", "base1 + len1 must be equal to base2."
    End If

    Dim vTemp() As Variant
    ReDim vTemp(0 To len2 - 1)

    Dim j As Long
    For j = 0 To len2 - 1
        AssignVariant vTemp(j), arr(base2 + j)
    Next j

    Dim i As Long: i = base1 + len1 - 1
    j = len2 - 1
    Dim k As Long: k = base2 + len2 - 1
    Do While i >= base1 And j >= 0
        If InternalCompare(arr(i), vTemp(j), comparator, descending) > 0 Then
            AssignVariant arr(k), arr(i)
            i = i - 1
        Else
            AssignVariant arr(k), vTemp(j)
            j = j - 1
        End If
        k = k - 1
    Loop

    ' Copy remaining elements of vTemp if any
    Do While j >= 0
        AssignVariant arr(k), vTemp(j)
        j = j - 1
        k = k - 1
    Loop

    ' No need to copy the first run (arr) because they are already in place
End Sub

'/ <summary>
'/ Ensures that the run stack has enough capacity to hold the required number of runs. If the stack is not large enough, it is resized to accommodate the required size.
'/ </summary>
'/ <param name="runBase">An array storing the starting indices of runs. This array may be resized if it does not have enough capacity.</param>
'/ <param name="runLen">An array storing the lengths of runs. This array may be resized if it does not have enough capacity. Must be resized in sync with runBase.</param>
'/ <param name="requiredSize">
'/ The required size of the run stack. Must be non-negative.
'/ If the current capacity of the stack is less than requiredSize, the stack will be resized to at least requiredSize.
'/ </param>
Private Sub EnsureStackCapacity( _
    ByRef runBase() As Long, _
    ByRef runLen() As Long, _
    ByVal requiredSize As Long _
)
    Dim vHasRunBase As Boolean: vHasRunBase = HasArrayAllocated(runBase)
    Dim vHasRunLen As Boolean: vHasRunLen = HasArrayAllocated(runLen)

    If requiredSize < 0 Then
        Err.Raise VbaTimSort.TIMSORT_ERR_STATE_REQUIRED_SIZE_NEGATIVE, "VbaTimSort.EnsureStackCapacity", "requiredSize must be non-negative."
    ElseIf vHasRunBase XOr vHasRunLen Then
        Err.Raise VbaTimSort.TIMSORT_ERR_STATE_RUN_STACK_MISMATCH, "VbaTimSort.EnsureStackCapacity", "runBase and runLen must be both initialized or both uninitialized."
    ElseIf requiredSize = 0 Then
        ' no required size, no need to allocate
        Exit Sub
    End If

    If Not (vHasRunBase And vHasRunLen) Then
        Dim vInitialSize As Long: vInitialSize = IIf(INITIAL_RUN_STACK_SIZE > requiredSize, INITIAL_RUN_STACK_SIZE, requiredSize)
        ReDim runBase(0 To vInitialSize - 1)
        ReDim runLen(0 To vInitialSize - 1)
        Exit Sub
    End If

    If UBound(runBase) < requiredSize - 1 Then
        Dim newSize As Long: newSize = IIf(UBound(runBase) = -1, 4, UBound(runBase) * 2 + 1)
        If newSize < requiredSize Then
            newSize = requiredSize
        End If
        ReDim Preserve runBase(0 To newSize - 1)
        ReDim Preserve runLen(0 To newSize - 1)
    End If
End Sub

'/ <summary>
'/ Assigns a value to a Variant variable, handling both object references and value types correctly. This is used to ensure that when we copy elements of the array, we maintain the correct semantics for objects and values.
'/ </summary>
'/ <param name="target">The Variant variable to which the value will be assigned. This variable will be modified by this function.</param>
'/ <param name="source">The value to be assigned to the target variable. This can be an object reference or a value type.</param>
Private Sub AssignVariant(ByRef target As Variant, ByVal source As Variant)
    If IsObject(source) Then
        Set target = source
    Else
        target = source
    End If
End Sub

' Checks whether an array should be treated as empty in this library.
' This includes unallocated arrays and the sentinel bounds (-1 To -1).
Private Function IsEmptyArray(ByRef arr As Variant) As Boolean
    If Not HasArrayAllocated(arr) Then
        IsEmptyArray = True
        Exit Function
    End If

    Dim vLb As Long: vLb = LBound(arr)
    Dim vUb As Long: vUb = UBound(arr)
    IsEmptyArray = (vUb < vLb) Or (vLb = -1 And vUb = -1)
End Function

' / <summary>
' / Checks whether an array is a multi-dimensional array. This is used to ensure that the sorting functions only operate on one-dimensional arrays, as multi-dimensional arrays are not supported.
' / </summary>
' / <param name="arr">The array to check. This should be a Variant variable that may or may not contain an array.</param>
' / <returns>True if the array is a multi-dimensional array, False if it is a one-dimensional array or not an array at all.</returns>
Private Function IsMultiDimensionalArray(ByRef arr As Variant) As Boolean
    On Error GoTo NotMultiDim
    Dim vTemp As Long
    vTemp = LBound(arr, 2)
    IsMultiDimensionalArray = True
    Goto Finally_IsMultiDimensionalArray
NotMultiDim:
    IsMultiDimensionalArray = False
Finally_IsMultiDimensionalArray:
    Err.Clear
End Function

'/ <summary>
'/ Checks if a dynamic array has been allocated. This is used to determine whether the run stack arrays have been initialized before trying to use them.
'/ </summary>
'/ <param name="arr">The array to check. This should be a dynamic array variable that may or may not have been allocated.</param>
'/ <returns>True if the array has been allocated and can be accessed without error, False if the array has not been allocated.</returns>
Private Function HasArrayAllocated(arr As Variant) As Boolean
    On Error GoTo NotAllocated
    Dim vTemp As Long
    vTemp = LBound(arr)
    HasArrayAllocated = True
    Goto Finally_HasArrayAllocated
NotAllocated:
    HasArrayAllocated = False
Finally_HasArrayAllocated:
    Err.Clear
End Function
