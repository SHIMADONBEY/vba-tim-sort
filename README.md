# vba-tim-sort

A library that implements Tim Sort's sorting algorithm in VBA.

English | [日本語](README.ja.md)

## How to Install

1. Importing Files
   - Open the VBA Editor, go to “File” → “Import File” in the menu, and add the following files to your project:  
     - [vba-files/VbaTimSort.bas](vba-files/VbaTimSort.bas#L1)  
     - [vba-files/IComparator.cls](vba-files/IComparator.cls#L1)  
   - Sample implementations of `IComparator` are located in `vba-files/test/comparators/`. Add them to your project as needed.

2. Usage (Arrays)

``` vb
Dim arr As Variant
arr = Array(5, 2, 9, 1)

' To sort in natural order (numbers, strings, dates), pass `Nothing` to the comparator.
Dim sorted As Variant
sorted = SortArrayInPlace(arr, Nothing)             ' Ascending

' To sort in descending order, pass True as the optional third argument.
Dim sortedDesc As Variant
sortedDesc = SortArrayInPlace(arr, Nothing, True)   ' Descending
```

3. Usage （Collection）

``` vb
Dim coll As New Collection
coll.Add "b"
coll.Add "a"
Dim sortedColl As Collection
Set sortedColl = SortCollection(coll, Nothing)
```

4. An example of implementing IComparator

``` vb
' MyComparator.cls
Implements IComparator

Private Function IComparator_Compare(ByVal a As Variant, ByVal b As Variant) As Integer
    ' Customized comparison
    If a < b Then
        IComparator_Compare = -1
    ElseIf a > b Then
        IComparator_Compare = 1
    Else
        IComparator_Compare = 0
    End If
End Function
```

``` vb
' example.bas
Dim comp As IComparator
Set comp = New MyComparator
Dim sortedWithComp As Variant
sortedWithComp = SortArrayInPlace(arr, comp)
```

5. Notes

- `SortArrayInPlace()` / `SortCollection()` does not modify the original array or collection; instead, it returns a new array or collection.
- When comparing objects, you must provide an `IComparator`. (Failure to do so will result in an error.)

## Open Source Project Guidelines

This project is maintained as an open source software (OSS) library.
The following notes describe common expectations for users and contributors.

### License

This project is distributed under the terms described in [LICENSE](LICENSE).
By using, modifying, or redistributing this project, you agree to follow that license.

### Contributing

Contributions are welcome.
Typical contribution flow:

1. Fork the repository.
2. Create a feature or fix branch.
3. Add or update tests when behavior changes.
4. Open a pull request with a clear summary.

Please keep changes focused, well-scoped, and easy to review.

### Issues and Requests

If you find a bug or want to request a feature, please open an issue.
When possible, include:

- Expected behavior
- Actual behavior
- Reproducible steps
- VBA / Excel version details

### Security

Please see [SECURITY](SECURITY.md)

### Support Scope

This project is maintained on a best-effort basis.
Response times and release schedules are not guaranteed.

### Acknowledgements

Thanks to everyone who reports issues, improves documentation, and contributes code.

