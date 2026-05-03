VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparatorString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/test/comparators
Option Explicit

Implements IComparator

'/ <summary>
'/ Compares two string values using case-insensitive text comparison.
'/ </summary>
Private Function IComparator_Compare(ByVal a As Variant, ByVal b As Variant) As Integer
    IComparator_Compare = StrComp(CStr(a), CStr(b), vbTextCompare)
End Function
