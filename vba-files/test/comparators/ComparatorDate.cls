VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparatorDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/test/comparators
Option Explicit

Implements IComparator

'/ <summary>
'/ Compares two Date values.
'/ </summary>
Private Function IComparator_Compare(ByVal a As Variant, ByVal b As Variant) As Integer
    Dim dA As Date: dA = CDate(a)
    Dim dB As Date: dB = CDate(b)
    If dA < dB Then
        IComparator_Compare = -1
    ElseIf dA > dB Then
        IComparator_Compare = 1
    Else
        IComparator_Compare = 0
    End If
End Function
