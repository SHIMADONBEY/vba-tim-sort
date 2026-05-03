VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparatorPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/test/comparators
Option Explicit

Implements IComparator

'/ <summary>
'/ Compares two PersonClass objects by Age.
'/ Used to verify sort stability: persons with equal Age must retain their original relative order after sort.
'/ </summary>
Private Function IComparator_Compare(ByVal a As Variant, ByVal b As Variant) As Integer
    Dim pA As PersonClass: Set pA = a
    Dim pB As PersonClass: Set pB = b
    If pA.Age < pB.Age Then
        IComparator_Compare = -1
    ElseIf pA.Age > pB.Age Then
        IComparator_Compare = 1
    Else
        IComparator_Compare = 0
    End If
End Function
