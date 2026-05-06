VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IComparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/<summary>
'/ Compares two objects and returns a value indicating whether one is less than, equal to, or greater than the other.
'/ </summary>
'/ <param name="a">The first object to compare.</param>
'/ <param name="b">The second object to compare.</param>
'/ <returns>
'/ A signed integer that indicates the relative values of a and b, as shown in the following table:
'/ Value Meaning
'/ Less than zero a is less than b.
'/ Zero a is equal to b.
'/ Greater than zero a is greater than b.
'/ </returns>
Public Function Compare(ByVal a As Variant, ByVal b As Variant) As Integer
End Function
