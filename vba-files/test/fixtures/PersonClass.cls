VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PersonClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/test/fixtures
Option Explicit

'/ <summary>
'/ A simple test fixture class representing a person with a name, age, and original insertion index.
'/ OriginalIndex is used to verify sort stability: after sorting by a key, elements with equal keys
'/ must retain their original relative order.
'/ </summary>

Private mName          As String
Private mAge           As Long
Private mOriginalIndex As Long

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)
    mName = value
End Property

Public Property Get Age() As Long
    Age = mAge
End Property

Public Property Let Age(ByVal value As Long)
    mAge = value
End Property

Public Property Get OriginalIndex() As Long
    OriginalIndex = mOriginalIndex
End Property

Public Property Let OriginalIndex(ByVal value As Long)
    mOriginalIndex = value
End Property

Public Function ToString() As String
    ToString = "Person(Name=" & mName & ", Age=" & mAge & ", OriginalIndex=" & mOriginalIndex & ")"
End Function
