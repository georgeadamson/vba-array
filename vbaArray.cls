VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBAArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit

'This class provides a zero-based array-like object with methods for use in VBA.
'
'TODO:
'Enable enumeration of this vbaArray object


'Properties:
' - Count
' - Values
'
'Methods:
' - First
' - Last
' - LastIndex
' - Push
' - Pop
' - Shift
' - Unshift

'Internally we store the values in a native vba array:
Private m_items() As Variant

'Return a native array of all the items:
Public Property Get Item(index) As Variant
    AssignValue Item, m_items(index)
End Property

'Return a native array of all the items:
Public Property Get Items() As Variant
    Items = m_items
End Property

'Re-populate our array object with the new values from a native array:
Public Property Let Items(value As Variant)

    Erase m_items
    Merge value

End Property

'Return the UBound of the array:
Public Property Get Count() As Long

    Count = 0
    On Error Resume Next
    Count = UBound(m_items) + 1

End Property

'Length is just an alias for Count:
Public Property Get Length() As Long
    
    Length = Me.Count

End Property

'Return the array's first item:
Public Function First() As Variant
    
    If Me.Count > 0 Then AssignValue First, m_items(0) Else First = Null

End Function

'Return the array's last item:
Public Function Last() As Variant
    
    If Me.Count > 0 Then AssignValue Last, m_items(Me.LastIndex) Else Last = Null

End Function

'Return the index of the array's last item: (Or -1 when array is empty)
Public Function LastIndex() As Long
    
    LastIndex = Me.Count - 1

End Function

'Append the array with a new value:
Public Function Push(value As Variant) As Variant

    Dim i As Long: i = Me.Count
    ReDim Preserve m_items(i) As Variant

    AssignValue m_items(i), value
    AssignValue Push, Me
    
End Function

'Remove and return the last item from the array:
Public Function Pop() As Variant

    AssignValue Pop, Me.Last
    If Me.Count > 1 Then
        ReDim Preserve m_items(Me.LastIndex - 1) As Variant
    Else
        Erase m_items
    End If
    
End Function

'Prepend the array with a new value:
Public Function Unshift(value As Variant) As Variant

    Dim i As Long
    ReDim Preserve m_items(Me.Count) As Variant
    For i = Me.LastIndex To 1 Step -1
        AssignValue m_items(i), m_items(i - 1)
    Next
    AssignValue m_items(0), value
    AssignValue Unshift, Me
    
End Function

'Remove and return the first item from the array:
Public Function Shift() As Variant

    Dim i As Long
    AssignValue Shift, Me.First
    If Me.Count > 1 Then
        For i = 0 To Me.LastIndex - 1
            AssignValue m_items(i), m_items(i + 1)
        Next
        ReDim Preserve m_items(Me.LastIndex - 1) As Variant
    Else
        Erase m_items
    End If
    
End Function

'Append our array object with the new values:
Public Function Merge(newArray As Variant) As VBAArray

    Dim Item As Variant
    
    For Each Item In makeArray(newArray)
        Me.Push Item
    Next
    
    Set Merge = Me
    
End Function

'Concat is really just an alias for Merge:
Public Function Concat(newArray As Variant) As VBAArray
    
    Set Concat = Merge(newArray)

End Function

'Return a string created by joining all the array values:
Public Function Join(Optional delimiter = " ") As String

    Dim i As Long
    Dim Item As Variant
    Dim tmpArray() As String
    ReDim tmpArray(Me.LastIndex) As String

    For Each Item In m_items
        tmpArray(i) = ItemToString(Item, delimiter)
        i = i + 1
    Next

    Join = VBA.Join(tmpArray, delimiter)

End Function

'Populate array from the substrings in the specified expression:
Public Function Split(expression As String, _
                    Optional delimiter = " ", _
                    Optional limit As Long = -1, _
                    Optional compare As VbCompareMethod = vbBinaryCompare) _
                    As VBAArray
    
    Dim tmpArray As Variant
    Items = VBA.Split(expression, delimiter, limit, compare)
    Set Split = Me

End Function





'Helper for converting an array or value to a native VB variant array:
Private Function makeArray(value As Variant) As Variant

    Dim tmpArray() As Variant

    If IsArray(value) Then

        'Return value as-is if it is already a Variant array: (VarType 8204)
        If VarType(value) = VarType(m_items) Then
            makeArray = value
        Else
            Dim i As Long
            Dim Item As Variant
            ReDim tmpArray(UBound(value)) As Variant
            For Each Item In value
                tmpArray(i) = Item
                i = i + 1
            Next
            makeArray = tmpArray
        End If

    ElseIf IsVbaArray(value) Then
        
        'Return the vbaArray items array as-is:
        makeArray = value.Items
        
    Else
    
        'Copy value into a new 1-item array:
        ReDim tmpArray(0) As Variant
        AssignValue tmpArray(0), value
        makeArray = tmpArray

    End If

End Function

'Helper for detecting whether arg is one of our custom array objects:
Private Function IsVbaArray(value As Variant) As Boolean
    
    IsVbaArray = (LCase(TypeName(value)) = "vbaarray")

End Function

'Helper to return a string representation of just about anything!
Private Function ItemToString(Item, Optional delimiter = " ") As String

    If IsMissing(Item) Then
        ItemToString = "[Missing]"
    ElseIf IsObject(Item) Then
        If Item Is Nothing Then
            ItemToString = "[Nothing]"
        ElseIf Len(TypeName(Item)) > 0 Then
            ItemToString = "[" & TypeName(Item) & "]"
        Else
            ItemToString = "[Object]"
        End If
    ElseIf IsNull(Item) Then
        ItemToString = "[Null]"
    ElseIf IsArray(Item) Then
        ItemToString = VBA.Join(Item, delimiter)
    Else
        ItemToString = Item
    End If
    
End Function

'Internal helper for setting a variable with a new value, irrespective of it's type:
Private Sub AssignValue(ByRef variable As Variant, ByRef newValue As Variant)

    If IsObject(newValue) Then
        Set variable = newValue
    Else
        variable = newValue
    End If
    
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
   'Set NewEnum = m_items.[_NewEnum]
End Property

Private Sub Class_Initialize()
    'Start with an erased array because it's easier to handle:
    Erase m_items
End Sub

'Public Property Get NewEnum() As IUnknown
'Attribute NewEnum.VB_UserMemId = -4
'   Set NewEnum = m_items.[_NewEnum]
'End Property

