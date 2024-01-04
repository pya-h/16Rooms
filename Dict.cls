VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Dict Class Module
Private keys() As Variant
Private values() As Variant

' DictionaryObject Class Module
Public Sub Add(key As Variant, value As Variant)
    Dim index As Integer
    index = GetIndex(key)
    
    If index = -1 Then
        index = UBound(keys)
        ReDim Preserve keys(UBound(keys) + 1)
        ReDim Preserve values(UBound(values) + 1)
        
    End If

    keys(index) = key
    values(index) = value
End Sub

Public Function Item(key As Variant) As Variant
    Dim index As Integer
    index = GetIndex(key)
    If index <> -1 Then
        Item = values(index)
    End If
End Function

Public Sub Del(key As Variant)
    Dim index As Integer
    index = GetIndex(key)

    If index <> -1 Then
        ' Shift elements to remove the item
        For i = index To UBound(keys) - 1
            keys(i) = keys(i + 1)
            values(i) = values(i + 1)
        Next i

        ' Resize arrays
        ReDim Preserve keys(UBound(keys) - 1)
        ReDim Preserve values(UBound(values) - 1)
    End If
End Sub

Public Sub Class_Initialize()
    ReDim Preserve keys(0)
    ReDim Preserve values(0)
End Sub

Private Function GetIndex(key As Variant) As Integer
    Dim i As Integer
    For i = LBound(keys) To UBound(keys)
        If keys(i) = key Then
            GetIndex = i
            Exit Function
        End If
    Next i

    ' Return -1 if key not found
    GetIndex = -1
End Function

