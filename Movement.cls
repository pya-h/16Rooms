VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Movement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mRow As Byte
Private mColumn As Byte

Public Property Get Row() As Variant
    Row = mRow
End Property

Public Property Let Row(ByVal val As Variant)
    If val >= 0 And val < TABLE_DIMENSION Then
        mRow = val
    End If
    
End Property

Public Sub PutOutOfTable()
    mRow = OutOfTableValue
    mColumn = OutOfTableValue
End Sub
Public Function IsInsideTable()
    IsInsideTable = mRow <> OutOfTableValue And mColumn <> OutOfTableValue
End Function
Public Sub Class_Initialize()

End Sub

Public Property Get Column() As Variant
    Column = mColumn
End Property

Public Property Let Column(ByVal val As Variant)
    If val >= 0 And val < TABLE_DIMENSION Then
        mColumn = val
    End If
    
End Property

Public Sub RandomizeMove()
    Randomize Timer
    
    ' Generate a random number between 1 and TABLE_DIMENSION - 1
    Row = Int((TABLE_DIMENSION - 1) * Rnd)
    Column = Int((TABLE_DIMENSION - 1) * Rnd)
End Sub

Public Function ToString() As String
    ToString = "(" & Row & ", " & Column & ")"
    
End Function


