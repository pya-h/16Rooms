VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Opponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mLevel As Byte
Private mName As Byte
Private mUnusedPieces() As Byte
Private mPreMove As New Movement, mNewMove As New Movement
Public Property Get Name() As Variant
    Name = mName
End Property

Public Property Let Name(ByVal val As Variant)
    If val = BodY Or val = HUMAN Then
        mName = val
    Else
        mName = WRONG
    End If
    
End Property

Public Sub Class_Initialize()
    Name = Game.BodY
    Level = Game.HARD
    ReDim mUnusedPieces(0 To (TABLE_DIMENSION - 1))
    Set mPreMove = New Movement
    Set mNewMove = New Movement
End Sub

Public Sub ResetPiecesToUnused(first As Byte)
    Dim i As Byte
    For i = 0 To UBound(mUnusedPieces)
        mUnusedPieces(i) = first + i
    Next
End Sub
Public Property Get Level() As Variant
    Level = mLevel
End Property

Public Property Let Level(ByVal val As Variant)
    If val = Game.EASY Or val = Game.HARD Then
        mLevel = val
    Else
        mLevel = WRONG
    End If
End Property

Public Function ToString() As String
    ToString = "Opponent: " & IIf(Name = Game.BodY, "BodY", "Human")
End Function

Public Property Get UnusedPieces()
    Dim length As Integer
    length = UBound(mUnusedPieces)
    If length > 0 Then
        UnusedPieces = mUnusedPieces(length)
        ReDim Preserve mUnusedPieces(0 To length - 1)
    Else
        UnusedPieces = NO_UNUSED_PIECES
    End If
End Property

Public Property Get PreMove() As Movement
    Set PreMove = mPreMove
End Property
Public Property Get NewMove() As Movement
    Set NewMove = mNewMove
End Property

Public Property Let PreMove(ByRef value As Movement)
    Set mPreMove = value
End Property
Public Property Let NewMove(ByRef value As Movement)
    Set mNewMove = value
End Property

Public Sub HoldMove()
    mPreMove.Row = mNewMove.Row
    mPreMove.Column = mNewMove.Column
End Sub
