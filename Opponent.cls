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
Private mLocked As Boolean
Private mPiece As Byte  ' selected piece ready to move

Public Property Get DragSpeed() As Integer
    DragSpeed = 50
    ' Change this comparing premove and newmove and obtain the speed as a percentage of the difference
End Property
Public Property Get Piece() As Byte
    Piece = mPiece
End Property

Public Property Let Piece(ByVal val As Byte)
    If (val >= 0 And val < TABLE_DIMENSION * 2) Or val = NO_UNUSED_PIECES Then
        mPiece = val
    End If
End Property

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
    ReDim mUnusedPieces(1 To TABLE_DIMENSION + 1)

    Set mNewMove = New Movement
    Set mPreMove = New Movement
End Sub

Public Sub ResetPiecesToUnused(first As Byte)
    Dim i As Byte
    mUnusedPieces(1) = 255
    For i = 2 To UBound(mUnusedPieces)
        mUnusedPieces(i) = first + i - 2
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
    If length > 1 Then
        UnusedPieces = mUnusedPieces(length)
        ReDim Preserve mUnusedPieces(1 To length - 1)
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

Public Property Get Locked() As Boolean
    Locked = mLocked
End Property

Public Sub LockMe()
    mLocked = True
End Sub

Public Sub UnlockMe()
    mLocked = False
End Sub


