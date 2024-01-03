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
Private mValue As Byte
Private mName As Byte
Private mUnusedPieces() As Byte
Private mPreMove As New Movement, mNewMove As New Movement
Private mLocked As Boolean
Private mPiece As Byte  ' selected piece ready to move

Dim DefenceWeights(TABLE_DIMENSION - 1, TABLE_DIMENSION - 1) As Integer, AttackWeights(TABLE_DIMENSION - 1, TABLE_DIMENSION - 1) As Integer

Public Sub ResetWeights()
    Dim i As Byte, j As Byte
    For i = 0 To TABLE_DIMENSION - 1
        For j = 0 To TABLE_DIMENSION - 1
            DefenceWeights(i, j) = AttackWeights(i, j) = 0
        Next j
    Next i
End Sub

Public Sub WeighTable()

    Call ResetWeights
    
    Dim i As Byte, j As Byte
    For i = 0 To TABLE_DIMENSION - 1
        For j = 0 To TABLE_DIMENSION - 1
            If table(i, j) <> EMPTY_CELL Then
                DefenceWeights(i, j) = AttackWeights(i, j) = OUT_OF_REACH_CELL
                
            Else
                Dim k As Byte, attackValue(0 To 2) As Integer, defenceValue(0 To 2) As Integer
                attackValue(0) = attackValue(1) = attackValue(2) = 1
                defenceValue(0) = defenceValue(1) = defenceValue(2) = 1
                For k = 0 To TABLE_DIMENSION - 1
                    If table(k, j) <> EMPTY_CELL Then
                        If table(k, j) = mValue Then
                            attackValue(0) = attackValue(0) * 2
                            defenceValue(0) = 0
                        Else
                            attackValue(0) = 0
                            defenceValue(0) = defenceValue(0) * 2
                        End If
                    End If
                    
                    If table(i, k) <> EMPTY_CELL Then
                        If table(i, k) = mValue Then
                            attackValue(1) = attackValue(1) * 2
                            defenceValue(1) = 0
                        Else
                            attackValue(1) = 0
                            defenceValue(1) = defenceValue(1) * 2
                        End If
                    End If
                Next k
                
                DefenceWeights(i, j) = Sum(defenceValue)
                AttackWeights(i, j) = Sum(attackValue)
            End If
            
        Next j
    Next i
End Sub

Public Property Get BestMove() As Movement
    Call WeighTable
    Dim i As Byte, j As Byte, r As Byte, c As Byte
    
    For i = 0 To TABLE_DIMENSION - 1
        For j = 0 To TABLE_DIMENSION - 1
            If AttackWeights(i, j) > AttackWeights(r, c) Then
                c = j
                r = i
                
            End If
        Next j
    Next i
    
    Set BestMove = New Movement
    BestMove.Row = r
    BestMove.Column = c
End Property
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

Public Property Let Value(val As Byte)
    mValue = val
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

Public Property Let PreMove(ByRef Value As Movement)
    Set mPreMove = Value
End Property
Public Property Let NewMove(ByRef Value As Movement)
    Set mNewMove = Value
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


