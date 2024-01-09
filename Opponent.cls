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

Private mValue As Byte
Private mName As Byte
Private mUnusedPieces() As Byte
Private mPreMove As New Movement, mNewMove As New Movement
Private mLocked As Boolean
Private mPiece As Byte  ' selected piece ready to move
Private mFirstPieceIndex As Byte, mLastPieceIndex As Byte
Private mGamePlayStyle As Byte
Private mDefenseThreshold As Integer, mAttackThreshold As Integer

Dim DefenceWeights(TABLE_DIMENSION - 1, TABLE_DIMENSION - 1) As Integer, AttackWeights(TABLE_DIMENSION - 1, TABLE_DIMENSION - 1) As Integer

Public Property Get FirstPieceIndex() As Byte
    FirstPieceIndex = mFirstPieceIndex
End Property
Public Property Get LastPieceIndex() As Byte
    LastPieceIndex = mLastPieceIndex
End Property

Public Property Let LastPieceIndex(value As Byte)
    mLastPieceIndex = IIf(value < TABLE_DIMENSION * 2, value, TABLE_DIMENSION * 2 - 1)
End Property

Public Property Let FirstPieceIndex(value As Byte)
    mFirstPieceIndex = value
End Property

Public Sub ResetWeights()
    Dim i As Byte, j As Byte
    For i = 0 To TABLE_DIMENSION - 1
        For j = 0 To TABLE_DIMENSION - 1
            DefenceWeights(i, j) = 0
            AttackWeights(i, j) = 0
        Next j
    Next i
End Sub

Public Sub WeighTable()

    Call ResetWeights
    
    Dim i As Byte, j As Byte
    For i = 0 To TABLE_DIMENSION - 1
        For j = 0 To TABLE_DIMENSION - 1
            If table(i, j) <> EMPTY_CELL Then
                DefenceWeights(i, j) = OUT_OF_REACH_CELL
                AttackWeights(i, j) = OUT_OF_REACH_CELL
                
            Else
                Dim weights As New Dict
                Set weights = GetCellWeight(i, j, mValue)
                
                DefenceWeights(i, j) = weights.Item("def")
                AttackWeights(i, j) = weights.Item("att")
                
            End If
            
        Next j
    Next i
End Sub

Public Property Get BestMove() As Movement
    Call WeighTable
    Dim i As Byte, j As Byte, attackRow As Byte, attackCol As Byte, defenseRow As Byte, defenseCol As Byte
    
    For i = 0 To TABLE_DIMENSION - 1
        For j = 0 To TABLE_DIMENSION - 1
            If AttackWeights(i, j) > AttackWeights(attackRow, attackCol) Then
                attackCol = j
                attackRow = i
                
            End If
            
            If DefenceWeights(i, j) > DefenceWeights(defenseRow, defenseCol) Then
                defenseCol = j
                defenseRow = i
                
            End If
        Next j
    Next i
    
    Set BestMove = New Movement
    If AttackWeights(attackRow, attackCol) >= mAttackThreshold Then
        ' TODO: Find The piece that isnt in the direction of weights
        BestMove.Row = attackRow
        BestMove.Column = attackCol
    ElseIf DefenceWeights(defenseRow, defenseCol) >= mDefenseThreshold Then ' if defense if crucial
        ' Now bugs are happening (Not pReventing 100% score threat)
        BestMove.Row = defenseRow
        BestMove.Column = defenseCol
    Else
        BestMove.Row = attackRow
        BestMove.Column = attackCol
    End If

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

Public Property Let value(val As Byte)
    mValue = val
End Property
Public Property Get value() As Byte
    value = mValue
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
    ReDim mUnusedPieces(1 To TABLE_DIMENSION + 1)
    GamePlayStyle = Game.AGGRESSIVE_GAMEPLAY_STYLE
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

Public Sub SelectLeastSignificantPiece(possibleWeights() As Dict)
    Dim p As Byte, minAttackIndex As Byte, minDeffenseIndex As Byte
    minAttackIndex = LBound(possibleWeights)
    minDeffenseIndex = LBound(possibleWeights)
    For p = LBound(possibleWeights) + 1 To UBound(possibleWeights)
        If possibleWeights(p).Item("att") < possibleWeights(minAttackIndex).Item("att") Then
            minAttackIndex = p
        End If
        
        If possibleWeights(p).Item("def") < possibleWeights(minDeffenseIndex).Item("def") Then
            minDeffenseIndex = p
        End If
    Next p
    Piece = IIf(possibleWeights(minDeffenseIndex).Item("def") < DEFFENSIVE_THRESHOLD, _
        minDeffenseIndex, minAttackIndex)

End Sub
Public Property Get GamePlayStyle() As Byte
    GamePlayStyle = mGamePlayStyle
End Property
Public Property Let GamePlayStyle(val As Byte)
    If val = Game.FULL_DEFFENSIVE_GAMEPLAY_STYLE Then
        mGamePlayStyle = Game.FULL_DEFFENSIVE_GAMEPLAY_STYLE
        mDefenseThreshold = Game.FULL_DEFENSIVE_THRESHOLD_VALUE
        mAttackThreshold = Game.FULL_DEFENSIVE_THRESHOLD_VALUE
    ElseIf val = Game.AGGRESSIVE_GAMEPLAY_STYLE Then
        mGamePlayStyle = Game.AGGRESSIVE_GAMEPLAY_STYLE
        mDefenseThreshold = Game.AGGRESSIVE_THRESHOLD_VALUE
        mAttackThreshold = Game.AGGRESSIVE_THRESHOLD_VALUE
    ElseIf val = Game.RANDOM_GAMEPLAY_STYLE Then
        mGamePlayStyle = Game.RANDOM_GAMEPLAY_STYLE
    End If
End Property
