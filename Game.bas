Attribute VB_Name = "Game"
Public Player2 As New Opponent
' options constants
Public Const TABLE_DIMENSION As Byte = 4, EMPTY_CELL As Byte = 0
Public Const WRONG As Byte = 2
Public Const BodY As Byte = 1, HUMAN As Byte = 0
Public Const AGGRESSIVE_GAMEPLAY_STYLE As Byte = 1, FULL_DEFFENSIVE_GAMEPLAY_STYLE As Byte = 0, RANDOM_GAMEPLAY_STYLE = 2
Public Const OutOfTableValue = 127, NO_UNUSED_PIECES = 255
Public Const OUT_OF_REACH_CELL As Integer = -1
Public table(TABLE_DIMENSION - 1, TABLE_DIMENSION - 1) As Byte

' Constant Values
Public Const AGGRESSIVE_THRESHOLD_VALUE = 2 * ((TABLE_DIMENSION - 1) ^ 2), _
    FULL_DEFENSIVE_THRESHOLD_VALUE = (TABLE_DIMENSION / 2) ^ 2


Public Function UserIsSure(msgTitle As String) As Boolean
    UserIsSure = MsgBox("Are you sure to " & msgTitle, vbYesNo + vbQuestion, msgTitle) = vbYes
End Function

Public Function Sum(arr() As Integer) As Integer
    Sum = 0
    Dim i As Integer
    
    For i = LBound(arr) To UBound(arr)
        Sum = Sum + arr(i)
    Next i
End Function

Public Function GetCellWeight(Row As Byte, col As Byte, us As Byte) As Dict
    Dim k As Byte, attackValue(0 To 3) As Integer, defenceValue(0 To 3) As Integer
    attackValue(0) = 1
    attackValue(1) = 1
    attackValue(2) = 1
    attackValue(3) = 1
    defenceValue(0) = 1
    defenceValue(1) = 1
    defenceValue(2) = 1
    defenceValue(3) = 1
    For k = 0 To TABLE_DIMENSION - 1
        If table(k, col) <> EMPTY_CELL Then
            If table(k, col) = us Then
                attackValue(0) = attackValue(0) + 1
                
                defenceValue(0) = 0
            Else
                attackValue(0) = 0
                defenceValue(0) = defenceValue(0) + 1

            End If
        End If
        
        If table(Row, k) <> EMPTY_CELL Then
            If table(Row, k) = us Then
                attackValue(1) = attackValue(1) + 1
                defenceValue(1) = 0
            Else
                attackValue(1) = 0
                defenceValue(1) = defenceValue(1) + 1
            End If
        End If
        If Row = col Then
            ' main diag move
            If table(k, k) = us Then
                attackValue(2) = attackValue(2) + 1
                defenceValue(2) = 0
            Else
                attackValue(2) = 0
                defenceValue(2) = defenceValue(2) + 1
            End If
        End If
        
        If Row + col = TABLE_DIMENSION - 1 Then
            ' other diag move
            If table(k, TABLE_DIMENSION - k - 1) = us Then
                attackValue(3) = attackValue(3) + 1
                defenceValue(3) = 0
            Else
                attackValue(3) = 0
                defenceValue(3) = defenceValue(3) + 1
            End If
        End If
    Next k

    For k = 0 To 3
        attackValue(k) = attackValue(k) ^ attackValue(k)
        defenceValue(k) = defenceValue(k) ^ defenceValue(k)
    Next k
    
    Set GetCellWeight = New Dict
    Call GetCellWeight.Add("def", Sum(defenceValue))
    Call GetCellWeight.Add("att", Sum(attackValue))


End Function
