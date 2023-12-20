Attribute VB_Name = "Game"
Public Player2 As New Opponent
Public Const TABLE_DIMENSION As Byte = 4, EMPTY_CELL As Byte = 0
'options
Public Const WRONG As Byte = 2
Public Const BodY As Byte = 1, HUMAN As Byte = 0
Public Const EASY As Byte = 1, HARD As Byte = 0
Public Const OutOfTableValue = 127, NO_UNUSED_PIECES = 255

Public Function UserIsSure(msgTitle As String) As Boolean
    UserIsSure = MsgBox("Are you sure to " & msgTitle, vbYesNo + vbQuestion, msgTitle) = vbYes
End Function
