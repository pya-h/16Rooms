Attribute VB_Name = "Game"
Public Player2 As New Opponent

'options
Public Const WRONG As Byte = 2
Public Const BodY As Byte = 1, Human As Byte = 0
Public Const Easy As Byte = 1, Hard As Byte = 0
Public Const OutOfTableValue = 127

Public Function UserIsSure(msgTitle As String) As Boolean
    UserIsSure = MsgBox("Are you sure to " & msgTitle, vbYesNo + vbQuestion, msgTitle) = vbYes
End Function


