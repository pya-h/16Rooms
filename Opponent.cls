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

Public Property Get Name() As Variant
    Name = mName
End Property

Public Property Let Name(ByVal val As Variant)
    If val = BodY Or val = Human Then
        mName = val
    Else
        mName = WRONG
    End If
End Property

Public Sub Class_Initialize()
    Name = Game.BodY
    Level = Game.Hard
End Sub

Public Property Get Level() As Variant
    Level = mLevel
End Property

Public Property Let Level(ByVal val As Variant)
    If val = Game.Easy Or val = Game.Hard Then
        mLevel = val
    Else
        mLevel = WRONG
    End If
End Property

Public Function ToString() As String
    ToString = "Opponent: " & IIf(Name = Game.BodY, "BodY", "Human")
End Function

