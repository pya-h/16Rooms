VERSION 5.00
Begin VB.Form frmOptions 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   FillColor       =   &H80000010&
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "frmOptions"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   7905
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancel 
      Height          =   495
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmOptions.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton btnOk 
      Height          =   495
      Left            =   7200
      Picture         =   "frmOptions.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   615
   End
   Begin VB.Frame frmOpponentAs 
      BackColor       =   &H80000005&
      Caption         =   "Opponent As:"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton optHuman 
         Caption         =   " Human"
         Height          =   615
         Left            =   1320
         Picture         =   "frmOptions.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optBodY 
         Caption         =   " BodY"
         Height          =   615
         Left            =   1320
         Picture         =   "frmOptions.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' check if any changes made by user then prompt a are u sure
' write code for changing the game level
' handle the X click
' add checkbox for asking the user for game objects location reset after each winning
' options for color and fonts and maybe sizes
' think more

Public TestPlayer As New Opponent

Private Sub btnCancel_Click()
    If UserIsSure("Discard The Changes?") Then
        Unload Me
    End If
End Sub

Private Sub btnOk_Click()
    If UserIsSure("Save The Changes?") Then
        Player2.Name = TestPlayer.Name
        Unload Me
    End If
End Sub

Private Sub Form_Load()
TestPlayer.Name = Player2.Name

If Player2.Name = BodY Then
    optBodY.Value = True
ElseIf Player2.Name = Human Then
    optHuman.Value = True
Else
    optBodY.Value = False
    optHuman.Value = False
End If

End Sub

Private Sub optBodY_Click()
    TestPlayer.Name = BodY
End Sub
 
Private Sub optHuman_Click()
    TestPlayer.Name = Human
    
End Sub
