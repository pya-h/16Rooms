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
   Begin VB.Frame frameGameplayStyle 
      BackColor       =   &H8000000B&
      Caption         =   "Gameplay Style"
      Height          =   3255
      Left            =   1200
      TabIndex        =   5
      Top             =   1920
      Width           =   4335
      Begin VB.OptionButton optRandomGameplay 
         Caption         =   "Random"
         Height          =   615
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2400
         Width           =   2175
      End
      Begin VB.OptionButton optFullDefensiveGameplay 
         Caption         =   "Full Defensive"
         Height          =   615
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton optAgressiveGameplay 
         Caption         =   "Agressive"
         Height          =   615
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   2175
      End
   End
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
      Enabled         =   0   'False
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
    Unload frmOptions
    
End Sub

Private Sub btnOk_Click()
    If TestPlayer.GamePlayStyle <> Player2.GamePlayStyle Or TestPlayer.Name <> Player2.Name Then  ' If anything changes
        If UserIsSure("Save The Changes?") Then
            Player2.Name = TestPlayer.Name
            Player2.GamePlayStyle = TestPlayer.GamePlayStyle
            Unload Me
        End If
    End If
End Sub


Private Sub Form_Load()
    TestPlayer.Name = Player2.Name
    
    If Player2.Name = BodY Then
        optBodY.value = True
    ElseIf Player2.Name = HUMAN Then
        optHuman.value = True
    Else
        optBodY.value = False
        optHuman.value = False
    End If
    If Player2.GamePlayStyle = Game.AGGRESSIVE_GAMEPLAY_STYLE Then
        optAgressiveGameplay.value = True
    ElseIf Player2.GamePlayStyle = Game.FULL_DEFFENSIVE_GAMEPLAY_STYLE Then
        optFullDefensiveGameplay.value = True
    ElseIf Player2.GamePlayStyle = Game.RANDOM_GAMEPLAY_STYLE Then
        optRandomGameplay.value = True
    Else
        optAgressiveGameplay.value = False
        optFullDefensiveGameplay.value = False
        optRandomGameplay.value = False
    End If
End Sub

Private Function SureAboutDiscard() As Boolean
    If TestPlayer.GamePlayStyle <> Player2.GamePlayStyle _
        Or TestPlayer.Name <> Player2.Name Then
            SureAboutDiscard = UserIsSure("Discard The Changes?")
    Else
        SureAboutDiscard = True
    End If

End Function


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpdateOkButtonState
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Cancel = Not SureAboutDiscard()
End Sub

Private Sub optAgressiveGameplay_Click()
    TestPlayer.GamePlayStyle = Game.AGGRESSIVE_GAMEPLAY_STYLE
    UpdateOkButtonState
End Sub

Private Sub optBodY_Click()
    TestPlayer.Name = BodY
    UpdateOkButtonState
End Sub
 
Private Sub optFullDefensiveGameplay_Click()
    TestPlayer.GamePlayStyle = Game.FULL_DEFFENSIVE_GAMEPLAY_STYLE
    UpdateOkButtonState

End Sub

Private Sub optHuman_Click()
    TestPlayer.Name = HUMAN
    UpdateOkButtonState
End Sub

Private Sub optRandomGameplay_Click()
    TestPlayer.GamePlayStyle = Game.RANDOM_GAMEPLAY_STYLE
    UpdateOkButtonState
End Sub

Private Sub UpdateOkButtonState()
    btnOk.Enabled = TestPlayer.GamePlayStyle <> Player2.GamePlayStyle Or TestPlayer.Name <> Player2.Name
End Sub
