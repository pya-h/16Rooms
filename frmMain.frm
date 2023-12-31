VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "16 Rooms"
   ClientHeight    =   9795
   ClientLeft      =   4425
   ClientTop       =   675
   ClientWidth     =   10155
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   10155
   Begin VB.Timer gameTimer 
      Interval        =   1000
      Left            =   10200
      Top             =   9120
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "0 - 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   3960
      TabIndex        =   3
      Top             =   8520
      Width           =   2520
   End
   Begin VB.Label lblState 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   1200
      TabIndex        =   2
      Top             =   9195
      Width           =   8760
   End
   Begin VB.Label lblStateLabel 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   9195
      Width           =   735
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpPlayer 
      Height          =   495
      Left            =   9960
      TabIndex        =   0
      Top             =   9120
      Visible         =   0   'False
      Width           =   495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   873
      _cy             =   873
   End
   Begin VB.Image imgPlayer 
      Height          =   480
      Index           =   4
      Left            =   7680
      Picture         =   "frmMain.frx":0442
      Top             =   8400
      Width           =   480
   End
   Begin VB.Image imgPlayer 
      Height          =   480
      Index           =   5
      Left            =   8280
      Picture         =   "frmMain.frx":0884
      Top             =   8400
      Width           =   480
   End
   Begin VB.Image imgPlayer 
      Height          =   480
      Index           =   6
      Left            =   8880
      Picture         =   "frmMain.frx":0CC6
      Top             =   8400
      Width           =   480
   End
   Begin VB.Image imgPlayer 
      Height          =   480
      Index           =   7
      Left            =   9480
      Picture         =   "frmMain.frx":1108
      Top             =   8400
      Width           =   480
   End
   Begin VB.Image imgPlayer 
      Height          =   480
      Index           =   3
      Left            =   2160
      Picture         =   "frmMain.frx":154A
      Top             =   8520
      Width           =   480
   End
   Begin VB.Image imgPlayer 
      Height          =   480
      Index           =   2
      Left            =   1560
      Picture         =   "frmMain.frx":1854
      Top             =   8520
      Width           =   480
   End
   Begin VB.Image imgPlayer 
      Height          =   480
      Index           =   1
      Left            =   960
      Picture         =   "frmMain.frx":1B5E
      Top             =   8520
      Width           =   480
   End
   Begin VB.Image imgPlayer 
      DragIcon        =   "frmMain.frx":1E68
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "frmMain.frx":22AA
      Top             =   8520
      Width           =   480
   End
   Begin VB.Line linHorizontal 
      BorderStyle     =   4  'Dash-Dot
      Index           =   4
      X1              =   360
      X2              =   9960
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line linHorizontal 
      BorderStyle     =   4  'Dash-Dot
      Index           =   3
      X1              =   360
      X2              =   9960
      Y1              =   6300
      Y2              =   6300
   End
   Begin VB.Line linHorizontal 
      BorderStyle     =   4  'Dash-Dot
      Index           =   2
      X1              =   360
      X2              =   9960
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line linHorizontal 
      BorderStyle     =   4  'Dash-Dot
      Index           =   1
      X1              =   360
      X2              =   9960
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Line linVertical 
      BorderStyle     =   4  'Dash-Dot
      Index           =   4
      X1              =   9960
      X2              =   9960
      Y1              =   360
      Y2              =   8280
   End
   Begin VB.Line linVertical 
      BorderStyle     =   4  'Dash-Dot
      Index           =   3
      X1              =   7560
      X2              =   7560
      Y1              =   360
      Y2              =   8280
   End
   Begin VB.Line linVertical 
      BorderStyle     =   4  'Dash-Dot
      Index           =   2
      X1              =   5160
      X2              =   5160
      Y1              =   360
      Y2              =   8280
   End
   Begin VB.Line linVertical 
      BorderStyle     =   4  'Dash-Dot
      Index           =   1
      X1              =   2760
      X2              =   2760
      Y1              =   360
      Y2              =   8280
   End
   Begin VB.Line linVertical 
      BorderStyle     =   4  'Dash-Dot
      Index           =   0
      X1              =   360
      X2              =   360
      Y1              =   360
      Y2              =   8280
   End
   Begin VB.Line linHorizontal 
      BorderStyle     =   4  'Dash-Dot
      Index           =   0
      X1              =   360
      X2              =   9960
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnuNew 
         Caption         =   "New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "Load Game"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Game"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Dim pressed As Boolean
Dim x0, y0 As Single
Const p1FirstPieceIndex As Byte = 0, p1Value As Byte = 1, LeftIndex As Byte = 0, TopIndex As Byte = 1 _
    , MaxDimensionIndex As Byte = TABLE_DIMENSION - 1, totalPiecesIndex As Byte = TABLE_DIMENSION * 2 - 1 'from 0
Const SendToBack As Byte = 1, BringToFront As Byte = 0

Dim PrimaryPositions(totalPiecesIndex, 2) As Integer, previousPositions(totalPiecesIndex, 2) As Integer
Dim DeltaCenter As Integer ' For Finding the imageview center
Dim NewMove As New Movement, PreMove As New Movement, playerTurn As Byte
Dim scores(1) As Integer

Private Sub btnReset_Click()
    ResetGame True
End Sub

Private Sub Command1_Click()
    Form1.Show
End Sub

'       TODO:
'       change the user interface maybe?
'       change the mouse cursor icon maybe      '
'       optimize the code       '    '

'       save game       '
'       creat a menu for editing the game interface     '

' I think primary and previous positions have no use !
' define the first player as an opponent object too.
' Make drag speed dynamic
Private Sub Form_Load()
    Set Player2 = New Opponent
    Player2.value = 2
    Player2.FirstPieceIndex = TABLE_DIMENSION
    Player2.LastPieceIndex = TABLE_DIMENSION * 2 - 1
    Set NewMove = New Movement
    Set PreMove = New Movement

    Dim i As Byte
    For i = 0 To totalPiecesIndex
        PrimaryPositions(i, LeftIndex) = imgPlayer(i).Left
        PrimaryPositions(i, TopIndex) = imgPlayer(i).Top
    Next i
    
    lblResult.ZOrder (SendToBack)
    ResetGame False

    DeltaCenter = imgPlayer(0).Width / 2
End Sub

Private Sub gameTimer_Timer()
    If Player2.Locked Then
        Dim dy As Integer, dx As Integer, x As Integer, Y As Integer
        x = (linVertical(Player2.NewMove.Column).X1 + linVertical(Player2.NewMove.Column + 1).X1) / 2 ' - DeltaCenter
        Y = (linHorizontal(Player2.NewMove.Row).Y1 + linHorizontal(Player2.NewMove.Row + 1).Y1) / 2 ' - DeltaCenter
        dx = IIf(x >= imgPlayer(Player2.Piece).Left, Player2.DragSpeed, -Player2.DragSpeed)
        dy = IIf(Y >= imgPlayer(Player2.Piece).Top, Player2.DragSpeed, -Player2.DragSpeed)
        Dim reachedX As Boolean, reachedY As Boolean
        reachedX = Abs(x - imgPlayer(Player2.Piece).Left) <= Player2.DragSpeed
        reachedY = Abs(Y - imgPlayer(Player2.Piece).Top) <= Player2.DragSpeed

        If Not reachedX Then
            imgPlayer(Player2.Piece).Left = imgPlayer(Player2.Piece).Left + dx
        End If
        If Not reachedY Then
            imgPlayer(Player2.Piece).Top = imgPlayer(Player2.Piece).Top + dy
        End If
        If reachedX And reachedY Then
            Player2.UnlockMe
            gameTimer.Interval = 1000
            Call SubmitMove(Player2.NewMove, Player2.PreMove, Player2.Piece)
        End If
    ElseIf playerTurn = 1 And Player2.Name = BodY Then
        Call DoBodYMove
    End If
    
    
End Sub

Private Sub imgPlayer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If IsThisPlayerTurn(Index) Then
        pressed = True
        x0 = x
        y0 = Y
        previousPositions(Index, LeftIndex) = imgPlayer(Index).Left
        previousPositions(Index, TopIndex) = imgPlayer(Index).Top
        If previousPositions(Index, LeftIndex) = PrimaryPositions(Index, LeftIndex) And previousPositions(Index, TopIndex) = PrimaryPositions(Index, TopIndex) Then
            Call PreMove.PutOutOfTable
        Else
            Set NewMove = GetPositionOnTable(CByte(Index))
            PreMove.Row = NewMove.Row
            PreMove.Column = NewMove.Column
        End If
    Else
        Beep
    End If
End Sub

Private Sub imgPlayer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    'the code for enabling imgPlayer dragging
    If pressed = True Then
        imgPlayer(Index).Left = imgPlayer(Index).Left + x - x0
        imgPlayer(Index).Top = imgPlayer(Index).Top + Y - y0
    End If
End Sub
Private Sub SubmitMove(ByRef newPlace As Movement, ByRef previousPlace As Movement, ByVal pieceIndex As Integer)
    
    If previousPlace.IsInsideTable Then
        table(previousPlace.Row, previousPlace.Column) = EMPTY_CELL
    End If
    
    wmpPlayer.URL = App.Path + "\moved.wav"  'play piece move sound
    table(newPlace.Row, newPlace.Column) = IIf(pieceIndex < Player2.FirstPieceIndex, p1Value, Player2.value)
    'Set the imgPlayer location at the center of the room
    imgPlayer(pieceIndex).Top = (linHorizontal(newPlace.Row).Y1 + linHorizontal(newPlace.Row + 1).Y1) / 2 - DeltaCenter
    imgPlayer(pieceIndex).Left = (linVertical(newPlace.Column).X1 + linVertical(newPlace.Column + 1).X1) / 2 - DeltaCenter
    Dim winner As Byte: winner = 0
    winner = CheckForWinner()
    If winner <> EMPTY_CELL Then
        Call ScoreNotification(winner)
    Else
        Call ManageTurns
    End If
End Sub
Private Sub imgPlayer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If IsThisPlayerTurn(Index) Then
        pressed = False
        x0 = 0
        y0 = 0
        
        If imgPlayer(Index).Left + DeltaCenter < linVertical(0).X1 Or imgPlayer(Index).Left + DeltaCenter > linVertical(MaxDimensionIndex + 1).X1 _
            Or imgPlayer(Index).Top + DeltaCenter < linHorizontal(0).Y1 Or imgPlayer(Index).Top + DeltaCenter > linHorizontal(MaxDimensionIndex + 1).Y1 Then
            
            ShowError "Destination is out of the table range!"
            RollbackMove Index
            
        Else
        
            Set NewMove = GetPositionOnTable(CByte(Index)) 'find the NewMove.Row and NewMove.Column variables values
                'edit table array value  with player data
            If table(NewMove.Row, NewMove.Column) <> EMPTY_CELL Then
                ShowError "Destination room is not empty!"
                RollbackMove Index
            Else
                Call SubmitMove(NewMove, PreMove, Index)
            End If
            
        End If
        
    End If
End Sub

Private Function GetPositionOnTable(Index As Byte) As Movement
    Dim position As Movement
    Set position = New Movement
    
    Dim r As Byte
    ' Find current location NewMove.Row
    For r = 0 To MaxDimensionIndex
        If imgPlayer(Index).Top + DeltaCenter <= linHorizontal(r + 1).Y1 Then
            position.Row = r
            Exit For
        End If
    Next r
    
    ' Find current location NewMove.Column
    For r = 0 To MaxDimensionIndex
        If imgPlayer(Index).Left + DeltaCenter <= linVertical(r + 1).X1 Then
            position.Column = r
            Exit For
        End If
    Next r
    Set GetPositionOnTable = position
End Function

Private Sub RollbackMove(Index As Integer)
    imgPlayer(Index).Left = previousPositions(Index, LeftIndex)
    imgPlayer(Index).Top = previousPositions(Index, TopIndex)
End Sub

Private Sub ResetGame(userRequestedTheReset As Boolean)
    
    playerTurn = 1 'see ManageTurns Sub and you'l see why:)
    Call ManageTurns
    pressed = False
    x0 = 0
    y0 = 0
    Call NewMove.PutOutOfTable
    Call PreMove.PutOutOfTable
    scores(0) = 0
    scores(1) = 0
    lblResult.Caption = "0 - 0"
    Dim i As Byte
    For i = 0 To totalPiecesIndex
        previousPositions(i, LeftIndex) = PrimaryPositions(i, LeftIndex)
        previousPositions(i, TopIndex) = PrimaryPositions(i, TopIndex)
        
        If userRequestedTheReset Then
            imgPlayer(i).Left = PrimaryPositions(i, LeftIndex)
            imgPlayer(i).Top = PrimaryPositions(i, TopIndex)
        End If
        
    Next i
    
    If userRequestedTheReset Then
        For i = 0 To MaxDimensionIndex
            Dim j As Byte
            For j = 0 To MaxDimensionIndex
                Game.table(i, j) = EMPTY_CELL
                
            Next j
        Next i
    End If
    Call Player2.ResetPiecesToUnused(Player2.FirstPieceIndex)
End Sub

Private Sub mnuNew_Click()
    If UserIsSure("Reset Game") Then
        ResetGame True
    End If
End Sub


Private Sub mnuOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub mnuQuit_Click()
    If UserIsSure("Quit The Game?") Then
        Unload Me
        End
    End If
End Sub

Private Sub ManageTurns()
    playerTurn = (playerTurn + 1) Mod 2
    lblState.ForeColor = vbBlue
    lblState.Caption = "Player " + Str(playerTurn + 1) + "'s Turn"
End Sub

Private Function IsThisPlayerTurn(Index As Integer) As Boolean
    IsThisPlayerTurn = (playerTurn = 0 And Index < Player2.FirstPieceIndex) Or _
        (playerTurn = 1 And Player2.Name = HUMAN And Index >= Player2.FirstPieceIndex)
End Function

Private Sub ShowError(text As String)
    wmpPlayer.URL = App.Path + "\error.wav"
    lblState.ForeColor = vbRed
    lblState.Caption = text
End Sub

Private Function CheckForWinner() As Byte
    Dim i, j, firstRoom As Byte
    Dim allTheSame As Boolean
    firstRoom = EMPTY_CELL
    
    ' horizontal check
    For i = 0 To MaxDimensionIndex
        firstRoom = table(i, 0)
        allTheSame = True
        
        For j = 1 To MaxDimensionIndex
            If table(i, j) <> firstRoom Then
                allTheSame = False
                Exit For
            End If
        Next j
        
        If firstRoom <> EMPTY_CELL And allTheSame = True Then
           CheckForWinner = firstRoom
           Exit Function
        End If
    Next i
    
    ' vertical check
    For i = 0 To MaxDimensionIndex
        firstRoom = table(0, i)
        allTheSame = True
        
        For j = 1 To MaxDimensionIndex
            If table(j, i) <> firstRoom Then
                allTheSame = False
                Exit For
            End If
        Next j
        
        If firstRoom <> EMPTY_CELL And allTheSame = True Then
           CheckForWinner = firstRoom
           Exit Function
        End If
    Next i
    
    ' X
    firstRoom = table(0, 0)
    allTheSame = True
    For i = 1 To MaxDimensionIndex
        If table(i, i) <> firstRoom Then
            allTheSame = False
            Exit For
        End If
    Next i
    
    If firstRoom <> EMPTY_CELL And allTheSame = True Then
        CheckForWinner = firstRoom
        Exit Function
    End If
    
    firstRoom = table(0, MaxDimensionIndex)
    allTheSame = True
    For i = 1 To MaxDimensionIndex
        If table(i, MaxDimensionIndex - i) <> firstRoom Then
            allTheSame = False
            Exit For
        End If
    Next i
    
    If firstRoom <> EMPTY_CELL And allTheSame = True Then
        CheckForWinner = firstRoom
        Exit Function
    End If
    
    CheckForWinner = EMPTY_CELL
End Function

Private Sub ScoreNotification(winner As Byte)
    wmpPlayer.URL = App.Path + "\win.wav"
    lblState.ForeColor = vbGreen
    lblState.Caption = "Player " & winner & " Scored!"
    scores(winner - 1) = scores(winner - 1) + 1
    lblResult.Caption = scores(p1Value - 1) & " - " & scores(Player2.value - 1)
End Sub

Private Sub DoBodYMove()
    If Player2.GamePlayStyle <> Game.RANDOM_GAMEPLAY_STYLE Then
        Player2.NewMove = Player2.BestMove
    Else
        Do
            Call Player2.NewMove.RandomizeMove
        Loop While table(Player2.NewMove.Row, Player2.NewMove.Column) <> EMPTY_CELL
    End If
    
    Player2.Piece = Player2.UnusedPieces
    
    If Player2.Piece = NO_UNUSED_PIECES Then
        ' Now: get each piece
        Dim i As Byte, position As Movement, temp As Byte
        Randomize Timer
        Player2.Piece = CByte(Rnd * (TABLE_DIMENSION - 1) + Player2.FirstPieceIndex) ' choose random in case the below algorythm couldnt find best match
        
        If Player2.GamePlayStyle <> Game.RANDOM_GAMEPLAY_STYLE Then
            ReDim possibleWeights(Player2.FirstPieceIndex To Player2.LastPieceIndex) As Dict
    
            For i = Player2.FirstPieceIndex To Player2.LastPieceIndex
                Set position = GetPositionOnTable(i)
                ' Now assume this position on table is empty, then calculate its value:
                temp = table(position.Row, position.Column)
                table(position.Row, position.Column) = EMPTY_CELL
                Set possibleWeights(i) = GetCellWeight(position.Row, position.Column, Player2.value)
    
                table(position.Row, position.Column) = temp  'Place the piece on its position again
            Next i
            
            ' calculate the value
            ' use the min valuest
            Call Player2.SelectLeastSignificantPiece(possibleWeights)
        End If
        Player2.PreMove = GetPositionOnTable(Player2.Piece)
    Else
        Call Player2.PreMove.PutOutOfTable
    End If
    ' READ PREVIOUS POSITION OF PIECE AND SET IT TO EMPTY
    ' OR DEFINE NewMove and PreMove fields for Player2, just like Player1
    ' or maybe define both of them as Opponent Object (BETTER)
    Player2.LockMe
    gameTimer.Interval = 1
    
End Sub

