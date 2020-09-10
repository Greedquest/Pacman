VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PacmanGame.Controller"
Option Explicit
Private Const mModuleName As String = "GameController"

Private Const MILLISECONDS_PER_TICK As Double = 0 '// sets a minimum amount of time (in milliseconds) that will pass between game ticks
Private Const TICK_CYCLE_RESOLUTION As Double = 10 '// helps faciliate game pieces moving at different speeds

'@MemberAttribute VB_VarHelpID, -1
Public WithEvents UIAdapter As ViewAdapter
Attribute UIAdapter.VB_VarHelpID = -1
Private WithEvents gameLoopTicker As timing.Metronome
Attribute gameLoopTicker.VB_VarHelpID = -1

Public Enum Direction
    dNone = 0
    dUp = -1
    dDown = 1
    dLeft = -2
    dRight = 2
End Enum

'//Encasulated Fields
Private Type TGameController
    IsGameOver As Boolean
    Maze() As Tile
    TickCounter As Long
    Ghosts As Collection
    GamePieces As Collection
    Player As PacmanModel
End Type

Private this As TGameController


Public Sub StartGame()
    '// this is here to temporarily provide a way for me to kick off the game from code
    UIAdapter_GameStarted
    
    '//play intro
    this.TickCounter = 0
    gameLoopTicker.Start
End Sub

Private Sub Class_Initialize()
    Set this.GamePieces = New Collection
    Set gameLoopTicker = Metronome.Create(MILLISECONDS_PER_TICK, delayedStart:=True)
End Sub

Private Sub Class_Terminate()
    Debug.Print TypeName(Me) & " terminating..."
    Set this.GamePieces = Nothing
    
    UIAdapter.Deconstruct
    
    Erase this.Maze
    Erase MapManager.Maze

    Set UIAdapter = Nothing
End Sub

Private Sub gameLoopTicker_Tick()
    this.TickCounter = this.TickCounter + 1
    'If  TickCounter = MaxCycles Then  TickCounter = 0
    If this.IsGameOver Then
        Set gameLoopTicker = Nothing
    Else
        On Error Resume Next
        OnTick
        'PrintMessageQueue Application.hWnd
    End If
End Sub

'// This is the main engine of the game that is called repeatedly until the game is over
Private Sub OnTick()

    
    Dim character As IGamePiece
    
    For Each character In this.GamePieces
    
        If character.CycleRemainder >= TICK_CYCLE_RESOLUTION Then
            character.CycleRemainder = character.CycleRemainder Mod TICK_CYCLE_RESOLUTION
            character.Move
                    
        Else
            If this.TickCounter Mod Round(TICK_CYCLE_RESOLUTION / (TICK_CYCLE_RESOLUTION * (1 - character.Speed)), 0) <> 0 Then
               character.CycleRemainder = character.CycleRemainder + TICK_CYCLE_RESOLUTION Mod (TICK_CYCLE_RESOLUTION * (1 - character.Speed))
               character.Move
            End If
            
            If Round(TICK_CYCLE_RESOLUTION / (TICK_CYCLE_RESOLUTION * (1 - character.Speed)), 0) = 1 Then
               character.CycleRemainder = character.CycleRemainder + TICK_CYCLE_RESOLUTION Mod (TICK_CYCLE_RESOLUTION * (1 - character.Speed))
            End If
            
        End If
    Next
    
    '// TODO: check if player died and/or there is a game over... account for player Lives > 1
    'If this.Player.IsDead Then IsGameOver = True
    
    '// update the view
    UIAdapter.AsCommandSender.UpdateComponents this.GamePieces
    

End Sub

'//ViewEvents Handling
Private Sub UIAdapter_DirectionalKeyPressed(vbKeyCode As KeyCode)
    Select Case vbKeyCode
        Case KeyCode.UpArrow
            this.Player.Heading = dUp
        Case KeyCode.DownArrow
            this.Player.Heading = dDown
        Case KeyCode.LeftArrow
            this.Player.Heading = dLeft
        Case KeyCode.RightArrow
            this.Player.Heading = dRight
    End Select
End Sub

Private Sub UIAdapter_GameStarted()
'// TODO: unbloat this a bit!

    '// initialize vars
        '//scoreboard
        '//
        
    '// initialize game peices
    Dim blinky As GhostModel
    Dim inky As GhostModel
    Dim pinky As GhostModel
    Dim clyde As GhostModel
    
    '// set up maze
    this.Maze = MapManager.LoadMapFromFile
    MapManager.Maze = this.Maze
    UIAdapter.AsCommandSender.CreateMap this.Maze
    
    '// set up pacman
    Set this.Player = New PacmanModel
    Set this.Player.CurrentTile = MapManager.GetMazeTile(46, 30)
    this.GamePieces.Add this.Player
    UIAdapter.AsCommandSender.CreatePacman this.Player
    
    '// set up ghosts
    Set blinky = BuildGhost("Blinky", vbRed, MapManager.GetMazeTile(22, 30), ShadowBehavior.Create(this.Player))
    this.GamePieces.Add blinky
    UIAdapter.AsCommandSender.CreateGhost blinky
    
    Set pinky = BuildGhost("Pinky", rgbLightPink, MapManager.GetMazeTile(22, 20), SpeedyBehavior.Create(this.Player))
    this.GamePieces.Add pinky
    UIAdapter.AsCommandSender.CreateGhost pinky
    
    Set inky = BuildGhost("Inky", vbCyan, MapManager.GetMazeTile(22, 34), BashfulBehavior.Create(this.Player, blinky))
    this.GamePieces.Add inky
    UIAdapter.AsCommandSender.CreateGhost inky
    
    Set clyde = BuildGhost("Clyde", rgbOrange, MapManager.GetMazeTile(22, 37), RandomBehavior.Create())
    Set clyde.ActiveBehavior = PokeyBehavior.Create(this.Player, clyde, MapManager.GetMazeTile( _
                                                                                                UBound(this.Maze, 1), _
                                                                                                LBound(this.Maze, 2) _
                                                                                                ) _
                                                    )
    this.GamePieces.Add clyde
    UIAdapter.AsCommandSender.CreateGhost clyde
    

        
End Sub

'//Private Helpers
Private Function BuildGhost(Name As String, _
                            Color As Long, _
                            startTile As Tile, behavior As IGhostBehavior) As GhostModel
    Dim newGhost As GhostModel
    Set newGhost = New GhostModel
    
    With newGhost
        .Name = Name
        .Color = Color
        Set .CurrentTile = startTile
        Set .ActiveBehavior = behavior
    End With
    
    Set BuildGhost = newGhost
End Function

Private Sub BuildGameBoard()
    UIAdapter.AsCommandSender.CreateMap this.Maze
End Sub


