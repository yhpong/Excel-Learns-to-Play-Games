VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEnvrion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================
'2023-04-25
' This is the environment for player to interact with
' here we define the rules of a game, the reward schemes etc.
' some global variables are common to all games, while some are game specific
' TakeStep_xx(), this subroutine defines how player's action would
' affect the game states, and what rewards would the player receive
' Init() defines and initialize state variables, and other relevant objects in the game
' Reset() resets a game to its starting state
'=============================================================

Private isInit As Boolean
Private pstrGame As String

Private ptime_step As Double
Private pmax_step As Long
Private pStates() As Double, pStatesNxt() As Double, pn_state As Long

Private pminx As Long, pmaxx As Long, pminy As Long, pmaxy As Long
Private pn_block As Long, pBlockPos() As Long
Private pStartPos() As Long, pTgtPos() As Long, pPlayerPos() As Long
Private pn_visited As Long, pVisitedPos() As Long

Private pturtle_velocity As Double, pturtleSpeedMin As Double, pturtleSpeedMax As Double
Private pPlayer_pos_init As Double, pturtle_pos_init As Double, pturtle_height As Double
Private pgravity As Double
Private pjump_speed As Double
Private half_at2 As Double, a_t As Double


Property Get States() As Double()
    States = pStates
End Property


'Initialize a game, this defines boundaries and physics of
'a game, what state variables are available etc.
Sub Init(strGame)
    
    pstrGame = Trim(UCase(strGame))
    
    If UCase(pstrGame) = "MAZE" Then
        
        pminx = 1: pmaxx = 7
        pminy = 1: pmaxy = 7
        pn_state = 7 * 7
        pmax_step = Int((pmaxy - pminy + 1) * (pmaxx - pminx + 1) * 1.2)
        ReDim pStates(1 To pn_state, 1 To 1)
        ReDim pStartPos(1 To 2, 1 To 1)
        ReDim pTgtPos(1 To 2, 1 To 1)
        ReDim pPlayerPos(1 To 2, 1 To 1)
        
        pStartPos(1, 1) = 1
        pStartPos(2, 1) = 7
        pTgtPos(1, 1) = 7
        pTgtPos(2, 1) = 1
        pPlayerPos(1, 1) = pStartPos(1, 1)
        pPlayerPos(2, 1) = pStartPos(2, 1)
        
        pn_visited = 1
        ReDim pVisitedPos(1 To 2, 1 To 1)
        pVisitedPos(1, 1) = pStartPos(1, 1)
        pVisitedPos(2, 1) = pStartPos(2, 1)
        
        pn_block = 15
        ReDim pBlockPos(1 To 2, 1 To pn_block)
        pBlockPos(1, 1) = 1: pBlockPos(2, 1) = 5
        pBlockPos(1, 2) = 2: pBlockPos(2, 2) = 2
        pBlockPos(1, 3) = 2: pBlockPos(2, 3) = 3
        pBlockPos(1, 4) = 2: pBlockPos(2, 4) = 5
        pBlockPos(1, 5) = 2: pBlockPos(2, 5) = 7
        pBlockPos(1, 6) = 3: pBlockPos(2, 6) = 3
        pBlockPos(1, 7) = 3: pBlockPos(2, 7) = 5
        pBlockPos(1, 8) = 4: pBlockPos(2, 8) = 1
        pBlockPos(1, 9) = 4: pBlockPos(2, 9) = 3
        pBlockPos(1, 10) = 4: pBlockPos(2, 10) = 6
        pBlockPos(1, 11) = 5: pBlockPos(2, 11) = 4
        pBlockPos(1, 12) = 5: pBlockPos(2, 12) = 6
        pBlockPos(1, 13) = 6: pBlockPos(2, 13) = 4
        pBlockPos(1, 14) = 7: pBlockPos(2, 14) = 5
        pBlockPos(1, 15) = 7: pBlockPos(2, 15) = 6
    
    ElseIf UCase(pstrGame) = "MAZEII" Then
    
        pn_state = 2 + 2 + 8 + 8
        pminx = 1: pmaxx = 13
        pminy = 1: pmaxy = 7
        pmax_step = Int((pmaxy - pminy + 1) * (pmaxx - pminx + 1) * 1.2)
        ReDim pStates(1 To pn_state, 1 To 1)
        ReDim pStartPos(1 To 2, 1 To 1)
        ReDim pTgtPos(1 To 2, 1 To 1)
        ReDim pPlayerPos(1 To 2, 1 To 1)
        
        pStartPos(1, 1) = 1
        pStartPos(2, 1) = 7
        pTgtPos(1, 1) = 13
        pTgtPos(2, 1) = 6
        pPlayerPos(1, 1) = pStartPos(1, 1)
        pPlayerPos(2, 1) = pStartPos(2, 1)
        
        pn_visited = 1
        ReDim pVisitedPos(1 To 2, 1 To 1)
        pVisitedPos(1, 1) = pStartPos(1, 1)
        pVisitedPos(2, 1) = pStartPos(2, 1)
        
        pn_block = 27
        ReDim pBlockPos(1 To 2, 1 To pn_block)
        pBlockPos(1, 1) = 1: pBlockPos(2, 1) = 5
        pBlockPos(1, 2) = 2: pBlockPos(2, 2) = 2
        pBlockPos(1, 3) = 2: pBlockPos(2, 3) = 3
        pBlockPos(1, 4) = 2: pBlockPos(2, 4) = 5
        pBlockPos(1, 5) = 2: pBlockPos(2, 5) = 7
        pBlockPos(1, 6) = 3: pBlockPos(2, 6) = 3
        pBlockPos(1, 7) = 3: pBlockPos(2, 7) = 5
        pBlockPos(1, 8) = 4: pBlockPos(2, 8) = 1
        pBlockPos(1, 9) = 4: pBlockPos(2, 9) = 3
        pBlockPos(1, 10) = 4: pBlockPos(2, 10) = 6
        pBlockPos(1, 11) = 5: pBlockPos(2, 11) = 4
        pBlockPos(1, 12) = 5: pBlockPos(2, 12) = 6
        pBlockPos(1, 13) = 6: pBlockPos(2, 13) = 4
        pBlockPos(1, 14) = 7: pBlockPos(2, 14) = 5
        pBlockPos(1, 15) = 7: pBlockPos(2, 15) = 6
        pBlockPos(1, 16) = 7: pBlockPos(2, 16) = 7
        pBlockPos(1, 17) = 8: pBlockPos(2, 17) = 2
        pBlockPos(1, 18) = 9: pBlockPos(2, 18) = 2
        pBlockPos(1, 19) = 10: pBlockPos(2, 19) = 2
        pBlockPos(1, 20) = 11: pBlockPos(2, 20) = 2
        pBlockPos(1, 21) = 12: pBlockPos(2, 21) = 2
        pBlockPos(1, 22) = 13: pBlockPos(2, 22) = 2
        pBlockPos(1, 23) = 9: pBlockPos(2, 23) = 5
        pBlockPos(1, 24) = 10: pBlockPos(2, 24) = 5
        pBlockPos(1, 25) = 11: pBlockPos(2, 25) = 5
        pBlockPos(1, 26) = 12: pBlockPos(2, 26) = 5
        pBlockPos(1, 27) = 13: pBlockPos(2, 27) = 5
    
    
    ElseIf UCase(pstrGame) = "STOMP" Then
        
        pn_state = 5 '1=turtle postion, 2=turtle speed, 3=player position, 4=player speed, 5=jump lock
        ReDim pStates(1 To pn_state, 1 To 1)
        
        ptime_step = 0.5
        pgravity = 1.5
        half_at2 = 0.5 * pgravity * (ptime_step ^ 2)
        a_t = pgravity * ptime_step
        pjump_speed = 7.8
        pturtleSpeedMin = -1.5
        pturtleSpeedMax = -0.8
        pturtle_pos_init = 30
        pPlayer_pos_init = 0
        pturtle_height = 10
        pmax_step = Int(-(pturtle_pos_init + 11) / (pturtleSpeedMax * ptime_step))
        
        ReDim pStartPos(1 To 2, 1 To 1)
        ReDim pTgtPos(1 To 2, 1 To 1)
        ReDim pPlayerPos(1 To 2, 1 To 1)
        
        pStartPos(1, 1) = 0
        pStartPos(2, 1) = pPlayer_pos_init
        pPlayerPos(1, 1) = pStartPos(1, 1)
        pPlayerPos(2, 1) = pStartPos(2, 1)
        pTgtPos(1, 1) = pturtle_pos_init
        pTgtPos(2, 1) = 0
        
        pStates(1, 1) = pturtle_pos_init / 20
        pStates(2, 1) = pturtleSpeedMin + Rnd() * (pturtleSpeedMax - pturtleSpeedMin)
        pStates(3, 1) = pPlayer_pos_init / 20
        pStates(4, 1) = 0
        pStates(5, 1) = 0
        
    End If
    
    isInit = True
    
End Sub


'Reset a game to starting state
Sub Reset(Optional strDisplaySheet As Variant, _
            Optional strChtName As String = "Chart 1", _
            Optional FrameRate As Double = 20)
Dim i As Long, j As Long, k As Long, ii As Long, jj As Long, m As Long

    If UCase(pstrGame) = "MAZE" Then
        
        ReDim pStates(1 To 49, 1 To 1)
        pStates((pStartPos(1, 1) - 1) * 7 + pStartPos(2, 1), 1) = 1
        pn_visited = 1
        ReDim pVisitedPos(1 To 2, 1 To 1)
        pVisitedPos(1, 1) = pStartPos(1, 1)
        pVisitedPos(2, 1) = pStartPos(2, 1)
        pPlayerPos(1, 1) = pStartPos(1, 1)
        pPlayerPos(2, 1) = pStartPos(2, 1)
    
    ElseIf UCase(pstrGame) = "MAZEII" Then

        pn_visited = 1
        ReDim pVisitedPos(1 To 2, 1 To 1)
        pVisitedPos(1, 1) = pStartPos(1, 1)
        pVisitedPos(2, 1) = pStartPos(2, 1)
        pPlayerPos(1, 1) = pStartPos(1, 1)
        pPlayerPos(2, 1) = pStartPos(2, 1)
        pStates(1, 1) = pPlayerPos(1, 1) / pmaxx
        pStates(2, 1) = pPlayerPos(2, 1) / pmaxy
        pStates(3, 1) = pPlayerPos(1, 1) / pmaxx
        pStates(4, 1) = pPlayerPos(2, 1) / pmaxy
        Call scan_obstacles(pStates)
 
    ElseIf UCase(pstrGame) = "STOMP" Then
    
        pPlayerPos(1, 1) = pStartPos(1, 1)
        pPlayerPos(2, 1) = pStartPos(2, 1)
        pTgtPos(1, 1) = pturtle_pos_init
        pTgtPos(2, 1) = 0
    
        pStates(1, 1) = pturtle_pos_init / 20
        pStates(2, 1) = pturtleSpeedMin + Rnd() * (pturtleSpeedMax - pturtleSpeedMin)
        pStates(3, 1) = pPlayer_pos_init / 20
        pStates(4, 1) = 0
        pStates(5, 1) = 0
        pmax_step = Int(-(pturtle_pos_init + 11) / (pStates(2, 1) * ptime_step))
    
    End If
    
    If Not IsMissing(strDisplaySheet) Then
        Application.ScreenUpdating = True
        Call DisplayOnScreen(strDisplaySheet, strDisplaySheet.ChartObjects(strChtName).Chart, pStates, FrameRate, 0, False, False, 0)
    End If
    
End Sub



Private Sub TakeStep_STOMP(x_action As Long, tmp_score As Double, isHit As Boolean, isDead As Boolean)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim x_turtle As Double, u_turtle As Double
Dim y_mario As Double, v_mario As Double
Dim tmp_x As Double
Dim jump_lock As Long

    tmp_score = 0
    
    x_turtle = pStates(1, 1) * 20
    u_turtle = pStates(2, 1)
    y_mario = pStates(3, 1) * 20
    v_mario = pStates(4, 1)
    jump_lock = pStates(5, 1)
    
    'Update States
    x_turtle = x_turtle + u_turtle * ptime_step
    If x_action = 1 And jump_lock = 0 Then
        'jump only works when mario is on the ground
        v_mario = pjump_speed
        jump_lock = 1
        tmp_score = tmp_score - 0.1 'Penalize excessive jumps
    ElseIf x_action = 1 And jump_lock = 1 Then
        tmp_score = tmp_score - 1 'Penalize jump command when it's not possible
    End If
    
    If v_mario <> 0 Or y_mario > 0 Then
        y_mario = y_mario + v_mario * ptime_step - half_at2
        v_mario = v_mario - a_t
        'touch down, reset jump
        If y_mario <= 0 Then
            y_mario = 0
            v_mario = 0
            jump_lock = 0
        End If
    End If
    
    pPlayerPos(2, 1) = y_mario
    pTgtPos(1, 1) = x_turtle
    
    pStates(1, 1) = x_turtle / 20
    pStates(2, 1) = u_turtle
    pStates(3, 1) = y_mario / 20
    pStates(4, 1) = v_mario
    pStates(5, 1) = jump_lock
    
    'Check for hit box collision
    If x_turtle >= -9 And x_turtle <= 0 And (y_mario >= 8 And y_mario <= 10.5 And v_mario < 0) Then
        tmp_x = 10 - 2 * Abs(x_turtle + 5)
        tmp_score = tmp_score + tmp_x
        isHit = True
    ElseIf x_turtle >= -10 And x_turtle < (10 - y_mario) / 2 And y_mario <= 8 Then
        tmp_score = tmp_score - 10
        isDead = True
    End If
    
End Sub


Private Sub TakeStep_Maze(x_action As Long, tmp_score As Double, isHit As Boolean)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim x_pos As Long, y_pos As Long
Dim x_pos_old As Long, y_pos_old As Long
Dim tmp_x As Double
Dim isVisited As Boolean

    tmp_score = 0
    
    x_pos = pPlayerPos(1, 1)
    y_pos = pPlayerPos(2, 1)
    x_pos_old = x_pos
    y_pos_old = y_pos
    
    If x_action = 1 Then
        y_pos = y_pos + 1
    ElseIf x_action = 2 Then
        x_pos = x_pos + 1
    ElseIf x_action = 3 Then
        y_pos = y_pos - 1
    ElseIf x_action = 4 Then
        x_pos = x_pos - 1
    End If
    
    If x_pos = pTgtPos(1, 1) And y_pos = pTgtPos(2, 1) Then
        tmp_score = tmp_score + 10
        isHit = True
        pPlayerPos(1, 1) = x_pos
        pPlayerPos(2, 1) = y_pos
        pStates((x_pos_old - 1) * 7 + y_pos_old, 1) = 0
        pStates((x_pos - 1) * 7 + y_pos, 1) = 1
        Exit Sub
    End If
    
    'Penalize long wandering
    tmp_score = tmp_score - 0.04
    
    'Penalize running into walls
    If x_pos < pminx Then
        x_pos = pminx
        tmp_score = tmp_score - 0.8
    ElseIf x_pos > pmaxx Then
        x_pos = pmaxx
        tmp_score = tmp_score - 0.8
    End If
    
    If y_pos < pminy Then
        y_pos = pminy
        tmp_score = tmp_score - 0.8
    ElseIf y_pos > pmaxy Then
        y_pos = pmaxy
        tmp_score = tmp_score - 0.8
    End If
    
    'Penalize running into blocks
    For i = 1 To pn_block
        If x_pos = pBlockPos(1, i) And y_pos = pBlockPos(2, i) Then
            x_pos = x_pos_old
            y_pos = y_pos_old
            tmp_score = tmp_score - 0.75
            Exit For
        End If
    Next i
    
    'Penalize re-visiting positions
    isVisited = False
    For i = 1 To pn_visited
        If x_pos = pVisitedPos(1, i) And y_pos = pVisitedPos(2, i) Then
            tmp_score = tmp_score - 0.25
            isVisited = True
            Exit For
        End If
    Next i
    
    'Award new discoveries
    If Not isVisited Then
        pn_visited = pn_visited + 1
        ReDim Preserve pVisitedPos(1 To 2, 1 To pn_visited)
        pVisitedPos(1, pn_visited) = x_pos
        pVisitedPos(2, pn_visited) = y_pos
        tmp_score = tmp_score + 0.5
    End If
    
    pPlayerPos(1, 1) = x_pos
    pPlayerPos(2, 1) = y_pos
    
    pStates((x_pos_old - 1) * 7 + y_pos_old, 1) = 0
    pStates((x_pos - 1) * 7 + y_pos, 1) = 1
    
End Sub


Private Sub TakeStep_MazeII(x_action As Long, tmp_score As Double, isHit As Boolean)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, ii As Long, jj As Long
Dim x_pos As Long, y_pos As Long
Dim x_pos_old As Long, y_pos_old As Long
Dim tmp_x As Double
Dim isVisited As Boolean

    tmp_score = 0
 
    x_pos = pPlayerPos(1, 1)
    y_pos = pPlayerPos(2, 1)
    x_pos_old = x_pos
    y_pos_old = y_pos
    
    If x_action = 1 Then
        y_pos = y_pos + 1
    ElseIf x_action = 2 Then
        x_pos = x_pos + 1
    ElseIf x_action = 3 Then
        y_pos = y_pos - 1
    ElseIf x_action = 4 Then
        x_pos = x_pos - 1
    End If

    If x_pos = pTgtPos(1, 1) And y_pos = pTgtPos(2, 1) Then
        tmp_score = tmp_score + 50
        isHit = True
        pPlayerPos(1, 1) = x_pos
        pPlayerPos(2, 1) = y_pos
        pStates(1, 1) = x_pos / pmaxx
        pStates(2, 1) = y_pos / pmaxy
        pStates(3, 1) = x_pos_old / pmaxx
        pStates(4, 1) = y_pos_old / pmaxy
        Call scan_obstacles(pStates)
        Exit Sub
    End If

    'Penalize long wandering
    tmp_score = tmp_score - 0.04
    
    'Penalize running into walls
    If x_pos < pminx Then
        x_pos = pminx
        tmp_score = tmp_score - 0.8
    ElseIf x_pos > pmaxx Then
        x_pos = pmaxx
        tmp_score = tmp_score - 0.8
    End If
    
    If y_pos < pminy Then
        y_pos = pminy
        tmp_score = tmp_score - 0.8
    ElseIf y_pos > pmaxy Then
        y_pos = pmaxy
        tmp_score = tmp_score - 0.8
    End If
    
    'Penalize running into blocks
    For i = 1 To pn_block
        If x_pos = pBlockPos(1, i) And y_pos = pBlockPos(2, i) Then
            x_pos = x_pos_old
            y_pos = y_pos_old
            tmp_score = tmp_score - 0.75
            Exit For
        End If
    Next i
    
    'Penalize re-visiting positions
    isVisited = False
    For i = 1 To pn_visited
        If x_pos = pVisitedPos(1, i) And y_pos = pVisitedPos(2, i) Then
            tmp_score = tmp_score - 0.25
            isVisited = True
            Exit For
        End If
    Next i
    
    'Award new discoveries
    If Not isVisited Then
        pn_visited = pn_visited + 1
        ReDim Preserve pVisitedPos(1 To 2, 1 To pn_visited)
        pVisitedPos(1, pn_visited) = x_pos
        pVisitedPos(2, pn_visited) = y_pos
        tmp_score = tmp_score + 0.5
    End If
    
    pPlayerPos(1, 1) = x_pos
    pPlayerPos(2, 1) = y_pos
    
    pStates(1, 1) = x_pos / pmaxx
    pStates(2, 1) = y_pos / pmaxy
    pStates(3, 1) = x_pos_old / pmaxx
    pStates(4, 1) = y_pos_old / pmaxy

    Call scan_obstacles(pStates)
    
End Sub


Private Sub scan_obstacles(x_states)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, ii As Long, jj As Long

    For i = 5 To 20
        pStates(i, 1) = 0
    Next i
    'pStates(5:12,1) shows whether that position is blocked
    If pPlayerPos(1, 1) = pminx Then
        pStates(5, 1) = 1
        pStates(8, 1) = 1
        pStates(10, 1) = 1
    End If
    If pPlayerPos(1, 1) = pmaxx Then
        pStates(7, 1) = 1
        pStates(9, 1) = 1
        pStates(12, 1) = 1
    End If
    If pPlayerPos(2, 1) = pminy Then
        pStates(5, 1) = 1
        pStates(6, 1) = 1
        pStates(7, 1) = 1
    End If
    If pPlayerPos(2, 1) = pmaxy Then
        pStates(10, 1) = 1
        pStates(11, 1) = 1
        pStates(12, 1) = 1
    End If
    For k = 1 To pn_block
        ii = pBlockPos(1, k): jj = pBlockPos(2, k)
        m = 4
        For j = -1 To 1
            For i = -1 To 1
                If Not (i = 0 And j = 0) Then
                    m = m + 1
                    If (pPlayerPos(1, 1) + i) = ii And (pPlayerPos(2, 1) + j) = jj Then
                        pStates(m, 1) = 1
                        Exit For
                    End If
                End If
            Next i
        Next j
    Next k
    
    'pStates(13:20,1) shows whether that position was visited before
    For k = 1 To pn_visited
        ii = pVisitedPos(1, k): jj = pVisitedPos(2, k)
        m = 12
        For j = -1 To 1
            For i = -1 To 1
                If Not (i = 0 And j = 0) Then
                    m = m + 1
                    If (pPlayerPos(1, 1) + i) = ii And (pPlayerPos(2, 1) + j) = jj Then
                        pStates(m, 1) = 1
                        Exit For
                    End If
                End If
            Next i
        Next j
    Next k
End Sub


Private Function TrainPlayerValueLayer(cMario As cPlayer, _
                Optional n_cycle As Long = 10, Optional n_episode As Long = 10, Optional n_epoch As Long = 50, Optional mini_batch As Long = 10, _
                Optional maxTimeStep As Long = -1, _
                Optional learn_rate_value As Double = 0.0001, _
                Optional discount_rate As Double = 0.95)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim i_episode As Long, i_cycle As Long
Dim x_score_avg As Double, n_T_avg As Double, winRate As Double
Dim step_size_value As Double
Dim episode_count As Long, converge_count As Long
Dim v As Variant
Dim x_score_avg_prv As Double, x_score As Double
Dim x_progress() As Double, value_progress() As Double
Dim v_err As Double, v_err_prv As Double
    

    converge_count = 0
    ReDim value_progress(1 To 2, 1 To 1)
    For i_cycle = 1 To n_cycle
        
        step_size_value = learn_rate_value * (1 - (i_cycle - 1) / n_cycle)
        
        'Generate multiple episodes of experiences
        For i_episode = 1 To n_episode
            
            episode_count = episode_count + 1
            If episode_count Mod 10 = 0 Then
                DoEvents
                Application.StatusBar = "TrainPlayer: Value Cycle: gererating episodes: " & i_cycle & "/" & n_cycle & ":" & i_episode & "/" & n_episode & _
                                            "...score=" & Round(x_score, 3)
            End If
            
            VBA.Randomize
            v = SimulateSingleEpsiode(cMario, maxTimeStep:=maxTimeStep, storeHist:=True, isGreedy:=False)
            
            Call cMario.SaveEpisode(v(3))
            
        Next i_episode
        
        DoEvents
        Application.StatusBar = "TrainPlayer: Learning Value only: " & i_cycle & "/" & n_cycle & "..."
        With cMario
            v_err = .LearnValueOnly(step_size_value, discount_rate, n_epoch, mini_batch)
            Call .ClearEpisode
            Call .ClearHist
        End With
        
        If i_cycle > 1 Then v_err = v_err_prv + 0.1 * (v_err - v_err_prv)
        
        ReDim Preserve value_progress(1 To 2, 1 To i_cycle)
        value_progress(1, i_cycle) = i_cycle
        value_progress(2, i_cycle) = v_err
        
        If i_cycle > 1 Then
            If (v_err <= v_err_prv And (v_err_prv - v_err) < (0.01 * v_err_prv)) Then
                converge_count = converge_count + 1
            Else
                converge_count = 0
            End If
        End If
        If converge_count >= 5 Then Exit For
        
        v_err_prv = v_err
        
        Debug.Print "Value Cycle: " & i_cycle & ", value error=" & Round(v_err, 6)
    Next i_cycle
    
    TrainPlayerValueLayer = value_progress
    
End Function


'Train player within this environment, player needs to be initialized first
'returns a double array of size(1:4,1:n_epoch), which stores the epoch counter,
'average scores, average number of steps, and win rate before each epoch
'Explaining n_cycle, n_episode, n_epoch and minibatch:
'During each cycle, n_episode of games are played and stored in player's buffer, who
'then learns from experience with n_epoch of mini-batch SGD, where mini_batch is the number of batches
'maxTimeStep        maxinum number of steps to let a game run, by default it uses the game specific values
'value_progress     if provided, will force player to run a learning process for the value layer only, after running the main learning cycles.
'                   value_progress will then store the training progress of value layer to check for convergence
'learn_rate         learning rate for policy network
'learn_rate_value   learning rate for value network
'discount_rate      discount rate used in discounting rewards

Function TrainPlayer(cMario As cPlayer, _
                Optional n_cycle As Long = 10, Optional n_episode As Long = 10, Optional n_epoch As Long = 50, Optional mini_batch As Long = 10, _
                Optional maxTimeStep As Long = -1, _
                Optional learn_rate As Double = 0.0001, Optional learn_rate_value As Double = 0.0001, _
                Optional discount_rate As Double = 0.95, _
                Optional value_progress As Variant, Optional learn_value_first As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim i_episode As Long, i_cycle As Long
Dim x_score_avg As Double, n_T_avg As Double, winRate As Double
Dim step_size As Double, step_size_value As Double
Dim episode_count As Long, converge_count As Long
Dim v As Variant
Dim x_score_avg_prv As Double, x_score As Double, x_score_sd As Double, x_score_sd_prv As Double
Dim x_progress() As Double
    
    If learn_value_first Then
        v = TrainPlayerValueLayer(cMario, _
                            n_cycle:=n_cycle, n_episode:=n_episode, _
                            n_epoch:=n_epoch, mini_batch:=mini_batch, _
                            maxTimeStep:=maxTimeStep, _
                            learn_rate_value:=learn_rate_value, _
                            discount_rate:=discount_rate)
                            
        'if input learn_rate is zero, that means we only want to learn the value layer with current policy
        If learn_rate = 0 Then
            TrainPlayer = v
            Exit Function
        End If
                            
    End If
    
    converge_count = 0
    episode_count = 0
    x_score = 0
    x_score_avg_prv = Exp(70)
    x_score_sd_prv = 0
    x_score_sd = 0
    ReDim x_progress(1 To 4, 1 To 1)
    
    For i_cycle = 1 To n_cycle

        'Generate multiple episodes of experiences
        x_score_avg = 0
        x_score_sd = 0
        n_T_avg = 0
        winRate = 0
        For i_episode = 1 To n_episode

            episode_count = episode_count + 1
            If episode_count Mod 10 = 0 Then
                DoEvents
                Application.StatusBar = "TrainPlayer: gererating episodes: " & i_cycle & "/" & n_cycle & ":" & i_episode & "/" & n_episode & _
                                            "...score=" & Round(x_score, 3)
            End If

            VBA.Randomize
            v = SimulateSingleEpsiode(cMario, maxTimeStep:=maxTimeStep, storeHist:=True, isGreedy:=False)

            x_score = x_score + 0.1 * (v(1) - x_score)
            x_score_avg = x_score_avg + v(1)
            x_score_sd = x_score_sd + v(1) ^ 2
            n_T_avg = n_T_avg + v(2)
            winRate = winRate + IIf(v(3) = 1, 1, 0)

            Call cMario.SaveEpisode(v(3))

        Next i_episode

        x_score_avg = x_score_avg / n_episode
        x_score_sd = (x_score_sd / n_episode - x_score_avg ^ 2) * n_episode / (n_episode - 1)
        If (x_score_sd > 0) Then
            x_score_sd = Sqr(x_score_sd)
        Else
            x_score_sd = 0
        End If
        n_T_avg = n_T_avg / n_episode
        winRate = winRate / n_episode

        ReDim Preserve x_progress(1 To 4, 1 To i_cycle)
        x_progress(1, i_cycle) = i_cycle
        x_progress(2, i_cycle) = x_score_avg
        x_progress(3, i_cycle) = n_T_avg
        x_progress(4, i_cycle) = winRate

        Debug.Print i_cycle & ", score=" & Round(x_score_avg, 4) & "(sd " & Round(x_score_sd, 4) & "), steps=" & Round(n_T_avg, 1) & ", winRate=" & Round(winRate, 4)

        'check for convergence
        If i_cycle > 1 And x_score_avg >= (x_score_avg_prv - 0.5 * x_score_sd_prv) And winRate > 0.95 Then
            converge_count = converge_count + 1
        Else
            converge_count = 0
        End If
        If converge_count >= 5 Then
            With cMario
                Call .ClearEpisode
                Call .ClearHist
            End With
            Exit For
        End If

        x_score_avg_prv = x_score_avg
        x_score_sd_prv = x_score_sd

        'Learn from episodes
        DoEvents
        Application.StatusBar = "TrainPlayer: Learning: " & i_cycle & "/" & n_cycle & "..."
        With cMario
            Call .Learn(learn_rate, learn_rate_value, discount_rate, n_epoch, mini_batch)
            Call .ClearEpisode
            Call .ClearHist
        End With

    Next i_cycle
    
    If Not IsMissing(value_progress) Then
        m = n_cycle
        If i_cycle < m Then m = i_cycle
        value_progress = TrainPlayerValueLayer(cMario, _
                                n_cycle:=m, n_episode:=n_episode, _
                                n_epoch:=n_epoch, mini_batch:=mini_batch, _
                                maxTimeStep:=maxTimeStep, _
                                learn_rate_value:=learn_rate_value, _
                                discount_rate:=discount_rate)
    End If

    TrainPlayer = x_progress
    Application.StatusBar = False
End Function


'Run the game multiple times from start to finish
'returns a length 3 vector storing average scores, average number of steps and win rate
'cMario         a cPlayer class object, needs to be first initialized to match the game inputs and outputs
'n_run          number of games to play
'maxTimeStep    maximum number of steps to play for each game, use game specific values by default
'storeHist      whether trajectories need to be stored for training purpose
'isGreedy       whether greedy actions should be taken
'strDisplaySheet    if supplied, will display game states on the worksheet every time step
'strChtName         if supplied together with strDisplaysheet, will refresh game states on the chart every time step
'FrameRate      how many frames to show per second. Needs to be adjusted manually, depends on PC speed and Excel version
'perfectPlay    for testing only, not implemented here anymore
Function Simulate(cMario As cPlayer, _
                Optional n_run As Long = 1, _
                Optional maxTimeStep As Long = -1, _
                Optional storeHist As Boolean = False, Optional isGreedy As Boolean = False, _
                Optional strDisplaySheet As Worksheet = Nothing, _
                Optional strChtName As String = "Chart 1", _
                Optional FrameRate As Double = 20, _
                Optional perfectPlay As Boolean = False) As Variant
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim n_steps As Long
Dim x_score As Double, winRate As Double
Dim v As Variant
    
    Simulate = CVErr(xlErrNA)
    x_score = 0
    n_steps = 0
    For i = 1 To n_run
        v = SimulateSingleEpsiode(cMario, maxTimeStep:=maxTimeStep, _
                                storeHist:=storeHist, isGreedy:=isGreedy, _
                                strDisplaySheet:=strDisplaySheet, _
                                strChtName:=strChtName, _
                                FrameRate:=FrameRate, perfectPlay:=perfectPlay)
        x_score = x_score + v(1)
        n_steps = n_steps + v(2)
        If v(3) = 1 Then
            winRate = winRate + 1
        End If
        
        If storeHist Then
            Call cMario.SaveEpisode(v(3))
        End If
        
    Next i
    
    ReDim v(1 To 3)
    v(1) = x_score / n_run
    v(2) = n_steps / n_run
    v(3) = winRate / n_run
    Simulate = v

End Function


'Perform a single run of the game from start to finish, this is wrapped under Simulate() so
'there should be no need to call this function directly outside this class
Private Function SimulateSingleEpsiode(cMario As cPlayer, Optional maxTimeStep As Long = -1, _
                            Optional storeHist As Boolean = False, Optional isGreedy As Boolean = False, _
                            Optional strDisplaySheet As Worksheet = Nothing, Optional strChtName As String = "Chart 1", _
                            Optional FrameRate As Double = 20, Optional perfectPlay As Boolean = False) As Variant
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long, n_max As Long
Dim tmp_x As Double, tmp_y As Double
Dim isDisplay As Boolean
Dim mycht As Chart
Dim x_states() As Double
Dim x_action As Long
Dim tmp_score As Double, x_score As Double
Dim isHit As Boolean, isDead As Boolean
Dim n_steps As Long
Dim output_tuple As Variant

    SimulateSingleEpsiode = VBA.CVErr(xlErrNA)
    
    'Initialize environment
    Call Reset
    
    'Use game default max steps if not specified
    If maxTimeStep <= 0 Then
        n_max = pmax_step
    Else
        n_max = maxTimeStep
    End If
    
    'Initialize scores
    n_steps = 0
    x_score = 0
    tmp_score = 0
    isHit = False
    isDead = False
    
    'Show current state on screen
    isDisplay = Not strDisplaySheet Is Nothing
    If isDisplay Then
        Application.ScreenUpdating = True
        Set mycht = strDisplaySheet.ChartObjects(strChtName).Chart
        Call DisplayOnScreen(strDisplaySheet, mycht, pStates, FrameRate, x_score, isHit, isDead, n_steps)
        'Call OutputFrame(strDisplaySheet, 0)
    End If
    
    'Start playing game
    For iterate = 1 To n_max
        
        'let player perform actions
        x_action = cMario.Interact(pStates, storeHist:=storeHist, isGreedy:=isGreedy)
        n_steps = n_steps + 1
        
        'Update states according to action
        If pstrGame = "MAZE" Then
            Call TakeStep_Maze(x_action, tmp_score, isHit)
        ElseIf pstrGame = "MAZEII" Then
            Call TakeStep_MazeII(x_action, tmp_score, isHit)
        ElseIf pstrGame = "STOMP" Then
            Call TakeStep_STOMP(x_action, tmp_score, isHit, isDead)
        End If
        
        'assign reward for this action, tmp_score is reward for this time step, x_score is the cumulative reward
        x_score = x_score + tmp_score
        If storeHist Then Call cMario.collect_reward(tmp_score)

        'Show current state on screen
        If isDisplay Then
            Call DisplayOnScreen(strDisplaySheet, mycht, pStates, FrameRate, x_score, isHit, isDead, n_steps)
            'Call OutputFrame(strDisplaySheet, iterate)
        End If
        
        'Terminate upon death, success or when steps exhausted
        If isHit Or isDead Then
            Exit For
        End If
    
    Next iterate

    'Return score as output
    ReDim output_tuple(1 To 3)
    output_tuple(1) = x_score
    output_tuple(2) = n_steps
    output_tuple(3) = IIf(isHit, 1, IIf(isDead, -1, 0))
    SimulateSingleEpsiode = output_tuple
    
End Function


'Game specific gimmick to shown states on screen
Private Sub DisplayOnScreen(mysht As Variant, mycht As Chart, x_states() As Double, FrameRate As Double, _
                            x_score As Double, isHit As Boolean, isDead As Boolean, n_step As Long)
Dim strtmp As String

    If UCase(pstrGame) = "MAZE" Or UCase(pstrGame) = "MAZEII" Then
    
        With mysht
            .Range("B5").Value = pPlayerPos(1, 1)
            .Range("C5").Value = pPlayerPos(2, 1)
            .Range("H5").Value = pTgtPos(1, 1)
            .Range("I5").Value = pTgtPos(2, 1)
            strtmp = ""
            If isHit Then strtmp = "WIN! " & n_step & " moves."
            If isDead Then strtmp = "LOSE! " & n_step & " moves."
            .Range("M5").Value = strtmp
            mycht.Refresh
            DoEvents
            Call temp_wait(1 / FrameRate)
            If isHit Or isDead Then Call temp_wait(2 / FrameRate)
        End With
        
    ElseIf UCase(pstrGame) = "STOMP" Then
    
        With mysht
            .Range("E5").Value = pTgtPos(1, 1)
            .Range("C5").Value = pPlayerPos(2, 1)
            strtmp = ""
            If isHit Then strtmp = "WIN ! Score " & IIf(x_score > 0, "+", "") & Round(x_score, 2)
            If isDead Then strtmp = "DEAD! Score " & IIf(x_score > 0, "+", "") & Round(x_score, 2)
            .Range("J5").Value = strtmp
            mycht.Refresh
            DoEvents
            DoEvents
            Call temp_wait(1 / FrameRate)
            If isHit Or isDead Then Call temp_wait(2 / FrameRate)
        End With
        
    End If
End Sub


'Temp work to help output frames as jpg to create animated gif
'Copied from the spreadsheetguru
' https://www.thespreadsheetguru.com/blog/vba-save-as-picture-file-excel
Private Sub OutputFrame(mysht As Variant, i_step As Long)
Dim strtmp As String
Dim cht As ChartObject
Dim ActiveShape As Shape
        
    'Copy/Paste Cell Range as a Picture
    mysht.Range("K5:Q25").Copy
    mysht.Pictures.Paste(link:=False).Select
    Set ActiveShape = mysht.Shapes(ActiveWindow.Selection.Name)
      
    'Create a temporary chart object (same size as shape)
    Set cht = mysht.ChartObjects.Add( _
            Left:=ActiveCell.Left, _
            Width:=ActiveShape.Width, _
            Top:=ActiveCell.Top, _
            Height:=ActiveShape.Height)
    
    'Format temporary chart to have a transparent background
    cht.ShapeRange.Fill.Visible = msoFalse
    cht.ShapeRange.Line.Visible = msoFalse
        
    'Copy/Paste Shape inside temporary chart
    ActiveShape.Copy
    cht.Activate
    ActiveChart.Paste
      
    'Save chart to User's Desktop as PNG File
    cht.Chart.Export ThisWorkbook.Path & "\TempImg\" & Format(i_step, "00") & ".jpg"
    
    'Delete temporary Chart
    cht.Delete
    ActiveShape.Delete
    
    'Re-Select Shape (appears like nothing happened!)
    'ActiveShape.Select

End Sub


'number of seconds to wait
Private Sub temp_wait(x_time As Double)
Dim i As Long
Dim x As Double, y As Double
    x = Timer
    Do While (Timer - x) < x_time
        DoEvents
    Loop
End Sub

