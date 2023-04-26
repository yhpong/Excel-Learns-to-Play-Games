VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===========================================================================
'2023-04-25
'This is the agent which has a policy layer for action, and a
'separate value layer for learning state value function. Some global
'variables store history as it interacts with the environment, these
'include: xt(), yt(), x_actiont(), probs(), vt(), px_rewards(), pEpisodeIdx(), pEpisodeStatus
'Learn() is the function to learn from episodes saved in memory. PPO is implemented and
'ADAM is used in all optimizations
'Interact() is the main interface for it to take an action in an environment
'collect_reward() is antoher interface for it to collect a single reward from the
'environment, which is a scalar append to the vector px_rewards()
'Init() statement is game specific as each game requires specific input and output size, and
'possibly different network structures
'PrintNetwork() and ReadNetwork() are used to print and read both the policy and value networks to a worksheet
'StateMatrix() is a case by case function that outputs the matrix of state vs action
'and relevant data inside, but this is not alwasy meaningful to be presented in 2D matrix.
'===========================================================================

Private cNNPolicy As cNeuralNet
Private pn_input As Long, pn_hidden As Long, pn_output As Long
Private xt() As Double, yt() As Double, x_actiont() As Long, probs() As Double
Private px_rewards() As Double, pdiscounted_rewards() As Double

Private pstrGame As String
Private cNNValue As cNeuralNet
Private pValue_n_hidden As Long
Private vt() As Double

Private pn_episode As Long
Private pEpisodeIdx() As Long, pEpisodeStatus() As Long

Private pdiscount_rate As Double


Sub Init(strGame)
Dim i As Long, j As Long
    
    pstrGame = UCase(Trim(strGame))
    
    If UCase(Trim(strGame)) = "MAZE" Then
        
        pdiscount_rate = 0.9
        
        pn_input = 7 * 7
        pn_hidden = 16
        pn_output = 4
        Set cNNPolicy = New cNeuralNet
        With cNNPolicy
            Call .Init(pn_input, pn_output, "SOFTMAX")
            Call .AddLayer(pn_hidden, "RELU")
        End With
        
        pValue_n_hidden = 16
        Set cNNValue = New cNeuralNet
        With cNNValue
            Call .Init(pn_input, 1, "LINEAR")
            Call .AddLayer(pValue_n_hidden, "RELU")
        End With
    
    ElseIf UCase(Trim(strGame)) = "MAZEII" Then
        
        pdiscount_rate = 0.9
        
        pn_input = 20
        pn_hidden = 16
        pn_output = 4
        Set cNNPolicy = New cNeuralNet
        With cNNPolicy
            Call .Init(pn_input, pn_output, "SOFTMAX")
            Call .AddLayer(pn_hidden, "RELU")
        End With
        
        pValue_n_hidden = 16
        Set cNNValue = New cNeuralNet
        With cNNValue
            Call .Init(pn_input, 1, "LINEAR")
            Call .AddLayer(pValue_n_hidden, "RELU")
        End With
        
    
    ElseIf UCase(Trim(strGame)) = "STOMP" Then
        
        pdiscount_rate = 0.9
        
        pn_input = 5
        pn_hidden = 16
        pn_output = 1
        Set cNNPolicy = New cNeuralNet
        With cNNPolicy
            Call .Init(pn_input, pn_output, "SIGMOID")
            Call .AddLayer(pn_hidden, "RELU")
        End With
        
        pValue_n_hidden = 16
        Set cNNValue = New cNeuralNet
        With cNNValue
            Call .Init(pn_input, 1, "LINEAR")
            Call .AddLayer(pValue_n_hidden, "RELU")
        End With

    End If
End Sub


Sub PrintNetwork(mysht As Worksheet)
Dim i As Long, j As Long, m As Long
    Call cNNPolicy.PrintNetwork(mysht)
    Call cNNValue.PrintNetwork(mysht, cNNPolicy.n_row + 1)
End Sub


Sub ReadNetwork(mysht As Worksheet)
Dim i As Long, j As Long, m As Long

    Set cNNPolicy = New cNeuralNet
    Set cNNValue = New cNeuralNet
    Call cNNPolicy.ReadNetwork(mysht)
    Call cNNValue.ReadNetwork(mysht, cNNPolicy.n_row + 1)
   
    pn_input = cNNPolicy.n_input
    pn_output = cNNPolicy.n_output

End Sub

Property Let discount_rate(x As Double)
    pdiscount_rate = x
End Property

'Total number of steps taken so far by policy network
Property Get n_hist() As Long
    n_hist = cNNPolicy.n_hist
End Property

'Collect a single reward and append it to px_rewards()
Sub collect_reward(x As Double)
Dim i As Long
    If ArrayIsEmpty(px_rewards) Then
        i = 0
        ReDim px_rewards(1 To 1)
    Else
        i = UBound(px_rewards, 1)
        ReDim Preserve px_rewards(1 To i + 1)
    End If
    px_rewards(i + 1) = x
End Sub

'Return total reward if m is not supplied, otherwise return reward from specified time step
Property Get x_reward(Optional m As Long = -1) As Double
Dim i As Long, tmp_x As Double
    If m > 0 Then
        x_reward = px_rewards(m)
    Else
        tmp_x = 0
        For i = 1 To UBound(px_rewards)
            tmp_x = tmp_x + px_rewards(i)
        Next i
        x_reward = tmp_x
    End If
End Property

'clear all saved rewards
Sub reset_reward()
    Erase px_rewards
End Sub

'clear all experience buffer
Sub ClearHist()
    pn_episode = 0
    Call cNNPolicy.ClearHist
    Call cNNValue.ClearHist
    Erase xt, yt, vt, x_actiont
    Erase px_rewards, pdiscounted_rewards
    Erase pEpisodeIdx, pEpisodeStatus
End Sub

'clear all experience buffer, I forget why this is here since it's the same as ClearHist, buy anyway
'Private Sub Clear()
'    Erase xt, yt, vt, x_actiont
'    Erase px_rewards, pdiscounted_rewards
'    Call cNNPolicy.ClearHist
'    Call cNNValue.ClearHist
'End Sub

'Reset episode counter
Sub ClearEpisode()
    pn_episode = 0
    Erase pEpisodeIdx, pEpisodeStatus
End Sub


'record end of an episode by saving its end status, and also an integer pointer to
'time steps in the full trajectory that represents ending points
Sub SaveEpisode(endStatus As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
    pn_episode = pn_episode + 1
    If pn_episode = 1 Then
        ReDim pEpisodeIdx(1 To pn_episode)
        ReDim pEpisodeStatus(1 To pn_episode)
    Else
        ReDim Preserve pEpisodeIdx(1 To pn_episode)
        ReDim Preserve pEpisodeStatus(1 To pn_episode)
    End If
    pEpisodeIdx(pn_episode) = UBound(xt, 2)
    pEpisodeStatus(pn_episode) = endStatus
End Sub

'No longer using these
'Sub ADAM_Init()
'    Call cNNPolicy.ADAM_Init
'    Call cNNValue.ADAM_Init
'End Sub
'
'Sub ADAM_Clear()
'    Call cNNPolicy.ADAM_Clear
'    Call cNNValue.ADAM_Clear
'End Sub
'
'Sub NNWgtCache()
'    Call cNNPolicy.CacheCurrentWgt
'    Call cNNValue.CacheCurrentWgt
'End Sub
'
'Sub NNWgtCacheClear()
'    Call cNNPolicy.ClearCacheWgt
'    Call cNNValue.ClearCacheWgt
'End Sub
'
'Sub NNWgtRestore()
'    Call cNNPolicy.RestoreWgt
'    Call cNNValue.RestoreWgt
'End Sub
'
'Sub RMSClear()
'    Call cNNPolicy.RMSProp_Clear
'    Call cNNValue.RMSProp_Clear
'End Sub


'Learn from saved episodes
'bread and butter of this exercise, implemented PPO by Schulman 2017 (https://arxiv.org/abs/1707.06347v2)
'implementation detail is similar to https://iclr-blog-track.github.io/2022/03/25/ppo-implementation-details/
'for simplicity a fixed value of 0.2 is hardcode as the clipping parameter
'learn_rate         learning rate for policy network
'learn_rate_value   learning rate for value network
'discount_rate      discount rate for reward allocation
'n_epoch            number of epochs to run in optimization
'mini_batch         number of mini-batchs to divide the trainingg data into batches
Function Learn(learn_rate As Double, learn_rate_value As Double, discount_rate As Double, _
                n_epoch As Long, mini_batch As Long, _
                Optional entropy_coef As Double = 0.01)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_T As Long
Dim i_epoch As Long, batch_start As Long, n_mb As Long, batch_size As Long, n_converge As Long, n_step As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double, step_size As Double
Dim grad_out() As Double, grad_value() As Double
Dim discounted_rewards() As Double
Dim x_err As Double, x_err_prv As Double, delta As Double, i_episode As Long, i_action As Long
Dim r_mean As Double, r_sd As Double
Dim batch_idx() As Long, mini_batch_idx As Variant
Dim x_ratio() As Double, x_loss() As Double, x_loss_mean As Double, x_obj As Double, x_obj_prv As Double
Dim xt_old() As Double, xt_old_sub As Variant
Dim vt_old() As Double
Dim x_advantages() As Double, x_advantages_sub As Variant
Dim probs_old() As Double, probs_old_sub As Variant
Dim x_actiont_old_sub As Variant, x_actiont_old() As Long
Dim v As Variant
Dim x_entropy() As Double, x_entropy_mean As Double
    
    'estimate size of each batch based on number of mini-batches
    n_T = UBound(yt, 2)
    ReDim batch_idx(1 To n_T)
    For i = 1 To n_T
        batch_idx(i) = i
    Next i
    batch_size = Int(n_T / mini_batch)

    'Save old outputs
    ReDim xt_old(1 To pn_input, 1 To n_T)
    ReDim vt_old(1 To 1, 1 To n_T)
    ReDim probs_old(1 To n_T)
    ReDim x_actiont_old(1 To n_T)
    For i = 1 To n_T
        probs_old(i) = probs(i)
        x_actiont_old(i) = x_actiont(i)
        vt_old(1, i) = vt(1, i)
        For j = 1 To pn_input
            xt_old(j, i) = xt(j, i)
        Next j
    Next i
    
    'Reshape Rewards - this is game specific, see if this
    'can be removed in the future with better reward shaping algorithm
    If pstrGame = "STOMP" Then
        m = 1
        For i_episode = 1 To pn_episode
            n = pEpisodeIdx(i_episode)
            If pEpisodeStatus(i_episode) <> 0 Then
                tmp_x = px_rewards(n)
                px_rewards(n) = 0
                For i = n - 1 To m Step -1
                    If xt(5, i) = 0 Then
                        px_rewards(i) = tmp_x
                        Exit For
                    End If
                Next i
            End If
            m = n + 1
            If m > n_T Then Exit For
        Next i_episode
    End If
    
    'no longer needed, discounted rewards is calculated in LearnValueOnly() when needed
'    'calculate discounted rewards
'    pdiscounted_rewards = Calc_discounted_rewards(px_rewards, pEpisodeIdx, discount_rate)

    'calculate advantages
    'according to ICRL's blog, they also estimate the value from next states when an episode is not finished
    'but in this case I assume all epsiodes are finished and no rollout is performed
    'this could be something to improve in the future
    ReDim x_advantages(1 To n_T)
    m = 1
    For i_episode = 1 To pn_episode
        n = pEpisodeIdx(i_episode)
        x_advantages(n) = px_rewards(n) - vt(1, n) '+ discount_rate * EstimateValue(xt_tmp, False)
        For i = n - 1 To m Step -1
            delta = px_rewards(i) - vt(1, i) + discount_rate * vt(1, i + 1)
            x_advantages(i) = delta + discount_rate * 1 * x_advantages(i + 1)
        Next i
        m = n + 1
        If m > n_T Then Exit For
    Next i_episode

    'temp codes to output data for debugging
'    With ActiveWorkbook.Sheets("Stomp")
'        .Range("U2:AD200").Clear
'        n = pEpisodeIdx(1)
'        For i = 1 To n
'            .Range("U" & 1 + i).Value = i
'            .Range("V" & 1 + i).Value = xt(1, i)
'            .Range("W" & 1 + i).Value = xt(2, i)
'            .Range("X" & 1 + i).Value = xt(3, i)
'            .Range("Y" & 1 + i).Value = xt(4, i)
'            .Range("Z" & 1 + i).Value = x_actiont(i)
'            .Range("AA" & 1 + i).Value = px_rewards(i)
'            .Range("AB" & 1 + i).Value = pdiscounted_rewards(i)
'            .Range("AC" & 1 + i).Value = x_advantages(i)
'            .Range("AD" & 1 + i).Value = probs(i)
'        Next i
'    End With
'    End
    
    'Update value network
    If learn_rate_value > 0 Then
        x_err = LearnValueOnly(learn_rate_value, discount_rate, n_epoch, mini_batch)
    End If

    'mini batch gradient ascent to update policy network and maximize objective
    n_converge = 0
    x_obj_prv = -Exp(70)
    Call cNNPolicy.ClearHist
    Erase xt, vt, yt, probs, x_actiont
    For i_epoch = 1 To n_epoch
        
        'shuffle time steps
        Call Shuffle(batch_idx)
        x_obj = 0: n_step = 0
        batch_start = 1
        Do While batch_start < n_T

            'Extract a segment of trajectories, note that the last segment could be shorter than batch_size
            mini_batch_idx = SubsetIdx(batch_idx, batch_start, batch_start + batch_size - 1)
            n_mb = UBound(mini_batch_idx)
            batch_start = batch_start + batch_size 'start of next batch
            
            xt_old_sub = SubsetIdx(xt_old, x_idx:=mini_batch_idx)
            x_actiont_old_sub = SubsetIdx(x_actiont_old, x_idx:=mini_batch_idx)
            x_advantages_sub = SubsetIdx(x_advantages, x_idx:=mini_batch_idx)
            probs_old_sub = SubsetIdx(probs_old, x_idx:=mini_batch_idx)

            'Normalize advantage, note that this is done on mini-batch level, not on the full trajectory
            r_mean = 0: r_sd = 0
            For i = 1 To n_mb
                r_mean = r_mean + x_advantages_sub(i)
                r_sd = r_sd + x_advantages_sub(i) ^ 2
            Next i
            r_mean = r_mean / n_mb
            r_sd = (r_sd / n_mb - r_mean * r_mean) / (n_mb - 1 + 0.000001)
            If r_sd > 0 Then
                r_sd = Sqr(r_sd)
            Else
                r_sd = 1    'in  rare case when r_sd is zero, simply de-mean
            End If
            For i = 1 To n_mb
                x_advantages_sub(i) = (x_advantages_sub(i) - r_mean) / (r_sd + 0.000000001)
            Next i

            'Find probabilities of performing the same actions under new network
            v = Interact(xt_old_sub, storeHist:=True, isGreedy:=False, fix_action:=x_actiont_old_sub)
            ReDim x_ratio(1 To n_mb)
            For i = 1 To n_mb
                x_ratio(i) = probs(i) / probs_old_sub(i)
            Next i

            'calculate clipped surrogate objective and its gradient
            'I did this by hand, hope it's right!
            x_loss_mean = 0
            ReDim x_loss(1 To n_mb)
            ReDim grad_out(1 To pn_output, 1 To n_mb)
            For i = 1 To n_mb
                i_action = x_actiont_old_sub(i)
                tmp_z = x_advantages_sub(i) / probs_old_sub(i)
                tmp_x = x_ratio(i) * x_advantages_sub(i)
                If x_ratio(i) > 1.2 Then
                    tmp_y = 1.2 * x_advantages_sub(i)
                ElseIf x_ratio(i) < 0.8 Then
                    tmp_y = 0.8 * x_advantages_sub(i)
                Else
                    tmp_y = tmp_x
                End If
                If tmp_x <= tmp_y Then
                    x_loss(i) = tmp_x
                    
                    If pn_output = 1 Then
                    
                        If i_action = 1 Then
                            grad_out(1, i) = yt(1, i) * (1 - yt(1, i)) * tmp_z
                        Else
                            grad_out(1, i) = -yt(1, i) * (1 - yt(1, i)) * tmp_z
                        End If

                    Else
                    
                        For j = 1 To pn_output
                            If j = i_action Then
                                grad_out(j, i) = yt(i_action, i) * (1 - yt(j, i)) * tmp_z
                            Else
                                grad_out(j, i) = -yt(i_action, i) * yt(j, i) * tmp_z
                            End If
                        Next j
                    
                    End If
                Else
                    x_loss(i) = tmp_y
                End If
                x_loss_mean = x_loss_mean + x_loss(i)
            Next i
            x_loss_mean = x_loss_mean

            'Calculate Entropy bonus and its gradient
            'again it's done by hand, hope it's right!
            If entropy_coef > 0 Then
                x_entropy = calcProbEntropy(yt)
                x_entropy_mean = 0
                For i = 1 To n_mb
                    x_entropy_mean = x_entropy_mean + x_entropy(i)
                Next i
                x_entropy_mean = x_entropy_mean
                
                For i = 1 To n_mb
                    For j = 1 To pn_output
                        grad_out(j, i) = grad_out(j, i) - entropy_coef * yt(j, i) * (Log(yt(j, i)) + x_entropy(i))
                    Next j
                Next i
            End If
            
            'combined objective
            x_obj = x_obj + x_loss_mean + entropy_coef * x_entropy_mean
            n_step = n_step + n_mb

            'Update policy
            With cNNPolicy
                Call .Backward(yt, grad_out, , isdEdx:=True)
                Call .UpdateWgt(-learn_rate, useSpeedUp:="ADAM")  'negative for ascent instead of descent
                Call .ResetWgtChg
                Call .ClearHist
            End With
            Erase xt, vt, yt, probs, x_actiont
        Loop
        
        'Evalulate this epoch and test for convergence
        If (i_epoch - 1) Mod 10 = 0 Then
            DoEvents
            Debug.Print "cPlayer: Learn: PolicyLayer: Epoch " & i_epoch & "/" & n_epoch & ", Objective=" & Round(x_obj, 4)
        End If
        
        x_obj = x_obj / n_step
        If x_obj >= x_obj_prv And (x_obj - x_obj_prv) < (Abs(x_obj_prv) * 0.001) Then
            n_converge = n_converge + 1
        Else
            n_converge = 0
        End If
        If n_converge >= 5 Then
            Exit For
        End If
        x_obj_prv = x_obj
        
    Next i_epoch
    Call cNNPolicy.ADAM_Clear
    
    'output error in value layer, for no particular reasons
    Learn = x_err
    
End Function



'Update value network from saved episodes
'input definition is same as Learn()
Function LearnValueOnly(learn_rate As Double, discount_rate As Double, n_epoch As Long, mini_batch As Long)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_T As Long, total_steps As Long, n_converge As Long
Dim i_epoch As Long, batch_start As Long, n_mb As Long, batch_size As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double, step_size As Double
Dim grad_value() As Double
Dim discounted_rewards() As Double
Dim x_err As Double, x_err_prv As Double, delta As Double, i_episode As Long
Dim r_mean As Double, r_sd As Double
Dim batch_idx() As Long, mini_batch_idx As Variant
Dim xt_old() As Double, xt_old_sub As Variant
Dim vt_old() As Double
Dim v As Variant
Dim discounted_rewards_old() As Double, discounted_rewards_old_sub As Variant

    n_T = UBound(vt, 2)
    ReDim batch_idx(1 To n_T)
    For i = 1 To n_T
        batch_idx(i) = i
    Next i
    batch_size = Int(n_T / mini_batch)

    'calculate discounted rewards
    pdiscounted_rewards = Calc_discounted_rewards(px_rewards, pEpisodeIdx, discount_rate)

    'Compute starting error
    x_err_prv = 0
    For i = 1 To n_T
        x_err_prv = x_err_prv + (vt(1, i) - pdiscounted_rewards(i)) ^ 2
    Next i
    x_err_prv = x_err_prv / (2 * n_T)

    'Save old outputs
    ReDim xt_old(1 To pn_input, 1 To n_T)
    ReDim vt_old(1 To 1, 1 To n_T)
    ReDim discounted_rewards_old(1 To n_T)
    For i = 1 To n_T
        vt_old(1, i) = vt(1, i)
        discounted_rewards_old(i) = pdiscounted_rewards(i)
        For j = 1 To pn_input
            xt_old(j, i) = xt(j, i)
        Next j
    Next i

    'mini batch SGD to update value layer
    Call cNNValue.ClearHist
    Erase vt
    n_converge = 0
    For i_epoch = 1 To n_epoch
    
        Call Shuffle(batch_idx)
        
        x_err = 0: total_steps = 0
        batch_start = 1
        Do While batch_start < n_T
            
            'Extract a segment of trajectories
            mini_batch_idx = SubsetIdx(batch_idx, batch_start, batch_start + batch_size - 1)
            n_mb = UBound(mini_batch_idx)
            batch_start = batch_start + batch_size
            
            xt_old_sub = SubsetIdx(xt_old, x_idx:=mini_batch_idx)
            discounted_rewards_old_sub = SubsetIdx(discounted_rewards_old, x_idx:=mini_batch_idx)
            
            'Estimate value using current value layer
            v = EstimateValue(xt_old_sub, True)
            
            'calculate error in value estimates and its gradient
            ReDim grad_value(1 To 1, 1 To n_mb)
            For i = 1 To n_mb
                grad_value(1, i) = vt(1, i) - discounted_rewards_old_sub(i)
                x_err = x_err + (vt(1, i) - discounted_rewards_old_sub(i)) ^ 2
            Next i
            total_steps = total_steps + n_mb
            
            'Update weights
            With cNNValue
                Call .Backward(vt, grad_value)
                Call .UpdateWgt(learn_rate, useSpeedUp:="ADAM")
                Call .ResetWgtChg
                Call .ClearHist
            End With
            Erase vt
            
        Loop
        
        'Evaluate current epoch and test for convergence
        x_err = x_err / (2 * total_steps)
        
        If Abs(x_err_prv - x_err) < (Abs(x_err_prv) * 0.001) Then 'use a lenient criteria as value layer is less important
            n_converge = n_converge + 1
        Else
            n_converge = 0
        End If
        If n_converge >= 5 Then
            Exit For
        End If
        x_err_prv = x_err
        
        If (i_epoch - 1) Mod 10 = 0 Then
            DoEvents
            Debug.Print "cPlayer: LearnValueOnly: Epoch " & i_epoch & "/" & n_epoch & ", err=" & Round(x_err, 4) & ", converge count=" & n_converge
        End If
        
    Next i_epoch
    Call cNNValue.ADAM_Clear
    
    'Output error in value estimates
    LearnValueOnly = x_err
    
End Function


Private Function Calc_discounted_rewards(x_rewards() As Double, EpisodeIdx() As Long, discount_rate As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_episode As Long
Dim n_T As Long, i_episode As Long
Dim discounted_rewards() As Double

    n_T = UBound(x_rewards, 1)
    n_episode = UBound(EpisodeIdx, 1)
    ReDim discounted_rewards(1 To n_T)
    
    m = 1
    For i_episode = 1 To n_episode
        n = EpisodeIdx(i_episode)
        discounted_rewards(n) = x_rewards(n)
        For i = n - 1 To m Step -1
            discounted_rewards(i) = x_rewards(i) + discount_rate * discounted_rewards(i + 1)
        Next i
        m = n + 1
        If m > n_T Then Exit For
    Next i_episode
    
    Calc_discounted_rewards = discounted_rewards
    
End Function


Function EvaluateValueError(discount_rate As Double) As Double
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_T As Long
Dim tmp_x As Double
Dim y() As Double, v() As Double
Dim discounted_rewards() As Double
Dim x_err As Double, delta As Double, i_episode As Long

    n_T = UBound(xt, 2)
    ReDim discounted_rewards(1 To n_T)
    
    v = cNNValue.FwdPass(xt, storeOutput:=True)
    
    discounted_rewards = Calc_discounted_rewards(px_rewards, pEpisodeIdx, discount_rate)
    x_err = 0
    For k = 1 To n_T
        x_err = x_err + (v(1, k) - discounted_rewards(k)) ^ 2
    Next k
    x_err = x_err / (2 * n_T)
    
    EvaluateValueError = x_err
    
End Function


'feed in state variables x and output a single integer that represents the choice of actions
'in learning phase when storeHist is True, state value is also estimated by the value layer
'x,             input state variables of size (1:pn_input, 1:n)
'storeHist,     save all intermediate outputs for training purpose when set to True
'isGreedy,      greedy action when set to True, otherwise action is sample from the output discrete distribution
'fix_action,    used only in PPO, by supplying a series of actions together with states variables, it
'               calculates the probability of performing those actions (under current policy) and store them in memory
Function Interact(x As Variant, Optional storeHist As Boolean = False, Optional isGreedy As Boolean = False, Optional fix_action As Variant = Null) As Long
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_T As Long, iterate As Long, n_T_old As Long
Dim tmp_x As Double
Dim y() As Double, v As Variant
Dim x_action() As Long
    
    n_T = UBound(x, 2)
    
    y = cNNPolicy.FwdPass(x, storeOutput:=storeHist)
    
    If IsArray(fix_action) Then
        ReDim x_action(1 To n_T)
        For i = 1 To n_T
            x_action(i) = fix_action(i)
        Next i
    Else
        x_action = SampleAction(y, isGreedy)
    End If
    
    If storeHist Then
        If ArrayIsEmpty(xt) Then
            n_T_old = 0
            ReDim xt(1 To pn_input, 1 To n_T)
            ReDim yt(1 To pn_output, 1 To n_T)
            ReDim probs(1 To n_T)
            ReDim x_actiont(1 To n_T)
        Else
            n_T_old = UBound(xt, 2)
            ReDim Preserve xt(1 To pn_input, 1 To n_T_old + n_T)
            ReDim Preserve yt(1 To pn_output, 1 To n_T_old + n_T)
            ReDim Preserve probs(1 To n_T_old + n_T)
            ReDim Preserve x_actiont(1 To n_T_old + n_T)
        End If
        For iterate = 1 To n_T
            k = n_T_old + iterate
            
            For i = 1 To pn_input
                xt(i, k) = x(i, iterate)
            Next i
            
            For i = 1 To pn_output
                yt(i, k) = y(i, iterate)
            Next i
            
            If pn_output = 1 Then
                If x_action(iterate) = 1 Then
                    probs(k) = y(1, iterate)
                Else
                    probs(k) = 1 - y(1, iterate)
                End If
            ElseIf pn_output > 1 Then
                probs(k) = y(x_action(iterate), iterate)
            End If

            x_actiont(k) = x_action(iterate)
            
        Next iterate

    End If
    
    Interact = x_action(n_T)
    
    'Evaluate value
    If Not storeHist Then Exit Function
    v = EstimateValue(x, storeHist:=storeHist)
    
End Function


Function EstimateValue(x As Variant, Optional storeHist As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_T As Long, iterate As Long, n_T_old As Long
Dim tmp_x As Double
Dim v() As Double
    n_T = UBound(x, 2)
    v = cNNValue.FwdPass(x, storeOutput:=storeHist)
    If storeHist Then
        If ArrayIsEmpty(vt) Then
            n_T_old = 0
            ReDim vt(1 To 1, 1 To n_T)
        Else
            n_T_old = UBound(vt, 2)
            ReDim Preserve vt(1 To 1, 1 To n_T_old + n_T)
        End If
        For iterate = 1 To n_T
            vt(1, n_T_old + iterate) = v(1, iterate)
        Next iterate
    End If
    EstimateValue = v(1, n_T)
End Function






Function StateMatrix(Optional strOutType As String = "ACTION")
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim ii As Long, jj As Long
Dim x() As Double
Dim y() As Double, v As Variant
Dim v_out() As Double
Dim tmp_x As Double, tmp_y As Double

    If UCase(pstrGame) = "MAZE" Then
        
        ReDim v_out(1 To 7 + 1, 1 To 7 + 1)
        For ii = 1 To 7
            tmp_y = ii
            v_out(1 + (7 - ii + 1), 1) = tmp_y
            For jj = 1 To 7
                tmp_x = jj
                v_out(1, 1 + jj) = tmp_x
                ReDim x(1 To pn_input, 1 To 1)
                x((tmp_x - 1) * 7 + tmp_y, 1) = 1
                If strOutType = "ACTION" Then
                    y = cNNPolicy.FwdPass(x, storeOutput:=False)
                    v = SampleAction(y, True)
                    v_out(1 + (7 - ii + 1), 1 + jj) = v(1)
                ElseIf strOutType = "VALUE" Then
                    v_out(1 + (7 - ii + 1), 1 + jj) = cNNValue.FwdPass(x, storeOutput:=False)(1, 1)
                End If
            Next jj
        Next ii
    
    ElseIf UCase(pstrGame) = "MAZEII" Then
        
        ReDim v_out(1 To 7 + 1, 1 To 13 + 1)
        For ii = 1 To 7
            tmp_y = ii
            v_out(1 + (7 - ii + 1), 1) = tmp_y
            For jj = 1 To 13
                tmp_x = jj
                v_out(1, 1 + jj) = tmp_x
                ReDim x(1 To pn_input, 1 To 1)
                x(1, 1) = tmp_x / 13
                x(2, 1) = tmp_y / 7
                x(3, 1) = tmp_x / 13
                x(4, 1) = tmp_y / 7
                If strOutType = "ACTION" Then
                    y = cNNPolicy.FwdPass(x, storeOutput:=False)
                    v = SampleAction(y, True)
                    v_out(1 + (7 - ii + 1), 1 + jj) = v(1)
                ElseIf strOutType = "VALUE" Then
                    v_out(1 + (7 - ii + 1), 1 + jj) = cNNValue.FwdPass(x, storeOutput:=False)(1, 1)
                End If
            Next jj
        Next ii
    
    ElseIf UCase(pstrGame) = "STOMP" Then
        
        ReDim x(1 To pn_input, 1 To 1)
        ReDim v_out(1 To 30 + 1, 1 To 10 + 1)
        For ii = 1 To 30
            tmp_y = 20 - ii + 1
            v_out(1 + ii, 1) = tmp_y
            For jj = 1 To 10
                tmp_x = -1.5 + (jj - 1) * 0.7 / 9
                v_out(1, 1 + jj) = tmp_x
                x(1, 1) = tmp_y / 20
                x(2, 1) = tmp_x
                If strOutType = "ACTION" Then
                    y = cNNPolicy.FwdPass(x, storeOutput:=False)
                    v_out(1 + ii, 1 + jj) = y(1, 1)
                ElseIf strOutType = "VALUE" Then
                    v_out(1 + ii, 1 + jj) = cNNValue.FwdPass(x, storeOutput:=False)(1, 1)
                End If
            Next jj
        Next ii
        
    End If
    
    StateMatrix = v_out
End Function



Private Function ArrayIsEmpty(x) As Boolean
Dim i As Long
    If Not IsArray(x) Then
        ArrayIsEmpty = True
    Else
        ArrayIsEmpty = False
        On Error Resume Next
        i = UBound(x, 1)
        If Err.Number <> 0 Then
            Err.Clear
            ArrayIsEmpty = True
        End If
    End If
End Function


Private Function SampleAction(y() As Double, isGreedy As Boolean)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim n_output As Long, n_T As Long, n_T_old As Long
Dim tmp_max As Double, cumprob As Double
Dim x_action() As Long

    n_T = UBound(y, 2)
    n_output = UBound(y, 1)
    ReDim x_action(1 To n_T)
    If Not isGreedy Then
        
        
        If n_output = 1 Then
            For iterate = 1 To n_T
                If y(1, iterate) >= Rnd() Then
                    x_action(iterate) = 1
                End If
            Next iterate
        Else
            For iterate = 1 To n_T
                tmp_max = Rnd()
                cumprob = y(1, iterate)
                If tmp_max < cumprob Then
                    x_action(iterate) = 1
                Else
                    For i = 2 To n_output
                        cumprob = cumprob + y(i, iterate)
                        If tmp_max < cumprob Then
                            x_action(iterate) = i
                            Exit For
                        End If
                    Next i
                End If
            Next iterate
        End If
        
    Else
        
        If n_output = 1 Then
        
            For iterate = 1 To n_T
                If y(1, iterate) > 0.5 Then
                    x_action(iterate) = 1
                End If
            Next iterate
        Else
            For iterate = 1 To n_T
                tmp_max = y(1, iterate)
                x_action(iterate) = 1
                For i = 2 To n_output
                    If y(i, iterate) > tmp_max Then
                        tmp_max = y(i, iterate)
                        x_action(iterate) = i
                    End If
                Next i
            Next iterate
        End If
        
    End If
    
    SampleAction = x_action

End Function




Private Function SubsetIdx(x As Variant, Optional m As Long, Optional n As Long, Optional x_idx As Variant = Null)
Dim i As Long, j As Long, k As Long
Dim y As Variant, y_idx() As Long

    If IsArray(x_idx) Then
        k = UBound(x_idx, 1)
        If getDimension(x) = 1 Then
            ReDim y(1 To k)
            For i = 1 To k
                y(i) = x(x_idx(i))
            Next i
        ElseIf getDimension(x) = 2 Then
            m = UBound(x, 1)
            ReDim y(1 To m, 1 To k)
            For i = 1 To k
                For j = 1 To m
                    y(j, i) = x(j, x_idx(i))
                Next j
            Next i
        End If
        SubsetIdx = y
        Exit Function
    End If
    
    k = n
    If n > UBound(x, 1) Then
        k = UBound(x, 1)
    End If
    
    ReDim y_idx(1 To k - m + 1)
    For i = 1 To (k - m + 1)
        y_idx(i) = x(m + i - 1)
    Next i
    
    SubsetIdx = y_idx
    
End Function


Private Function calcProbEntropy(x_probs() As Double) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim n_output As Long, n_T As Long
Dim tmp_x As Double
Dim x_entropy() As Double, x_entropy_mean As Double

    n_output = UBound(x_probs, 1)
    n_T = UBound(x_probs, 2)
    ReDim x_entropy(1 To n_T)
    
    If n_output = 1 Then
        For iterate = 1 To n_T
            tmp_x = x_probs(i, iterate)
            tmp_x = -tmp_x * Log(tmp_x) - (1 - tmp_x) * Log(1 - tmp_x)
            x_entropy(iterate) = tmp_x
        Next iterate
    Else
        For iterate = 1 To n_T
            tmp_x = 0
            For i = 1 To n_output
                tmp_x = tmp_x - x_probs(i, iterate) * Log(x_probs(i, iterate))
            Next i
            x_entropy(iterate) = tmp_x
        Next iterate
    End If
    calcProbEntropy = x_entropy
End Function

Private Function getDimension(A As Variant) As Long
    Dim i As Long, j As Long
    i = 0
    On Error GoTo getDimension_Err:
    Do While True:
        i = i + 1
        j = UBound(A, i)
    Loop
getDimension_Err:
    getDimension = i - 1
End Function

Private Sub Shuffle(x() As Long)
Dim i As Long, j As Long, n As Long
Dim k As Long
Dim vtmp As Variant
    n = UBound(x)
    Randomize
    For i = n To 2 Step -1
        j = Int(Rnd() * i) + 1  'Random_Integer(1, i)
        vtmp = x(j)
        x(j) = x(i)
        x(i) = vtmp
    Next i
End Sub
