VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNeuralNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================
'2023-04-25
'This is a general purpose feed-forward neutral network
'Structure of the network needs to be initialized manually with Init() and AddLayer()
'The two main methods to fit data are Fit() and FitCV()
'either ADAM or RMSProp can be used to speed up convergence
'networks can be printed to a worksheet using PrintNetwork(), and read by ReadNetwork()
'A few frequently used functions like Label2Dummy(), Dummy2Label() and SampleSoftmax()
'are also included
'======================================================================================

Private pOutputType As String
Private pn_input As Long, pn_output As Long, pn_unit As Long
Private pn_layer As Long, pn_hidden() As Long, pLayerType() As String
Private vWgts() As Variant, vBias As Variant
Private pWOut() As Double, pbOut() As Double

Private vdWgts() As Variant, vdBias As Variant
Private dWOut() As Double, dbOut() As Double

Private vWgts_tmp() As Variant, vBias_tmp As Variant
Private pWOut_tmp() As Double, pbOut_tmp() As Double

Private vdW2() As Variant, vdB2() As Variant
Private dWOut2() As Double, dbOut2() As Double

Private vdW1() As Variant, vdB1() As Variant
Private dWOut1() As Double, dbOut1() As Double

Private vh_out() As Variant, px_in() As Double

Private pn_row As Long, pn_hist As Long
Private pTrainProgress() As Double
Private pADAM_count As Long, pRMS_count As Long


Property Get OutputType() As String
    OutputType = pOutputType
End Property

Property Get n_input() As Long
    n_input = pn_input
End Property

Property Get n_output() As Long
    n_output = pn_output
End Property

Property Get n_unit() As Long
    n_unit = pn_unit
End Property

Property Get n_layer() As Long
    n_layer = pn_layer
End Property

Property Get n_row() As Long
    n_row = pn_row
End Property

Property Get n_hist() As Long
    n_hist = pn_hist
End Property

Property Get WgtHidden()
    WgtHidden = vWgts
End Property

Property Get BiasHidden()
    BiasHidden = vBias
End Property

Property Get WOut()
    WOut = pWOut
End Property

Property Get bOut()
    bOut = pbOut
End Property

Property Get n_hidden()
    n_hidden = pn_hidden
End Property

Property Get LayerType()
    LayerType = pLayerType
End Property

Property Get Hist_x_in() As Double()
    Hist_x_in = px_in
End Property

Property Get TrainProgress()
    TrainProgress = pTrainProgress
End Property

'Clone properties from another network
Sub Clone(cNN As cNeuralNet)
    pOutputType = cNN.OutputType
    pn_input = cNN.n_input
    pn_output = cNN.n_output
    pn_unit = cNN.n_unit
    pn_layer = cNN.n_layer
    pn_hidden = cNN.n_hidden
    pLayerType = cNN.LayerType
    pWOut = cNN.WOut
    pbOut = cNN.bOut
    vWgts = cNN.WgtHidden
    vBias = cNN.BiasHidden
End Sub

Sub RMSProp_Clear()
    pRMS_count = 0
    Erase vdW2, vdB2, dWOut2, dbOut2
End Sub

Private Sub RMSProp_CalcRMS()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim n_input As Long, n_output As Long
Dim dWgt() As Double, dBias() As Double
Dim dW2() As Double, db2() As Double
Dim tmp_x As Double
    
    If pRMS_count = 0 Then
        ReDim vdW2(1 To pn_layer)
        ReDim vdB2(1 To pn_layer)
        ReDim dWOut2(1 To pn_output, 1 To pn_hidden(pn_layer))
        ReDim dbOut2(1 To pn_output)
        For iterate = 1 To pn_layer
            If iterate = 1 Then
                n_input = pn_input
            Else
                n_input = pn_hidden(iterate - 1)
            End If
            n_output = pn_hidden(iterate)
            ReDim dW2(1 To n_output, 1 To n_input)
            ReDim db2(1 To n_output)
            vdW2(iterate) = dW2
            vdB2(iterate) = db2
        Next iterate
    End If
    
    For iterate = 1 To pn_layer
        dWgt = vdWgts(iterate)
        dBias = vdBias(iterate)
        n_output = UBound(dBias, 1)
        n_input = UBound(dWgt, 2)
        
        If pRMS_count = 0 Then
            ReDim dW2(1 To n_output, 1 To n_input)
            ReDim db2(1 To n_output)
            For j = 1 To n_input
                For i = 1 To n_output
                    dW2(i, j) = dWgt(i, j) ^ 2
                Next i
            Next j

            For i = 1 To n_output
                db2(i) = dBias(i) ^ 2
            Next i

        Else
            dW2 = vdW2(iterate)
            db2 = vdB2(iterate)
            For j = 1 To n_input
                For i = 1 To n_output
                    dW2(i, j) = 0.9 * dW2(i, j) + 0.1 * dWgt(i, j) ^ 2
                Next i
            Next j
            For i = 1 To n_output
                db2(i) = 0.9 * db2(i) + 0.1 * dBias(i) ^ 2
            Next i
        End If

        vdW2(iterate) = dW2
        vdB2(iterate) = db2
        
    Next iterate

    If pRMS_count = 0 Then
        For j = 1 To pn_hidden(pn_layer)
            For i = 1 To pn_output
                dWOut2(i, j) = dWOut(i, j) ^ 2
            Next i
        Next j

        For i = 1 To pn_output
            dbOut2(i) = dbOut(i) ^ 2
        Next i
    Else
        For j = 1 To pn_input
            For i = 1 To pn_output
                dWOut2(i, j) = 0.9 * dWOut2(i, j) + 0.1 * dWOut(i, j) ^ 2
            Next i
        Next j
        
        For i = 1 To pn_output
            dbOut2(i) = 0.9 * dbOut2(i) + 0.1 * dbOut(i) ^ 2
        Next i
    End If

    pRMS_count = pRMS_count + 1

End Sub

'Clear and reset memories used in ADAM
Sub ADAM_Clear()
    pADAM_count = 0
    Erase vdW1, vdB1
    Erase vdW2, vdB2
    Erase dWOut1, dbOut1
    Erase dWOut2, dbOut2
End Sub

Private Sub ADAM_Init()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim n_input As Long, n_output As Long
Dim dWgt() As Double, dBias() As Double
Dim dW2() As Double, db2() As Double
Dim dW1() As Double, db1() As Double
    
    pADAM_count = 0
    
    ReDim vdW1(1 To pn_layer)
    ReDim vdB1(1 To pn_layer)
    ReDim dWOut1(1 To pn_output, 1 To pn_hidden(pn_layer))
    ReDim dbOut1(1 To pn_output)
    
    ReDim vdW2(1 To pn_layer)
    ReDim vdB2(1 To pn_layer)
    ReDim dWOut2(1 To pn_output, 1 To pn_hidden(pn_layer))
    ReDim dbOut2(1 To pn_output)
    
    For iterate = 1 To pn_layer
        dW1 = vWgts(iterate)
        n_output = UBound(dW1, 1)
        n_input = UBound(dW1, 2)
        
        ReDim dW1(1 To n_output, 1 To n_input)
        ReDim db1(1 To n_output)
        ReDim dW2(1 To n_output, 1 To n_input)
        ReDim db2(1 To n_output)
            
        vdW1(iterate) = dW1
        vdB1(iterate) = db1
        vdW2(iterate) = dW2
        vdB2(iterate) = db2
    Next iterate
    
End Sub

Private Sub ADAM_CalcMoment(Optional isFirst As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim n_input As Long, n_output As Long
Dim dWgt() As Double, dBias() As Double
Dim dW2() As Double, db2() As Double
Dim dW1() As Double, db1() As Double
Dim tmp_x As Double
    
    If pADAM_count = 0 Then Call ADAM_Init
    
    pADAM_count = pADAM_count + 1
    
    For iterate = 1 To pn_layer
        dWgt = vdWgts(iterate)
        dBias = vdBias(iterate)
        n_output = UBound(dWgt, 1)
        n_input = UBound(dWgt, 2)

        dW1 = vdW1(iterate)
        db1 = vdB1(iterate)
        dW2 = vdW2(iterate)
        db2 = vdB2(iterate)
        For j = 1 To n_input
            For i = 1 To n_output
                dW1(i, j) = 0.9 * dW1(i, j) + 0.1 * dWgt(i, j)
                dW2(i, j) = 0.999 * dW2(i, j) + 0.001 * dWgt(i, j) ^ 2
            Next i
        Next j
        
        For i = 1 To n_output
            db1(i) = 0.9 * db1(i) + 0.1 * dBias(i)
            db2(i) = 0.999 * db2(i) + 0.001 * dBias(i) ^ 2
        Next i

        vdW1(iterate) = dW1
        vdB1(iterate) = db1
        vdW2(iterate) = dW2
        vdB2(iterate) = db2
        
    Next iterate
    
    n_input = UBound(dWOut, 2)
    For j = 1 To n_input
        For i = 1 To pn_output
            dWOut1(i, j) = 0.9 * dWOut1(i, j) + 0.1 * dWOut(i, j)
            dWOut2(i, j) = 0.999 * dWOut2(i, j) + 0.001 * dWOut(i, j) ^ 2
        Next i
    Next j
    
    For i = 1 To pn_output
        dbOut1(i) = 0.9 * dbOut1(i) + 0.1 * dbOut(i)
        dbOut2(i) = 0.999 * dbOut2(i) + 0.001 * dbOut(i) ^ 2
    Next i

End Sub



'Cache current weights
Sub CacheCurrentWgt()
    vWgts_tmp = vWgts
    vBias_tmp = vBias
    pWOut_tmp = pWOut
    pbOut_tmp = pbOut
End Sub

'Clear cache weights
Sub ClearCacheWgt()
    Erase vWgts_tmp, vBias_tmp, pWOut_tmp, pbOut_tmp
End Sub

'Restore weights to cached values
Sub RestoreWgt()
    vWgts = vWgts_tmp
    vBias = vBias_tmp
    pWOut = pWOut_tmp
    pbOut = pbOut_tmp
End Sub

'clear intermediate outputs from memory
Sub ClearHist()
    pn_hist = 0
    Erase vh_out, px_in
End Sub

'Reset gradients
Sub ResetWgtChg()
    Erase dWOut, dbOut, vdWgts, vdBias
End Sub

'Initialize with one input layer and one output layer of specified activation function
Sub Init(n_input As Long, n_output As Long, Optional strOutputType As String = "SIGMOID")
Dim i As Long, j As Long
Dim tmp_x As Double

    pn_input = n_input
    pn_output = n_output
    pn_layer = 0
    pOutputType = UCase(Trim(strOutputType))
    
    If pn_output = 1 And pOutputType = "SOFTMAX" Then
        Debug.Print "cNeuralNet: Init: when output is binary, use SIGMOID instead of SOFTMAX"
        End
    End If
    
    VBA.Randomize
    
    tmp_x = Sqr(2 / pn_input)
    ReDim pWOut(1 To pn_output, 1 To pn_input)
    ReDim pbOut(1 To pn_output)
    For i = 1 To pn_output
        pbOut(i) = (-0.1 + Rnd() * 0.2) * tmp_x
        For j = 1 To pn_input
            pWOut(i, j) = (-0.5 + Rnd()) * tmp_x
        Next j
    Next i
    
    pn_unit = pn_output * (pn_input + 1)
    pn_hist = 0
    pADAM_count = 0
    pRMS_count = 0
    
End Sub


'Insert a hidden layer to current network, layer is inserted between the last hidden layer and the output layer
Sub AddLayer(n_hidden As Long, Optional strType As String = "SIGMOID")
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim xW() As Double, xb() As Double
Dim tmp_x As Double
    pn_layer = pn_layer + 1
    If pn_layer = 1 Then
        ReDim vWgts(1 To pn_layer)
        ReDim vBias(1 To pn_layer)
        ReDim pn_hidden(1 To pn_layer)
        ReDim pLayerType(1 To pn_layer)
        m = pn_input
    Else
        ReDim Preserve vWgts(1 To pn_layer)
        ReDim Preserve vBias(1 To pn_layer)
        ReDim Preserve pn_hidden(1 To pn_layer)
        ReDim Preserve pLayerType(1 To pn_layer)
        m = UBound(vBias(pn_layer - 1), 1)
    End If
    
    tmp_x = Sqr(2 / m)
    pn_hidden(pn_layer) = n_hidden
    pLayerType(pn_layer) = UCase(Trim(strType))
    ReDim xW(1 To n_hidden, 1 To m)
    ReDim xb(1 To n_hidden)
    For i = 1 To n_hidden
        xb(i) = (-0.1 + Rnd() * 0.2) * tmp_x
        For j = 1 To m
            xW(i, j) = (-0.5 + Rnd()) * tmp_x
        Next j
    Next i
    
    vWgts(pn_layer) = xW
    vBias(pn_layer) = xb
    
    'Reconnect output layer to new hidden layer
    tmp_x = Sqr(2 / n_hidden)
    ReDim pWOut(1 To pn_output, 1 To n_hidden)
    For i = 1 To pn_output
        pbOut(i) = (-0.1 + Rnd() * 0.2) * tmp_x
        For j = 1 To n_hidden
            pWOut(i, j) = (-0.5 + Rnd()) * tmp_x
        Next j
    Next i
    
    pn_unit = pn_output * (n_hidden + 1)
    For i = 1 To pn_layer
        If i > 1 Then
            pn_unit = pn_unit + pn_hidden(i) * (pn_hidden(i - 1) + 1)
        Else
            pn_unit = pn_unit + pn_hidden(i) * (pn_input + 1)
        End If
    Next i
    
End Sub


'Forward pass with input x
'Returns double array y() of dimension y(1:pn_output, 1:n), where n is same as length of x
'If storeOutput is set to True, intermediate activation and input x is stored
Function FwdPass(x, Optional storeOutput As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_len As Long, n_input As Long, n_hidden As Long
Dim iterate As Long
Dim xW() As Double, xb() As Double
Dim xin() As Double
Dim h() As Double, h_tmp() As Double
Dim v As Variant
Dim strType As String
Dim y() As Double
Dim n_hist_prv As Long

    n_len = UBound(x, 2)
    n_input = UBound(x, 1)
    If n_input <> pn_input Then
        Debug.Print "cNeuralNet:FwdPass: input size incorrect. n_input=" & n_input & ", pn_input=" & pn_input
        End
    End If
    
    n_hist_prv = pn_hist
    If storeOutput Then pn_hist = n_hist_prv + n_len
    
    'store output of each hidden layer
    If storeOutput And n_hist_prv = 0 Then
        ReDim vh_out(1 To pn_layer)
    End If
    
    ReDim xin(1 To n_input, 1 To n_len)
    For k = 1 To n_len
        For i = 1 To n_input
            xin(i, k) = x(i, k)
        Next i
    Next k
    
    'store input to be used later in back propagation
    If storeOutput Then
        If n_hist_prv = 0 Then
            ReDim px_in(1 To n_input, 1 To n_len)
            For k = 1 To n_len
                For i = 1 To n_input
                    px_in(i, k) = xin(i, k)
                Next i
            Next k
        Else
            ReDim Preserve px_in(1 To n_input, 1 To n_hist_prv + n_len)
            For k = 1 To n_len
                For i = 1 To n_input
                    px_in(i, n_hist_prv + k) = xin(i, k)
                Next i
            Next k
        End If
    End If
    
    For iterate = 1 To pn_layer

        xW = vWgts(iterate)
        xb = vBias(iterate)
        n_hidden = pn_hidden(iterate)
        strType = pLayerType(iterate)
        
        ReDim h(1 To n_hidden, 1 To n_len)
        v = WorksheetFunction.MMult(xW, xin)
        If n_hidden = 1 Then
            For k = 1 To n_len
                h(1, k) = f_Activate(v(k) + xb(1), strType)
            Next k
        Else
            For k = 1 To n_len
                For i = 1 To n_hidden
                    h(i, k) = f_Activate(v(i, k) + xb(i), strType)
                Next i
            Next k
        End If
        
'        ReDim h(1 To n_hidden, 1 To n_len)
'        For k = 1 To n_len
'            For i = 1 To pn_hidden
'                tmp_x = xb(i)
'                For j = 1 To n_input
'                    tmp_x = tmp_x + xW(i, j) * x(j, k)
'                Next j
'                h(i, k) = f_Activate(tmp_x, strType)
'            Next i
'        Next k
        
        If storeOutput And n_hist_prv = 0 Then
            vh_out(iterate) = h
        ElseIf storeOutput And n_hist_prv > 0 Then
            h_tmp = vh_out(iterate)
            ReDim Preserve h_tmp(1 To n_hidden, 1 To n_hist_prv + n_len)
            For k = 1 To n_len
                For i = 1 To n_hidden
                    h_tmp(i, n_hist_prv + k) = h(i, k)
                Next i
            Next k
            vh_out(iterate) = h_tmp
        End If
        
        'Assign output as input to next layer
        ReDim xin(1 To n_hidden, 1 To n_len)
        For k = 1 To n_len
            For i = 1 To n_hidden
                xin(i, k) = h(i, k)
            Next i
        Next k
        n_input = n_hidden

    Next iterate

    'Output layer
    ReDim y(1 To pn_output, 1 To n_len)
    v = WorksheetFunction.MMult(pWOut, xin)
    If pOutputType = "SOFTMAX" Then
        ReDim h(1 To pn_output)
        For k = 1 To n_len
            For i = 1 To pn_output
                h(i) = v(i, k) + pbOut(i)
            Next i
            h = f_Activate(h, "SOFTMAX")
            For i = 1 To pn_output
                y(i, k) = h(i)
            Next i
        Next k
    Else
        If pn_output = 1 Then
            For k = 1 To n_len
                y(1, k) = f_Activate(v(k) + pbOut(1), pOutputType)
            Next k
        Else
            For k = 1 To n_len
                For i = 1 To pn_output
                    y(i, k) = f_Activate(v(i, k) + pbOut(i), pOutputType)
                Next i
            Next k
        End If
    End If
    
    FwdPass = y
    
End Function


'Calculate and accumulate gradients with respect to all weights
'gradients are stored in memory
Sub Backward(y_out, Optional grad_out, Optional y_tgt As Variant, Optional isdEdx As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_output As Long, n_hidden As Long, n_len As Long, n_input As Long
Dim iterate As Long
Dim xW() As Double, xb() As Double, dW() As Double, dBias() As Double
Dim y() As Double
Dim h() As Double
Dim v As Variant, x() As Double
Dim strType As String
Dim grad_curr As Variant, grad_nxt As Variant, grad_loc As Variant
Dim n_hidden_nxt As Long

    n_len = UBound(y_out, 2)
    n_output = UBound(y_out, 1)
    If n_output <> pn_output Then
        Debug.Print "cNeuralNet: Backward: output size incorrect."
        End
    End If
    
    'Placeholder to store changes in weights
    m = pn_hidden(pn_layer)
    ReDim dWOut(1 To pn_output, 1 To m)
    ReDim dbOut(1 To pn_output)
    ReDim vdWgts(1 To pn_layer)
    ReDim vdBias(1 To pn_layer)
    
    If isdEdx Then
        'dEdx is supplied instead of dEdy
        ReDim grad_curr(1 To n_output, 1 To n_len)
        For k = 1 To n_len
            For i = 1 To n_output
                grad_curr(i, k) = grad_out(i, k)
            Next i
        Next k
    Else
    
        ReDim grad_curr(1 To n_output, 1 To n_len)
        
        If pOutputType = "SOFTMAX" Then
        
            'For softmax directly calculate grad_nxt * grad_loc
            For k = 1 To n_len
                For i = 1 To n_output
                    grad_curr(i, k) = y_out(i, k) - y_tgt(i, k)
                Next i
            Next k
            
        Else
    
            grad_loc = f_LocalGrad(y_out, pOutputType)
            For k = 1 To n_len
                For i = 1 To n_output
                    grad_curr(i, k) = grad_out(i, k) * grad_loc(i, k)
                Next i
            Next k
    
        End If
    End If
    x = vh_out(pn_layer)
    For k = n_len To 1 Step -1
        For i = 1 To pn_output
            dbOut(i) = dbOut(i) + grad_curr(i, k)
            For j = 1 To m
                dWOut(i, j) = dWOut(i, j) + grad_curr(i, k) * x(j, k)
            Next j
        Next i
    Next k

    grad_nxt = grad_curr
    n_hidden_nxt = pn_output
    For iterate = pn_layer To 1 Step -1

        n_hidden = pn_hidden(iterate)
        strType = pLayerType(iterate)

        'local gradient
        grad_loc = f_LocalGrad(vh_out(iterate), strType)
 
        'Wgts connected to next layer
        If iterate = pn_layer Then
            xW = pWOut
        Else
            xW = vWgts(iterate + 1)
        End If

        'Input to current layer
        If iterate = 1 Then
            n_input = pn_input
            x = px_in
        Else
            n_input = pn_hidden(iterate - 1)
            x = vh_out(iterate - 1)
        End If

        ReDim dW(1 To n_hidden, 1 To n_input)
        ReDim dBias(1 To n_hidden)

        'Wgt sum of next layer's gradient
        ReDim grad_curr(1 To n_hidden, 1 To n_len)
        
'        v = Application.WorksheetFunction.MMult(Application.WorksheetFunction.Transpose(xW), grad_nxt)
'        For k = n_len To 1 Step -1
'            For i = 1 To n_hidden
'                tmp_x = v(i, k) * grad_loc(i, k)
'                grad_curr(i, k) = tmp_x
'                dBias(i) = dBias(i) + tmp_x
'                For j = 1 To n_input
'                    dW(i, j) = dW(i, j) + tmp_x * x(j, k)
'                Next j
'            Next i
'        Next k
        
        For k = n_len To 1 Step -1
            For i = 1 To n_hidden
                For j = 1 To n_hidden_nxt
                    grad_curr(i, k) = grad_curr(i, k) + grad_nxt(j, k) * xW(j, i) * grad_loc(i, k)
                Next j
            Next i
        Next k
        
        For k = n_len To 1 Step -1
            For i = 1 To n_hidden
                dBias(i) = dBias(i) + grad_curr(i, k)
                For j = 1 To n_input
                    dW(i, j) = dW(i, j) + grad_curr(i, k) * x(j, k)
                Next j
            Next i
        Next k

        vdWgts(iterate) = dW
        vdBias(iterate) = dBias
        
        grad_nxt = grad_curr
        n_hidden_nxt = n_hidden
    Next iterate
    
End Sub


'Applied stored gradients to update weights
'stored gradients are erased once used
Sub UpdateWgt(learn_rate As Double, Optional useSpeedUp As String = "")
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_len As Long
Dim iterate As Long, n_hidden As Long
Dim xW() As Double, xb() As Double
Dim dW() As Double, db() As Double
Dim dW1() As Double, db1() As Double
Dim dW2() As Double, db2() As Double
Dim adam_discount1 As Double, adam_discount2 As Double

    If UCase(useSpeedUp) = "RMS" Then
        
        Call RMSProp_CalcRMS
        
        m = pn_hidden(pn_layer)
        For i = 1 To pn_output
            pbOut(i) = pbOut(i) - learn_rate * dbOut(i) / (Sqr(dbOut2(i)) + 0.00000001)
        Next i

        For j = 1 To m
            For i = 1 To pn_output
                pWOut(i, j) = pWOut(i, j) - learn_rate * dWOut(i, j) / (Sqr(dWOut2(i, j)) + 0.00000001)
            Next i
        Next j

        For iterate = 1 To pn_layer
    
            xW = vWgts(iterate)
            xb = vBias(iterate)
            dW = vdWgts(iterate)
            db = vdBias(iterate)
            dW2 = vdW2(iterate)
            db2 = vdB2(iterate)
            
            n = UBound(xW, 1)
            m = UBound(xW, 2)
            For i = 1 To n
                xb(i) = xb(i) - learn_rate * db(i) / (Sqr(db2(i)) + 0.00000001)
            Next i

            For j = 1 To m
                For i = 1 To n
                    xW(i, j) = xW(i, j) - learn_rate * dW(i, j) / (Sqr(dW2(i, j)) + 0.00000001)
                Next i
            Next j

            vWgts(iterate) = xW
            vBias(iterate) = xb

        Next iterate
        
    ElseIf UCase(useSpeedUp) = "ADAM" Then
        
        Call ADAM_CalcMoment
        
        adam_discount1 = 1 / (1 - 0.9 ^ pADAM_count)
        adam_discount2 = 1 / (1 - 0.999 ^ pADAM_count)
        
        m = pn_hidden(pn_layer)
        For i = 1 To pn_output
            pbOut(i) = pbOut(i) - learn_rate * dbOut1(i) * adam_discount1 / (Sqr(dbOut2(i) * adam_discount2) + 0.00000001)
        Next i
        For j = 1 To m
            For i = 1 To pn_output
                pWOut(i, j) = pWOut(i, j) - learn_rate * dWOut1(i, j) * adam_discount1 / (Sqr(dWOut2(i, j) * adam_discount2) + 0.00000001)
            Next i
        Next j
        
        For iterate = 1 To pn_layer
    
            xW = vWgts(iterate)
            xb = vBias(iterate)
'            dW = vdWgts(iterate)
'            db = vdBias(iterate)
            dW1 = vdW1(iterate)
            db1 = vdB1(iterate)
            dW2 = vdW2(iterate)
            db2 = vdB2(iterate)
            
            n = UBound(xW, 1)
            m = UBound(xW, 2)
            
            For i = 1 To n
                xb(i) = xb(i) - learn_rate * db1(i) * adam_discount1 / (Sqr(db2(i) * adam_discount2) + 0.00000001)
            Next i
            For j = 1 To m
                For i = 1 To n
                    xW(i, j) = xW(i, j) - learn_rate * dW1(i, j) * adam_discount1 / (Sqr(dW2(i, j) * adam_discount2) + 0.00000001)
                Next i
            Next j
            
            vWgts(iterate) = xW
            vBias(iterate) = xb
            
        Next iterate
    
    
    Else
    
        m = pn_hidden(pn_layer)
        For i = 1 To pn_output
            pbOut(i) = pbOut(i) - learn_rate * dbOut(i)
        Next i
        For j = 1 To m
            For i = 1 To pn_output
                pWOut(i, j) = pWOut(i, j) - learn_rate * dWOut(i, j)
            Next i
        Next j
        
        For iterate = 1 To pn_layer
    
            xW = vWgts(iterate)
            xb = vBias(iterate)
            dW = vdWgts(iterate)
            db = vdBias(iterate)
     
            n = UBound(xW, 1)
            m = UBound(xW, 2)
            
            For i = 1 To n
                xb(i) = xb(i) - learn_rate * db(i)
            Next i
            For j = 1 To m
                For i = 1 To n
                    xW(i, j) = xW(i, j) - learn_rate * dW(i, j)
                Next i
            Next j
            vWgts(iterate) = xW
            vBias(iterate) = xb
            
        Next iterate
    
    End If
    Erase vdWgts, vdBias, dWOut, dbOut
    
End Sub


'Fit with cross validation, arguments are similar to Fit, additional arguments include
'outputs a scalar which is the average k-fold validation error
'n_fold     number of validation sets
'strLossCV  loss function used in measuring cross-validation performance. Default is the same as strLossType
Function FitCV(x, y_tgt, Optional n_fold As Long = 10, Optional n_epoch As Long = 5, Optional batch_size As Long = -1, _
            Optional learn_rate As Double = 0.0001, _
            Optional strLossType As String = "SSE", Optional strLossCV As String = "", _
            Optional err_tol As Double = 0.000000001, _
            Optional err_tol_rel As Double = 0.0001, _
            Optional useSpeedUp As String = "", _
            Optional learnSchedule As String = "", _
            Optional statusShown As Long = 10)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, ii As Long, jj As Long
Dim n_len As Long, n_input As Long, n_output As Long
Dim i_fold As Long, fold_size As Long
Dim x_idx() As Long, x_idx_fit() As Long, x_idx_test() As Long
Dim x_err As Double, x_err_min As Double, x_err_kfold As Double
Dim y() As Double
Dim x_sub As Variant, y_tgt_sub As Variant
Dim cNNtest As cNeuralNet
Dim cNNOut As cNeuralNet
Dim strLossCV2 As String
    
    If strLossCV = "" Then
        strLossCV2 = strLossType
    Else
        strLossCV2 = strLossCV
    End If

    n_len = UBound(x, 2)
    n_input = UBound(x, 1)
    n_output = UBound(y_tgt, 1)
    fold_size = Int(n_len / n_fold)
    
    'Shuffle raw data
    ReDim x_idx(1 To n_len)
    For i = 1 To n_len
        x_idx(i) = i
    Next i
    Call Shuffle(x_idx)
    
    x_err_kfold = 0
    x_err_min = Exp(70)
    For i_fold = 1 To n_fold
        
        DoEvents
        Application.StatusBar = "cNeuralNet: FitCV: Fold: " & i_fold & "/" & n_fold & "..."
        
        Set cNNtest = New cNeuralNet
        Call cNNtest.Clone(Me)
        
        'divide data into training and test set
        ReDim x_idx_test(1 To fold_size)
        ReDim x_idx_fit(1 To n_len - fold_size)
        m = 0: n = 0
        For i = 1 To n_len
            If i > ((i_fold - 1) * fold_size) And i <= (i_fold * fold_size) Then
                n = n + 1
                x_idx_test(n) = x_idx(i)
            Else
                m = m + 1
                x_idx_fit(m) = x_idx(i)
            End If
        Next i
        ReDim Preserve x_idx_test(1 To n)
        ReDim Preserve x_idx_fit(1 To m)

        'Fit model on training set
        x_sub = SubsetIdx(x, x_idx:=x_idx_fit)
        y_tgt_sub = SubsetIdx(y_tgt, x_idx:=x_idx_fit)
        Call cNNtest.Fit(x_sub, y_tgt_sub, n_epoch:=n_epoch, batch_size:=batch_size, _
                        learn_rate:=learn_rate, strLossType:=strLossType, _
                        err_tol:=err_tol, err_tol_rel:=err_tol_rel, _
                        learnSchedule:=learnSchedule, _
                        useSpeedUp:=useSpeedUp, _
                        statusShown:=statusShown)
                        
        'evaluate performance on testing set
        x_sub = SubsetIdx(x, x_idx:=x_idx_test)
        y_tgt_sub = SubsetIdx(y_tgt, x_idx:=x_idx_test)
        y = cNNtest.FwdPass(x_sub, False)
        x_err_kfold = x_err_kfold + f_calcLoss(y, y_tgt_sub, strLossCV2)

        'evaluate performance on full set
        y = cNNtest.FwdPass(x, False)
        x_err = f_calcLoss(y, y_tgt, strLossCV2)
        
        If x_err < x_err_min Then
            x_err_min = x_err
            Set cNNOut = New cNeuralNet
            Call cNNOut.Clone(cNNtest)
        End If
        
    Next i_fold
    
    FitCV = x_err_kfold / n_fold
    
    Call Me.Clone(cNNOut)
    Set cNNOut = Nothing
    Set cNNtest = Nothing
End Function


'Fit input x to response y_tgt
'Upon finish, training progress can be returned using property TrainProgress,
'which shows the loss function after each epoch
'n_epoch        number of epochs, 1 epoch means going over the whole data set once
'batch_size     size of mini-batch to use in SGD. Default is -1 where all data is used in a single epoch
'strLossType    can be "SSE" or "CROSSENTROPY"
'learn_rate     learning rate
'err_tol        criteria for convergence if loss is less than err_tol
'err_tol_rel    criteria for convergence if change in loss function is less than err_rol_rel * loss from previous epoch
'useSpeedUp     can be either "", "RMS" or "ADAM"
'learnSchedule  can be "", "DECAY", or "AGGRESSIVE"
'statusShown    show progress in status bar every n epochs
Sub Fit(x, y_tgt, Optional n_epoch As Long = 5, Optional batch_size As Long = -1, _
            Optional learn_rate As Double = 0.0001, _
            Optional strLossType As String = "SSE", _
            Optional err_tol As Double = 0.000000001, _
            Optional err_tol_rel As Double = 0.0001, _
            Optional useSpeedUp As String = "", _
            Optional learnSchedule As String = "AGGRESSIVE", _
            Optional statusShown As Long = 10)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_len As Long
Dim ii As Long, jj As Long
Dim i_epoch As Long, i_batch As Long
Dim y() As Double
Dim grad_out() As Double
Dim x_err As Double, x_err_prv As Double
Dim step_size As Double
Dim batch_count As Long, converge_count As Long

    step_size = learn_rate
    strLossType = UCase(Trim(strLossType))
    If pn_output = 1 And strLossType = "CROSSENTROPY" And pOutputType <> "SIGMOID" Then
        Debug.Print "cNeuralNet: Train: CROSSENTROPY must be used with SIGMOID output layer for binary classification."
        End
    End If
    
    batch_count = 0
    converge_count = 0
    n_len = UBound(x, 2)
    x_err_prv = Exp(70)
    ReDim pTrainProgress(1 To 3, 1 To 1)

    For i_epoch = 1 To n_epoch
        
        DoEvents
        If (i_epoch - 1) Mod statusShown = 0 Then
            Application.StatusBar = "cNeuralNet: Train: Epoch " & i_epoch & "/" & n_epoch & "..."
        End If
        
        Call CacheCurrentWgt
        
        If batch_size < 0 Or batch_size >= n_len Then
            y = FwdPass(x, storeOutput:=True)
            If pOutputType = "SOFTMAX" Then
                Call Backward(y, , y_tgt)
            Else
                grad_out = f_calcdEdy(y, y_tgt, strLossType)
                Call Backward(y, grad_out)
            End If
            
            batch_count = batch_count + 1

            Call UpdateWgt(step_size, UCase(useSpeedUp))
            Call ClearHist
            
        Else
            ii = 0
            jj = 0
            Do While (jj + 1) <= n_len
                
                'Extract a mini-batch
                ii = jj + 1
                jj = ii + batch_size - 1
                If jj > n_len Then jj = n_len
                ReDim xx(1 To pn_input, 1 To (jj - ii + 1))
                ReDim yy(1 To pn_output, 1 To (jj - ii + 1))
                For m = ii To jj
                    k = m - ii + 1
                    For i = 1 To pn_input
                        xx(i, k) = x(i, m)
                    Next i
                    For i = 1 To pn_output
                        yy(i, k) = y_tgt(i, m)
                    Next i
                Next m
                batch_count = batch_count + 1

                y = FwdPass(xx, storeOutput:=True)

                If pOutputType = "SOFTMAX" Then
                    Call Backward(y, , yy)
                Else
                    grad_out = f_calcdEdy(y, yy, strLossType)
                    Call Backward(y, grad_out)
                End If

                Call UpdateWgt(step_size, UCase(useSpeedUp))
                Call ClearHist
            Loop
        End If
        
        'Evalulate this update
        y = FwdPass(x, storeOutput:=False)
        x_err = f_calcLoss(y, y_tgt, strLossType)
        
        ReDim Preserve pTrainProgress(1 To 3, 1 To i_epoch)
        pTrainProgress(1, i_epoch) = i_epoch
        pTrainProgress(2, i_epoch) = x_err
        pTrainProgress(3, i_epoch) = step_size
        
        'Check for convergence
        If (x_err <= x_err_prv) And ((x_err_prv - x_err) <= (err_tol_rel * Abs(x_err_prv)) Or x_err <= err_tol) Then
            converge_count = converge_count + 1
        Else
            converge_count = 0
        End If
        If converge_count >= 5 Then Exit For
        
        'Increase step size if this update is good
        If UCase(learnSchedule) = "AGGRESSIVE" Then
            If x_err <= x_err_prv Then
                step_size = step_size * 1.05
                x_err_prv = x_err
            Else
                step_size = step_size * 0.1
                Call RestoreWgt
                If UCase(useSpeedUp) = "ADAM" Then Call ADAM_Init
                If step_size < 0.00000000001 Then Exit For
            End If
        ElseIf UCase(learnSchedule) = "DECAY" Then
            step_size = step_size * (1 - (i_epoch - 1) * 1# / (n_epoch - 1))
            x_err_prv = x_err
        ElseIf UCase(learnSchedule) = "" Then
            x_err_prv = x_err
        Else
            Debug.Print "cNeuralNet: Fit: learnSchedule must be """", ""DECAY"" or ""AGGRESSIVE"""
        End If
        Call ClearCacheWgt
        Call ResetWgtChg

    Next i_epoch
    
    If UCase(useSpeedUp) = "RMS" Then Call RMSProp_Clear
    If UCase(useSpeedUp) = "ADAM" Then Call ADAM_Clear
    
    Application.StatusBar = False
End Sub


'Derivative of loss function w.r.t. output
'output is a double array of dimension (1:pn_output, 1:n)
Function f_calcdEdy(y, y_tgt, strLossType)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_len As Long
Dim dEdy() As Double
    
    n_len = UBound(y, 2)
    ReDim dEdy(1 To pn_output, 1 To n_len)
    
    Select Case UCase(strLossType)
    Case "SSE"
        For m = 1 To n_len
            For i = 1 To pn_output
                dEdy(i, m) = y(i, m) - y_tgt(i, m)
            Next i
        Next m
    
    Case "CROSSENTROPY"
    
        If pn_output = 1 Then
            For m = 1 To n_len
                dEdy(1, m) = (y(1, m) - y_tgt(1, m)) / (y(1, m) * (1 - y(1, m)))
            Next m
        End If
        
    End Select
    
    f_calcdEdy = dEdy
    
End Function


'Calculate loss function
'output is a single scalar
'strLossType can be "SSE", "CROSSENTROPY" or "MISCLASS"
Function f_calcLoss(y, y_tgt, strLossType)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_len As Long
Dim x_error As Double, tmp_x As Double
Dim i_max As Long
Dim x_class() As Long

    n_len = UBound(y, 2)
    
    Select Case UCase(strLossType)
    Case "SSE"
    
        x_error = 0
        For m = 1 To n_len
            For i = 1 To pn_output
                x_error = x_error + (y(i, m) - y_tgt(i, m)) ^ 2
            Next i
        Next m
        x_error = 0.5 * x_error / n_len
        
    Case "CROSSENTROPY"
        
        If pn_output = 1 Then
            
            x_error = 0
            For m = 1 To n_len
                x_error = x_error - y_tgt(i, m) * Log(y(i, m)) - (1 - y_tgt(i, m)) * Log(1 - y(i, m))
            Next m
            x_error = x_error / n_len

        Else
            x_error = 0
            For m = 1 To n_len
                For i = 1 To pn_output
                    x_error = x_error - y_tgt(i, m) * Log(y(i, m))
                Next i
            Next m
            x_error = x_error / n_len
        End If
        
    Case "MISCLASS"
        
        x_error = 0
        x_class = SampleSoftmax(y, False)
        For m = 1 To n_len
            If y_tgt(x_class(m), m) = 1 Then x_error = x_error + 1
        Next m
        x_error = 1 - x_error / n_len
        
    End Select
    
    f_calcLoss = x_error
    
End Function


'Activation function
'for softmax, accepts a vector input and outputs a vector
'otherwise,  accepts a scalar input and outputd a scalar
Private Function f_Activate(x As Variant, strType As String)
Dim i As Long, n As Long
Dim y() As Double
Dim tmp_x As Double
    Select Case strType
    Case "RELU"
        If x <= 0 Then
            f_Activate = 0
        Else
            f_Activate = x
        End If
    Case "LEAKYRELU"
        If x <= 0 Then
            f_Activate = 0.01 * x
        Else
            f_Activate = x
        End If
    Case "SIGMOID"
    
        If x > 25 Then
            f_Activate = 1
        ElseIf x < -25 Then
            f_Activate = 0
        Else
            f_Activate = 1 / (1 + Exp(-x))
        End If
        
    Case "LINEAR"
    
        f_Activate = x
    
    Case "TANH"
    
        If x > 25 Then
            f_Activate = 1
        ElseIf x < -25 Then
            f_Activate = -1
        Else
            f_Activate = (Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x))
        End If
        
    Case "SOFTMAX"
    
        tmp_x = 0
        n = UBound(x, 1)
        ReDim y(1 To n)
        For i = 1 To n
            y(i) = Exp(x(i))
            tmp_x = tmp_x + y(i)
        Next i
        For i = 1 To n
            y(i) = y(i) / tmp_x
        Next i
        f_Activate = y
        
    Case Else
        
        Debug.Print "cNeuralNet: f_Activate: " & strType & " is not supported"
        End
        
    End Select
End Function


'Derivative of activation function w.r.t. its input
'returns a double array of size (1:n_output, 1: n)
Private Function f_LocalGrad(y As Variant, strType As String)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_len As Long
Dim grad() As Double
Dim tmp_x As Double
    
    n = UBound(y, 1)
    n_len = UBound(y, 2)
    ReDim grad(1 To n, 1 To n_len)
    
    Select Case strType
    Case "RELU"
    
        For k = 1 To n_len
            For i = 1 To n
                If y(i, k) > 0 Then grad(i, k) = 1
            Next i
        Next k
    Case "LEAKYRELU"
    
        For k = 1 To n_len
            For i = 1 To n
                If y(i, k) > 0 Then
                    grad(i, k) = 1
                ElseIf y(i, k) < 0 Then
                    grad(i, k) = 0.01
                End If
            Next i
        Next k
    Case "SIGMOID"
        
        For k = 1 To n_len
            For i = 1 To n
                grad(i, k) = y(i, k) * (1 - y(i, k))
            Next i
        Next k
        
    Case "LINEAR"
    
        For k = 1 To n_len
            For i = 1 To n
                grad(i, k) = 1
            Next i
        Next k
        
    Case "TANH"
    
        For k = 1 To n_len
            For i = 1 To n
                grad(i, k) = 1 - y(i, k) ^ 2
            Next i
        Next k
        
    Case Else
        
        Debug.Print "cNeuralNet: f_LocalGrad: does not support " & strType
        End
        
    End Select
    
    f_LocalGrad = grad
End Function


'Output network parameters onto an Excel sheet so it can be replicated later
Sub PrintNetwork(wksht As Worksheet, Optional startRow As Long = 1)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim xW() As Double, xb() As Double
Dim i_row As Long
With wksht
    .Cells(startRow, 1).Value = pOutputType
    .Cells(startRow + 1, 1).Value = pn_input
    .Cells(startRow + 2, 1).Value = pn_output
    .Cells(startRow + 3, 1).Value = pn_layer
    .Cells(startRow + 4, 1).Resize(1, pn_layer).Value = pLayerType
    .Cells(startRow + 5, 1).Resize(1, pn_layer).Value = pn_hidden
    i_row = startRow + 5
    For iterate = 1 To pn_layer
        xW = vWgts(iterate)
        xb = vBias(iterate)
        m = UBound(xW, 1)
        n = UBound(xW, 2)
        .Cells(i_row + 1, 1).Resize(m, n).Value = xW
        .Cells(i_row + 1 + m, 1).Resize(1, m).Value = xb
        i_row = i_row + 1 + m
    Next iterate
    
    n = UBound(pWOut, 2)
    .Cells(i_row + 1, 1).Resize(pn_output, n).Value = pWOut
    .Cells(i_row + 1 + pn_output, 1).Resize(1, pn_output).Value = pbOut
    
    pn_row = i_row + 1 + pn_output - startRow + 1
    
End With
End Sub

'Read in saved network
Sub ReadNetwork(wksht As Worksheet, Optional startRow As Long = 1)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim n_input As Long, n_hidden As Long
Dim xW() As Double, xb() As Double
Dim i_row As Long
With wksht

    pn_hist = 0
    pOutputType = .Cells(startRow, 1).Value
    pn_input = .Cells(startRow + 1, 1).Value
    pn_output = .Cells(startRow + 2, 1).Value
    pn_layer = .Cells(startRow + 3, 1).Value
    
    ReDim vWgts(1 To pn_layer)
    ReDim vBias(1 To pn_layer)
    ReDim pn_hidden(1 To pn_layer)
    ReDim pLayerType(1 To pn_layer)
    ReDim pWOut(1 To pn_output, 1 To pn_input)
    ReDim pbOut(1 To pn_output)
    
    For j = 1 To pn_layer
        pLayerType(j) = .Cells(startRow + 4, j).Value
        pn_hidden(j) = .Cells(startRow + 5, j).Value
    Next j

    i_row = startRow + 5
    n_input = pn_input
    For iterate = 1 To pn_layer
        n_hidden = pn_hidden(iterate)
        ReDim xW(1 To n_hidden, 1 To n_input)
        ReDim xb(1 To n_hidden)
        For i = 1 To n_input
            For j = 1 To n_hidden
                xW(j, i) = .Cells(i_row + j, i).Value
            Next j
        Next i
        For j = 1 To n_hidden
            xb(j) = .Cells(i_row + n_hidden + 1, j).Value
        Next j
        
        vWgts(iterate) = xW
        vBias(iterate) = xb
        
        i_row = i_row + 1 + n_hidden
        n_input = n_hidden
    Next iterate

    ReDim pWOut(1 To pn_output, 1 To n_input)
    ReDim pbOut(1 To pn_output)
    For i = 1 To n_input
        For j = 1 To pn_output
            pWOut(j, i) = .Cells(i_row + j, i).Value
        Next j
    Next i
    For j = 1 To pn_output
        pbOut(j) = .Cells(i_row + pn_output + 1, j).Value
    Next j
    
    pn_row = i_row + 1 + pn_output - startRow + 1
    
End With
End Sub

'Convert a vector of labels to dummy variables array
'also return a dictonary of unique_labels
Function Label2Dummy(x_label As Variant, unique_labels As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim n_output As Long, isMatch As Boolean
Dim x_dummy() As Long, iArr() As Long
    
    n = UBound(x_label, 1)
    ReDim unique_labels(1 To n)
    ReDim iArr(1 To n)
    
    n_output = 1
    unique_labels(1) = x_label(1)
    iArr(1) = 1
    
    For i = 2 To n
        isMatch = False
        For j = 1 To n_output
            If unique_labels(j) = x_label(i) Then
                iArr(i) = j
                isMatch = True
                Exit For
            End If
        Next j
        If Not isMatch Then
            n_output = n_output + 1
            iArr(i) = n_output
            unique_labels(n_output) = x_label(i)
        End If
    Next i
    ReDim Preserve unique_labels(1 To n_output)
    
    
    ReDim x_dummy(1 To n_output, 1 To n)
    For i = 1 To n
        x_dummy(iArr(i), i) = 1
    Next i
    
    Label2Dummy = x_dummy
End Function


'convert dummy variables to a vector of lables, according to the supplied dictionary
Function Dummy2Label(x_dummy As Variant, unique_labels As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_output As Long
Dim x_label As Variant
Dim tmp_x As Double

    n_output = UBound(x_dummy, 1)
    n = UBound(x_dummy, 2)
    ReDim x_label(1 To n)
    
    For i = 1 To n
        tmp_x = x_dummy(1, i)
        x_label(i) = unique_labels(1)
        For j = 2 To n_output
            If x_dummy(j, i) > tmp_x Then
                tmp_x = x_dummy(j, i)
                x_label(i) = unique_labels(j)
            End If
        Next j
    Next i

    Dummy2Label = x_label
End Function



'Sample outputs from discrete distribution
Function SampleSoftmax(y As Variant, Optional isStochastic As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim n_T As Long, n_output As Long
Dim tmp_max As Double, cumprob As Double
Dim x_class() As Long

    n_T = UBound(y, 2)
    n_output = UBound(y, 1)
    ReDim x_class(1 To n_T)
    
    If isStochastic Then
        
        If n_output = 1 Then
            For iterate = 1 To n_T
                If y(1, iterate) > Rnd() Then
                    x_class(iterate) = 1
                End If
            Next iterate
        Else
        
            For iterate = 1 To n_T
                tmp_max = Rnd()
                cumprob = y(1, iterate)
                If tmp_max < cumprob Then
                    x_class(iterate) = 1
                Else
                    For i = 2 To n_output
                        cumprob = cumprob + y(i, iterate)
                        If tmp_max < cumprob Then
                            x_class(iterate) = i
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
                    x_class(iterate) = 1
                End If
            Next iterate
        Else
            For iterate = 1 To n_T
                tmp_max = y(1, iterate)
                x_class(iterate) = 1
                For i = 2 To n_output
                    If y(i, iterate) > tmp_max Then
                        tmp_max = y(i, iterate)
                        x_class(iterate) = i
                    End If
                Next i
            Next iterate
        End If
        
    End If
    
    SampleSoftmax = x_class

End Function


'Return a subset of input array x()
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

'Shuffle an integer vector x()
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



'=====================================================================================================================================
'Experiment with BFGS since it seems to perform batter for small problems like linear regresion and low dimension curve fitting
'line search algorithm was copied from"Numerical Optimization" by Jorge Ncedal and Stephen Wright
'=====================================================================================================================================


'Convert network wgts to 1D vector for easier manipulation
'isReverse, convert 1D vector back to network wgts when set to True
Private Sub flatten_wgts(wgts_vec() As Double, Optional isGrad As Boolean = False, Optional isReverse As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_input As Long, n_output As Long
Dim iterate As Long
Dim xW() As Double, xb() As Double
    
    If isReverse Then
        
        n = 0
        For iterate = 1 To pn_layer
            If isGrad Then
                xW = vdWgts(iterate)
                xb = vdBias(iterate)
            Else
                xW = vWgts(iterate)
                xb = vBias(iterate)
            End If
            n_output = UBound(xW, 1)
            n_input = UBound(xW, 2)
            For i = 1 To n_input
                For j = 1 To n_output
                    n = n + 1
                    xW(j, i) = wgts_vec(n)
                Next j
            Next i
            For j = 1 To n_output
                n = n + 1
                xb(j) = wgts_vec(n)
            Next j
            If isGrad Then
                vdWgts(iterate) = xW
                vdBias(iterate) = xb
            Else
                vWgts(iterate) = xW
                vBias(iterate) = xb
            End If
        Next iterate
        
        n_input = UBound(pWOut, 2)
        If isGrad Then
            For i = 1 To n_input
                For j = 1 To pn_output
                    n = n + 1
                    dWOut(j, i) = wgts_vec(n)
                Next j
            Next i
            For j = 1 To pn_output
                n = n + 1
                dbOut(j) = wgts_vec(n)
            Next j
        Else
            For i = 1 To n_input
                For j = 1 To pn_output
                    n = n + 1
                    pWOut(j, i) = wgts_vec(n)
                Next j
            Next i
            For j = 1 To pn_output
                n = n + 1
                pbOut(j) = wgts_vec(n)
            Next j
        End If
        
        Exit Sub
    End If

    
    ReDim wgts_vec(1 To pn_unit)
    n = 0
    For iterate = 1 To pn_layer
        If isGrad Then
            xW = vdWgts(iterate)
            xb = vdBias(iterate)
        Else
            xW = vWgts(iterate)
            xb = vBias(iterate)
        End If
        n_output = UBound(xW, 1)
        n_input = UBound(xW, 2)
        For i = 1 To n_input
            For j = 1 To n_output
                n = n + 1
                wgts_vec(n) = xW(j, i)
            Next j
        Next i
        For j = 1 To n_output
            n = n + 1
            wgts_vec(n) = xb(j)
        Next j
    Next iterate
    
    n_input = UBound(pWOut, 2)
    If isGrad Then
        For i = 1 To n_input
            For j = 1 To pn_output
                n = n + 1
                wgts_vec(n) = dWOut(j, i)
            Next j
        Next i
        For j = 1 To pn_output
            n = n + 1
            wgts_vec(n) = dbOut(j)
        Next j
    Else
        For i = 1 To n_input
            For j = 1 To pn_output
                n = n + 1
                wgts_vec(n) = pWOut(j, i)
            Next j
        Next i
        For j = 1 To pn_output
            n = n + 1
            wgts_vec(n) = pbOut(j)
        Next j
    End If

End Sub

'Finite difference to approximate Hessian
Private Function CalcHessian(x, y_tgt, strLossType)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim ii As Long, jj As Long
Dim y() As Double, grad_out() As Double
Dim wgt_vec() As Double, grad_vec() As Double
Dim wgt_vec_new() As Double, grad_vec_new() As Double
Dim xHessian() As Double
Dim tmp_x As Double
Dim v As Variant

    ReDim xHessian(1 To pn_unit, 1 To pn_unit)
    
    'Evalulate current gradients
    y = FwdPass(x, storeOutput:=True)
    If pOutputType = "SOFTMAX" Then
        Call Backward(y, , y_tgt)
    Else
        grad_out = f_calcdEdy(y, y_tgt, strLossType)
        Call Backward(y, grad_out)
    End If
    
    'Save current wgts and gradients as vectors
    Call flatten_wgts(wgt_vec, False, False)
    Call flatten_wgts(grad_vec, True, False)
    Call ResetWgtChg
    Call ClearHist

    For i = 1 To pn_unit
        
        'change single wgt
        wgt_vec_new = wgt_vec
        wgt_vec_new(i) = wgt_vec_new(i) + 0.0001
        Call flatten_wgts(wgt_vec_new, False, True)
        
        'find new gradient after wgt change
        y = FwdPass(x, storeOutput:=True)
        If pOutputType = "SOFTMAX" Then
            Call Backward(y, , y_tgt)
        Else
            grad_out = f_calcdEdy(y, y_tgt, strLossType)
            Call Backward(y, grad_out)
        End If
        Call flatten_wgts(grad_vec_new, True, False)
        Call ResetWgtChg
        Call ClearHist
        
        'restore old wgt
        Call flatten_wgts(wgt_vec, False, True)
        
        'input changes in gradients into Hessian
        For j = i To pn_unit
            xHessian(i, j) = (grad_vec_new(j) - grad_vec(j)) / 0.0001
            xHessian(j, i) = xHessian(i, j)
        Next j
        
    Next i
    
    CalcHessian = xHessian
    
End Function


'Guess initial Hessian according to  Eq(6.20)
Private Function HessianGuess(x, y_tgt, strLossType)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim y() As Double, grad_out() As Double
Dim wgt_vec() As Double, grad_vec() As Double
Dim wgt_vec_new() As Double, grad_vec_new() As Double
Dim xHessian() As Double
Dim tmp_x As Double, tmp_y As Double
Dim s_vec() As Double, y_vec() As Double

    'Evalulate current gradients
    y = FwdPass(x, storeOutput:=True)
    If pOutputType = "SOFTMAX" Then
        Call Backward(y, , y_tgt)
    Else
        grad_out = f_calcdEdy(y, y_tgt, strLossType)
        Call Backward(y, grad_out)
    End If
    
    'Save current wgts and gradients as vectors
    Call flatten_wgts(wgt_vec, False, False)
    Call flatten_wgts(grad_vec, True, False)
    Call ResetWgtChg
    Call ClearHist
    
    'try updating wgts
    ReDim s_vec(1 To pn_unit)
    ReDim wgt_vec_new(1 To pn_unit)
    For i = 1 To pn_unit
        s_vec(i) = -0.001 * grad_vec(i)
        wgt_vec_new(i) = wgt_vec(i) + s_vec(i)
    Next i
    
    'import new wgts into network and evaluate new gradients
    Call flatten_wgts(wgt_vec_new, False, True)
    y = FwdPass(x, storeOutput:=True)
    If pOutputType = "SOFTMAX" Then
        Call Backward(y, , y_tgt)
    Else
        grad_out = f_calcdEdy(y, y_tgt, strLossType)
        Call Backward(y, grad_out)
    End If
    Call flatten_wgts(grad_vec_new, True, False)
    Call ResetWgtChg
    Call ClearHist
    
    'Evaluate y
    ReDim y_vec(1 To pn_unit)
    For i = 1 To pn_unit
        y_vec(i) = grad_vec_new(i) - grad_vec(i)
    Next i
    
    'Guess initial Hessian (y^T s) / (y^T y)
    tmp_x = 0: tmp_y = 0
    ReDim xHessian(1 To pn_unit, 1 To pn_unit)
    For i = 1 To pn_unit
        tmp_x = tmp_x + s_vec(i) * y_vec(i)
        tmp_y = tmp_y + y_vec(i) * y_vec(i)
    Next i
    For i = 1 To pn_unit
        xHessian(i, i) = tmp_x / tmp_y
    Next i
    
    'retore old wgts in network
    Call flatten_wgts(wgt_vec, False, True)
    
    HessianGuess = xHessian
    
End Function


'Fit with BFGS, notation used here are from Wikipedia for easy reference
'https://en.wikipedia.org/wiki/Broyden%E2%80%93Fletcher%E2%80%93Goldfarb%E2%80%93Shanno_algorithm)
Sub FitBFGS(x, y_tgt, Optional n_epoch As Long = 5, _
            Optional strLossType As String = "SSE", _
            Optional err_tol As Double = 0.000000001, _
            Optional err_tol_rel As Double = 0.0001, _
            Optional statusShown As Long = 10)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim ii As Long, jj As Long, i_epoch As Long
Dim converge_count As Long, n_linesearchfail As Long
Dim y() As Double, grad_out() As Double
Dim x_err As Double, x_err_prv As Double, x_curve As Double, x_curve_prv As Double
Dim step_size As Double
Dim wgt_vec() As Double, grad_vec() As Double
Dim wgt_vec_new() As Double, grad_vec_new() As Double
Dim xHessian As Variant, xDir() As Double
Dim tmp_x As Double
Dim y_vec() As Double, s_vec() As Double, Hy_vec() As Double, sy As Double

    strLossType = UCase(Trim(strLossType))
    If pn_output = 1 And strLossType = "CROSSENTROPY" And pOutputType <> "SIGMOID" Then
        Debug.Print "cNeuralNet: Train: CROSSENTROPY must be used with SIGMOID output layer for binary classification."
        End
    End If
    
    converge_count = 0
    n_linesearchfail = 0
    ReDim pTrainProgress(1 To 3, 1 To 1)
    
    'Initialize Hessian
    xHessian = HessianGuess(x, y_tgt, strLossType)  'rough guess
'    xHessian = CalcHessian(x, y_tgt, strLossType)  'finite difference
'    ReDim xHessian(1 To pn_unit, 1 To pn_unit) 'Simple Identity
'    For i = 1 To pn_unit
'        xHessian(i, i) = 1
'    Next i
     
    'Evalulate current objective and gradients
    y = FwdPass(x, storeOutput:=True)
    x_err_prv = f_calcLoss(y, y_tgt, strLossType)
    If pOutputType = "SOFTMAX" Then
        Call Backward(y, , y_tgt)
    Else
        grad_out = f_calcdEdy(y, y_tgt, strLossType)
        Call Backward(y, grad_out)
    End If
    
    'Save current wgts and gradients as vectors
    Call flatten_wgts(wgt_vec, False, False)
    Call flatten_wgts(grad_vec, True, False)
    Call ResetWgtChg
    Call ClearHist
    
    For i_epoch = 1 To n_epoch

        DoEvents
        If (i_epoch - 1) Mod statusShown = 0 Then
            Application.StatusBar = "cNeuralNet: FitBFGS: Epoch " & i_epoch & "/" & n_epoch & "..."
        End If

        'Compute direction
        ReDim xDir(1 To pn_unit)
        For i = 1 To pn_unit
            tmp_x = 0
            For j = 1 To pn_unit
                tmp_x = tmp_x + xHessian(i, j) * grad_vec(j)
            Next j
            xDir(i) = -tmp_x
        Next i

        x_curve = 0
        For i = 1 To pn_unit
            x_curve = x_curve + xDir(i) * grad_vec(i)
        Next i
        
        'wolfe line seawrch algorithm, Algorithm3.5 from "Numerical Optimization, Nocedal"
        step_size = BFGS_LineSearch(x, y_tgt, strLossType, x_err_prv, x_curve, wgt_vec, grad_vec, xDir)
        
        'Apply step size
        Call BFGS_ApplyStepSize(step_size, x, y_tgt, strLossType, _
                                wgt_vec, grad_vec, xDir, _
                                x_err, wgt_vec_new, grad_vec_new)
        
        'Compute new Hessian
        sy = 0
        ReDim s_vec(1 To pn_unit)
        ReDim y_vec(1 To pn_unit)
        For i = 1 To pn_unit
            s_vec(i) = step_size * xDir(i)
            y_vec(i) = grad_vec_new(i) - grad_vec(i)
            sy = sy + s_vec(i) * y_vec(i)
        Next i
        
        ReDim Hy_vec(1 To pn_unit)
        For i = 1 To pn_unit
            tmp_x = 0
            For j = 1 To pn_unit
                tmp_x = tmp_x + xHessian(i, j) * y_vec(j)
            Next j
            Hy_vec(i) = tmp_x
        Next i
        
        tmp_x = 0
        For i = 1 To pn_unit
            tmp_x = tmp_x + y_vec(i) * Hy_vec(i)
        Next i
        tmp_x = (sy + tmp_x) / (sy * sy)

        For i = 1 To pn_unit
            For j = 1 To pn_unit
                xHessian(i, j) = xHessian(i, j) + tmp_x * s_vec(i) * s_vec(j) - (Hy_vec(i) * s_vec(j) + s_vec(i) * Hy_vec(j)) / sy
            Next j
        Next i

        'Evalulate this update
        ReDim Preserve pTrainProgress(1 To 3, 1 To i_epoch)
        pTrainProgress(1, i_epoch) = i_epoch
        pTrainProgress(2, i_epoch) = x_err
        pTrainProgress(3, i_epoch) = step_size

        'Check for convergence
        If (x_err <= x_err_prv) And ((x_err_prv - x_err) <= (err_tol_rel * Abs(x_err_prv)) Or x_err <= err_tol) Then
            converge_count = converge_count + 1
        Else
            converge_count = 0
        End If
        If converge_count >= 5 Then Exit For

        x_err_prv = x_err
        For i = 1 To pn_unit
            wgt_vec(i) = wgt_vec_new(i)
            grad_vec(i) = grad_vec_new(i)
        Next i

    Next i_epoch

    Application.StatusBar = False
End Sub





'=====================================
'Algorithm 3.5 (Line Search Algorithm)
'=====================================
'Note that during this algorithm, the network wgts keep getting changed, but wgt_vec() is not changed
'so we can always go back to use wgt_vec() to restore network wgts
Function BFGS_LineSearch(x, y_tgt, strLossType As String, x_err_start As Double, x_curve_start As Double, _
                wgt_vec() As Double, grad_vec() As Double, xDir() As Double) As Double
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim ii As Long, jj As Long, i_epoch As Long
Dim converge_count As Long, n_linesearchfail As Long
Dim y() As Double, grad_out() As Double
Dim x_err As Double, x_err_prv As Double
Dim step_size As Double, step_size_max As Double, step_size_best As Double, step_size_prv As Double
Dim wgt_vec_new() As Double, grad_vec_new() As Double
Dim x_err_new As Double
Dim tmp_x As Double, x_curve As Double, wolfe_tmp_y As Double

    x_err_prv = x_err_start
    step_size_max = 10
    step_size_prv = 0
    step_size = 1
    For iterate = 1 To 20
    
        'Evaluate current step size
        Call BFGS_ApplyStepSize(step_size, x, y_tgt, strLossType, _
                                wgt_vec, grad_vec, xDir, _
                                x_err, wgt_vec_new, grad_vec_new)
        
        x_curve = 0
        For i = 1 To pn_unit
            x_curve = x_curve + xDir(i) * grad_vec_new(i)
        Next i
        
        If (x_err > (x_err_start + 0.0001 * step_size * x_curve_start) Or (x_err >= x_err_prv And iterate > 1)) Then
            step_size = BFGS_LineSearch_Zoom(step_size_prv, step_size, _
                                x, y_tgt, strLossType, wgt_vec, grad_vec, xDir, _
                                x_err_prv, x_err_start, x_curve_start)
            Exit For
        End If
        
        If Abs(x_curve) <= (-0.9 * x_curve_start) Then
            Exit For
        End If
        
        If x_curve >= 0 Then
            step_size = BFGS_LineSearch_Zoom(step_size, step_size_prv, _
                                x, y_tgt, strLossType, wgt_vec, grad_vec, xDir, _
                                x_err, x_err_start, x_curve_start)
            Exit For
        End If
        
        step_size = step_size + (step_size_max - step_size) * 0.5
        x_err_prv = x_err
        
    Next iterate
    
    BFGS_LineSearch = step_size

End Function


'======================
'Algorithm 3.6 (zoom)
'======================
Private Function BFGS_LineSearch_Zoom(step_size_lo As Double, step_size_hi As Double, x, y_tgt, strLossType As String, _
                            wgt_vec() As Double, grad_vec() As Double, xDir() As Double, _
                            x_err_lo As Double, x_err_start As Double, x_curve_start As Double) As Double
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim ii As Long, jj As Long, i_epoch As Long
Dim converge_count As Long, n_linesearchfail As Long
Dim y() As Double, grad_out() As Double
Dim x_err As Double, x_err_prv As Double, x_err_search As Double
Dim step_size As Double, step_size_max As Double, step_size_best As Double, step_size_prv As Double
Dim wgt_vec_new() As Double, grad_vec_new() As Double
Dim x_err_new As Double
Dim tmp_x As Double, x_curve As Double, x_curve_new As Double
Dim y_vec() As Double, s_vec() As Double, Hy_vec() As Double, sy As Double
Dim isFound As Boolean

    For iterate = 1 To 100
    
        step_size = (step_size_lo + step_size_hi) * 0.5
        
        'Evaluate current step size
        Call BFGS_ApplyStepSize(step_size, x, y_tgt, strLossType, _
                                wgt_vec, grad_vec, xDir, _
                                x_err, wgt_vec_new, grad_vec_new)

        x_curve = 0
        For i = 1 To pn_unit
            x_curve = x_curve + xDir(i) * grad_vec_new(i)
        Next i
        
        If (x_err > (x_err_start + 0.0001 * step_size * x_curve_start) Or x_err >= x_err_lo) Then
            step_size_hi = step_size
        Else
            
            If Abs(x_curve) <= (-0.9 * x_curve_start) Then
                Exit For
            End If
            If x_curve * (step_size_hi - step_size_lo) >= 0 Then
                step_size_hi = step_size_lo
            End If
            step_size_lo = step_size
            x_err_lo = x_err
        End If
        
    Next iterate
    
    BFGS_LineSearch_Zoom = step_size
    
End Function


'Apply step size to update network
'The outputs are x_err, wgt_vec_new() and grad_vec_new()
Private Sub BFGS_ApplyStepSize(step_size, x, y_tgt, strLossType As String, _
                                wgt_vec() As Double, grad_vec() As Double, xDir() As Double, _
                                x_err As Double, wgt_vec_new() As Double, grad_vec_new() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim y() As Double, grad_out() As Double
Dim tmp_x As Double, x_curve As Double

    'Update wgts and insert them into network
    ReDim wgt_vec_new(1 To pn_unit)
    For i = 1 To pn_unit
        wgt_vec_new(i) = wgt_vec(i) + step_size * xDir(i)
    Next i
    Call flatten_wgts(wgt_vec_new, isGrad:=False, isReverse:=True)
    
    'Evalulate loss function and new gradients
    y = FwdPass(x, storeOutput:=True)
    x_err = f_calcLoss(y, y_tgt, strLossType)
    If pOutputType = "SOFTMAX" Then
        Call Backward(y, , y_tgt)
    Else
        grad_out = f_calcdEdy(y, y_tgt, strLossType)
        Call Backward(y, grad_out)
    End If
    Call flatten_wgts(grad_vec_new, True, isReverse:=False)
    Call ResetWgtChg
    Call ClearHist
    
End Sub
