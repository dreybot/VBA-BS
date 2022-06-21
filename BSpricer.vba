Function BSCalc(req, CP, spot, strike, tExp, div, sigma, hRate, cRate)

    DF = Exp(-hRate * tExp)
    cFactor = Exp(-(cRate - hRate) * tExp)
    
    fwd = spot * Exp((hRate - div) * tExp)
    
    d1 = 1 / sigma / (tExp ^ 0.5) * (Log(fwd / strike) + sigma * sigma / 2 * tExp)
    d2 = d1 - sigma * tExp ^ 0.5
    
    
    'call
    CallPrice = WorksheetFunction.Max(DF * (fwd * normDist(d1) - strike * normDist(d2)), 0) * cFactor
    CallDelta = normDist(d1) * cFactor
    CallGamma = normDist(d1, , , False) / (spot * sigma * tExp ^ 0.5) * spot / 100 * cFactor
    CallVega = spot * normDist(d1, , , False) * tExp ^ 0.5 / 100 * cFactor
    CallTheta = (-spot * normDist(d1, , , False) * sigma / (2 * tExp ^ 0.5) _
                - hRate * strike * DF * normDist(d2)) * 1 / 365 * cFactor
    CallRho = strike * tExp * DF * normDist(d2) / 100 * cFactor
    
    'put
    'PutPrice = WorksheetFunction.Max(CallPrice + (-Fwd + strike) * DF)
    PutPrice = WorksheetFunction.Max(DF * (strike * normDist(-d2) - fwd * normDist(-d1)), 0) * cFactor
    PutDelta = -normDist(-d1) * cFactor
    PutGamma = normDist(d1, , , False) / (spot * sigma * tExp ^ 0.5) * spot / 100 * cFactor
    PutVega = spot * normDist(d1, , , False) * tExp ^ 0.5 / 100 * cFactor
    PutTheta = (-spot * normDist(d1, , , False) * sigma / (2 * tExp ^ 0.5) _
               + hRate * strike * DF * normDist(-d2)) * 1 / 365 * cFactor
    PutRho = -strike * tExp * DF * normDist(-d2) / 100 * cFactor
    
    ret = 0
    
    If LCase(Left(CP, 1)) = "c" Then
        Select Case req
            Case 1: ret = CallPrice
            Case 2: ret = CallDelta
            Case 3: ret = CallGamma
            Case 4: ret = CallTheta
            Case 5: ret = CallVega
            Case 8: ret = CallRho
            Case Else: ret = "N/A"
                   
        End Select
        
    Else
        Select Case req
            Case 1: ret = PutPrice
            Case 2: ret = PutDelta
            Case 3: ret = PutGamma
            Case 4: ret = PutTheta
            Case 5: ret = PutVega
            Case 8: ret = PutRho
            Case Else: ret = "N/A"
                   
        End Select
    End If
    
    BSCalc = ret
    
End Function

Function normDist(x, Optional mean = 0, Optional sigma = 1, Optional CDF = True)

    normDist = WorksheetFunction.normDist(x, mean, sigma, CDF)

End Function

Function getIvol(CP, spot, strike, tExp, div, price, hRate, cRate, Optional VolGuess = 1, Optional tolerance = 1e-05, Optional maxIter = 25)
'uses iterative method to solve for implied vol
    
    iter = 0
    
    p = BSCalc(1, CP, spot, strike, tExp, div, VolGuess, hRate, cRate)
    vega = BSCalc(5, CP, spot, strike, tExp, div, VolGuess, hRate, cRate)
    
    Do While Abs((price - p) / spot) >= tolerance
        
        If iter > maxIter Then Exit Do
        
        dV = (price - p) / vega
        
        VolGuess = VolGuess + dV / 100
        
        p = BSCalc(1, CP, spot, strike, tExp, div, VolGuess, hRate, cRate)
        vega = BSCalc(5, CP, spot, strike, tExp, div, VolGuess, hRate, cRate)
        
        iter = iter + 1
        
    Loop
    
    If iter > maxIter Then
        getIvol = "TimedOut"
    Else
        getIvol = VolGuess
    End If
End Function
