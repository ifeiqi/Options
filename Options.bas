Attribute VB_Name = "Module1"
Function d1(S, K, T, r, volatility, dividend)
    d1 = (Log(S / K) + (r - dividend + 0.5 * volatility ^ 2) * T) / (volatility * (Sqr(T)))
End Function

Function Nd1(S, K, T, r, volatility, dividend)
    Nd1 = Application.NormSDist(d1(S, K, T, r, volatility, dividend))
End Function

Function Nprimed1(S, K, T, r, volatility, dividend)
    Nprimed1 = Exp(-0.5 * d1(S, K, T, r, volatility, dividend) ^ 2) / Sqr(2 * WorksheetFunction.Pi())
End Function

Function d2(S, K, T, r, volatility, dividend)
    d2 = d1(S, K, T, r, volatility, dividend) - volatility * Sqr(T)
End Function

Function Nd2(S, K, T, r, volatility, dividend)
    Nd2 = Application.NormSDist(d2(S, K, T, r, volatility, dividend))
End Function

Function OptionPrice(OptionType, S, K, T, r, volatility, dividend)
    If OptionType = "C" Then
        OptionPrice = Exp(-dividend * T) * S * Nd1(S, K, T, r, volatility, dividend) - K * Exp(-r * T) * Application.NormSDist(d1(S, K, T, r, volatility, dividend) - volatility * Sqr(T))
    ElseIf OptionType = "P" Then
        OptionPrice = K * Exp(-r * T) * (1 - Nd2(S, K, T, r, volatility, dividend)) - Exp(-dividend * T) * S * (1 - Nd1(S, K, T, r, volatility, dividend))
    End If
End Function
 
Function OptionDelta(OptionType, S, K, T, r, volatility, dividend)
    If OptionType = "C" Then
        OptionDelta = Exp(-dividend * T) * Nd1(S, K, T, r, volatility, dividend)
    ElseIf OptionType = "P" Then
        OptionDelta = Exp(-dividend * T) * (Nd1(S, K, T, r, volatility, dividend) - 1)
    End If
End Function

Function OptionGamma(S, K, T, r, volatility, dividend)
    OptionGamma = Exp(-dividend * T) * Nprimed1(S, K, T, r, volatility, dividend) / (S * volatility * Sqr(T))
End Function
 
Function OptionVega(S, K, T, r, volatility, dividend)
    OptionVega = Exp(-dividend * T) * S * Sqr(T) * Nprimed1(S, K, T, r, volatility, dividend)
End Function

Function OptionTheta(OptionType, S, K, T, r, volatility, dividend)
    If OptionType = "C" Then
        OptionTheta = -S * volatility * Nprimed1(S, K, T, r, volatility, dividend) / (2 * Sqr(T)) _
            + Exp(-dividend * T) * dividend * S * Nd1(S, K, T, r, volatility, dividend) _
            - r * K * Exp(-r * (T)) * Nd2(S, K, T, r, volatility, dividend)
    ElseIf OptionType = "P" Then
        OptionTheta = -S * volatility * Exp(-dividend * T) * Nprimed1(S, K, T, r, volatility, dividend) / (2 * Sqr(T)) _
            - Exp(-dividend * T) * dividend * S * (1 - Nd1(S, K, T, r, volatility, dividend)) _
            + r * K * Exp(-r * (T)) * (1 - Nd2(S, K, T, r, volatility, dividend))
    End If
End Function
 
Function OptionRho(OptionType, S, K, T, r, volatility, dividend)
    If OptionType = "C" Then
        OptionRho = K * T * Exp(-r * T) * Nd2(S, K, T, r, volatility, dividend)
    ElseIf OptionType = "P" Then
        OptionRho = -K * T * Exp(-r * T) * (1 - Nd2(S, K, T, r, volatility, dividend))
    End If
End Function

'Newton-Raphson Method iterative estimation of implied volatility

Function ImpliedVolatility(OptionType, S, K, T, r, price, dividend, Optional guess = 0.5)
    dVol = 0.000001
    epsilon = 0.000001
    maxIter = 500
    vol_1 = guess
    i = 1
    Do
       Value_1 = OptionPrice(OptionType, S, K, T, r, vol_1, dividend)
       vol_2 = vol_1 - dVol
       Value_2 = OptionPrice(OptionType, S, K, T, r, vol_2, dividend)
       dx = (Value_2 - Value_1) / dVol
       If Abs(dx) < epsilon Or i = maxIter Then Exit Do
       vol_1 = vol_1 - (price - Value_1) / dx
       i = i + 1
    Loop
    ImpliedVolatility = vol_1
End Function

