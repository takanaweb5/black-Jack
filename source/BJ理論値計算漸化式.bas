Attribute VB_Name = "BJ理論値計算漸化式"
Option Explicit
Private FOpenCard  As Long

'*****************************************************************************
'[概要] オープンカードの確率
'[引数] x:数値
'[戻値] OpenCardの確率
'*****************************************************************************
Public Function OpenRate(ByVal x As Long)
    If Not (1 <= x And x <= 10) Then
        Call Err.Raise(513, , "OpenRate範囲外")
    End If
    
    Dim Result
    If x = FOpenCard Then
        Result = 1
    Else
        Result = 0
    End If
    OpenRate = Result
End Function

'*****************************************************************************
'[概要] xのカードを引く確率
'[引数] x:数値
'[戻値] xのカードを引く確率
'*****************************************************************************
Public Function HitRate(ByVal x As Long)
    If Not (1 <= x And x <= 10) Then
        Call Err.Raise(513, , "HitRate範囲外")
    End If
    
    Dim Result
    If x = 10 Then
        Result = 16 / 52
    Else
        Result = 4 / 52
    End If
    HitRate = Result
End Function

'*****************************************************************************
'[概要] ハードハンドの確率
'[引数] x:数値
'[戻値] ハードハンドxになる確率
'*****************************************************************************
Public Function HRate(ByVal x As Long, Optional lngOpen = 0)
    If lngOpen <> 0 Then FOpenCard = lngOpen

    If Not (4 <= x And x <= 26) Then
        Call Err.Raise(513, , "HRate範囲外")
    End If
    
    Dim Result
    Dim i As Long
    Dim hit As Long
    
    'オープンカードからの確率
    For i = 2 To 10
        hit = x - i
        If 2 <= hit And hit <= 10 Then
            Result = Result + OpenRate(i) * HitRate(hit)
        End If
    Next
    
    'ハードハンドからの確率
    For i = 4 To 16
        hit = x - i
        If hit = 1 Then
            If i + 11 > 21 Then
                'Aを1としてカウント
                Result = Result + HRate(i) * HitRate(1)
            End If
        Else
            If 2 <= hit And hit <= 10 Then
                Result = Result + HRate(i) * HitRate(hit)
            End If
        End If
    Next
        
    'ソフトハンドからの確率
    For i = 12 To 16
        hit = x - i + 10
        If i + hit > 21 Then
            If 6 <= hit And hit <= 10 Then
                Result = Result + SRate(i) * HitRate(hit)
            End If
        End If
    Next
    
    HRate = Result
End Function

'*****************************************************************************
'[概要] ソフトハンドの確率
'[引数] x:数値
'[戻値] ソフトハンドxになる確率
'*****************************************************************************
Public Function SRate(ByVal x As Long, Optional lngOpen = 0)
    If lngOpen <> 0 Then FOpenCard = lngOpen
    
    If Not (12 <= x And x <= 21) Then
        Call Err.Raise(513, , "SRate範囲外")
    End If
    
    Dim Result
    Dim i As Long
    Dim hit As Long
    
    'オープンカードからの確率
    For i = 1 To 10
        If i = 1 Then
            hit = x - 11
            If 1 <= hit And hit <= 10 Then
                Result = Result + OpenRate(i) * HitRate(hit)
            End If
        Else
            hit = x - i
            If hit = 11 Then
                Result = Result + OpenRate(i) * HitRate(1)
            End If
        End If
    Next
    
    'ハードハンドからの確率
    For i = 4 To 10
        hit = x - i
        If hit = 11 Then
            Result = Result + HRate(i) * HitRate(1)
        End If
    Next
        
    'ソフトハンドからの確率
    For i = 12 To 16
        hit = x - i
        If 1 <= hit And hit <= 10 Then
            Result = Result + SRate(i) * HitRate(hit)
        End If
    Next
    
    SRate = Result
End Function

Function 漸化率(ByVal OpenCard As Long, x As Long, Optional blnSoft = False)
    FOpenCard = OpenCard
    If blnSoft Then
        漸化率 = SRate(x)
    Else
        漸化率 = HRate(x)
    End If
End Function
