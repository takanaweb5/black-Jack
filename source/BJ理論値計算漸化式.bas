Attribute VB_Name = "BJ理論値計算漸化式"
Option Explicit
Private FOpenCard  As Long

'*****************************************************************************
'[概要] オープンカードの確率
'[引数] x:数値
'[戻値] OpenCardの確率
'*****************************************************************************
Public Function OpenRate(ByVal x As Long)
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

'    If Not (4 <= x And x <= 26) Then
'        HRate = 99
'        Exit Function
'    End If
    
    Dim Result
    Dim i As Long
    Dim old As Long
    
    'Aから順番にHitする想定
    For i = 1 To 10
        'オープンカードからの確率
        old = x - i
        If 2 <= old And old <= 10 Then
            If i <> 1 Then 'A以外
                Result = Result + OpenRate(old) * HitRate(i)
            Else
                'ソフトハンド
            End If
        End If
    
        'ハードハンドからの確率
        old = x - i
        If 4 <= old And old <= 16 Then
            If i = 1 Then
                If old >= 11 Then
                    'Aを1としてカウント
                    Result = Result + HRate(old) * HitRate(i)
                Else
                    'ソフトハンド
                End If
            Else
                Result = Result + HRate(old) * HitRate(i)
            End If
        End If
        
        'ソフトハンドからの確率
        old = x - i + 10
        If 12 <= old And old <= 16 Then
            If old + i > 21 Then
                Result = Result + SRate(old) * HitRate(i)
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
    
'    If Not (12 <= x And x <= 21) Then
'        SRate = 99
'        Exit Function
'    End If
    
    Dim Result
    Dim i As Long
    Dim old As Long
    
    'Aから順番にHitする想定
    For i = 1 To 10
        'オープンカードからの確率
        old = x - i - 10
        If i = 1 Then '今回がA
            If 1 <= old And old <= 10 Then
                Result = Result + OpenRate(old) * HitRate(i)
            End If
        ElseIf old = 1 Then '前回がA
            Result = Result + OpenRate(old) * HitRate(i)
        End If
        
        'ハードハンドからの確率
        old = x - i - 10
        If i = 1 Then
            If 4 <= old And old <= 10 Then
                Result = Result + HRate(old) * HitRate(i)
            End If
        End If
        
        'ソフトハンドからの確率
        old = x - i
        If 12 <= old And old <= 16 Then
            Result = Result + SRate(old) * HitRate(i)
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
