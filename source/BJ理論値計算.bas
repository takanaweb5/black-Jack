Attribute VB_Name = "BJ理論値計算"
Option Explicit

Private Cards(1 To 10) As Long 'A～10カードの山の枚数
Private CardsCount As Long 'カードのトータル枚数
Private HardHandRate(1 To 26) As Double 'ハードハンドの確率の計算結果
Private SoftHandRate(1 To 26) As Double 'ソフトハンドの確率の計算結果

'*****************************************************************************
'[概要] 親のハンドの確率を配列数式で返す
'[引数] OpenCard:親の初手(0の時は、初手の出現確率も含めて計算する)
'       Decks:トランプの組数
'[戻値] 配列数式
'*****************************************************************************
Public Function CalcHandRate(ByVal OpenCard As Long, ByVal Decks As Double) As Variant
    Call Initilize(Decks)
    Call SetHand(OpenCard)
    
    Dim i As Long
    Dim Result(17 To 22)
    For i = 17 To 21
        Result(i) = HardHandRate(i) + SoftHandRate(i)
    Next
    'バーストの確率は22に設定
    Result(22) = HardHandRate(22) + SoftHandRate(22) _
               + HardHandRate(23) + SoftHandRate(23) _
               + HardHandRate(24) + SoftHandRate(24) _
               + HardHandRate(25) + SoftHandRate(25) _
               + HardHandRate(26) + SoftHandRate(26)
    CalcHandRate = Result
End Function

'*****************************************************************************
'[概要] ハードハンドの確率を配列数式で返す
'[引数] OpenCard:親の初手(0の時は、初手の出現確率も含めて計算する)
'       Decks:トランプの組数
'[戻値] 配列数式
'*****************************************************************************
Public Function CalcHardHandRate(ByVal OpenCard As Long, ByVal Decks As Double) As Variant
    Call Initilize(Decks)
    Call SetHand(OpenCard)
    
    Dim i As Long
    Dim Result(1 To 22)
    For i = 1 To 21
        Result(i) = HardHandRate(i)
    Next
    'バーストの確率は22に設定
    Result(22) = HardHandRate(22) + HardHandRate(23) + _
                 HardHandRate(24) + HardHandRate(25) + HardHandRate(26)
    CalcHardHandRate = Result
End Function

'*****************************************************************************
'[概要] ソフトハンドの確率を配列数式で返す
'[引数] OpenCard:親の初手(0の時は、初手の出現確率も含めて計算する)
'       Decks:トランプの組数
'[戻値] 配列数式
'*****************************************************************************
Public Function CalcSoftHandRate(ByVal OpenCard As Long, ByVal Decks As Double) As Variant
    Call Initilize(Decks)
    Call SetHand(OpenCard)
    
    Dim i As Long
    Dim Result(1 To 22)
    For i = 1 To 21
        Result(i) = SoftHandRate(i)
    Next
    'バーストの確率は22に設定
    Result(22) = SoftHandRate(22) + SoftHandRate(23) + _
                 SoftHandRate(24) + SoftHandRate(25) + SoftHandRate(26)
    CalcSoftHandRate = Result
End Function

'*****************************************************************************
'[概要] 配列の初期化など
'[引数] Decks:トランプの組数
'*****************************************************************************
Private Sub Initilize(ByVal Decks As Double)
    Erase HardHandRate()
    Erase SoftHandRate()
    
    Dim i As Long
    For i = 1 To 9
        Cards(i) = Decks * 4
    Next
    Cards(10) = Decks * 16
    CardsCount = Decks * 52
End Sub

'*****************************************************************************
'[概要] 各ハンドの確率を設定する
'[引数] 親の初手(0の時は、初手の出現確率も含めて計算する)
'*****************************************************************************
Private Sub SetHand(ByVal OpenCard As Long)
    If OpenCard = 0 Then
        'オープンカードの出現確率も含めて確率を計算する時
        Dim i As Long
        Dim Rate As Double
        For i = 1 To 10
            Rate = Cards(i) / CardsCount 'iのカードの出現確率を計算
            Call DecCard(i) 'iのカードを山から1枚減らす
            If i = 1 Then
                'Aの時は、ソフトハンドの11とみなして計算する
                Call SetHandRate(11, Rate, True)
            Else
                Call SetHandRate(i, Rate, False)
            End If
            Call IncCard(i) 'iのカードを山に戻す
        Next
        Exit Sub
    End If
    
    Call DecCard(OpenCard) 'オープンカードを山から1枚減らす
    If OpenCard = 1 Then
        'Aの時は、ソフトハンドの11とみなして計算する
        Call SetHandRate(11, 1, True)
    Else
        Call SetHandRate(OpenCard, 1, False)
    End If
End Sub

'*****************************************************************************
'[概要] 各ハンドの確率を設定する(多重ループを再帰関数で実現する)
'[引数] Hand:現在の手,HandRate:現在の手の出現確率,IsSoftHand:ソフトハンドの時True
'*****************************************************************************
Private Sub SetHandRate(ByVal Hand As Long, ByVal HandRate As Double, ByVal IsSoftHand As Boolean)
    Dim i As Long
    Dim Rate As Double
    Dim NextHand As Long
    Dim NextIsSoftHand As Boolean

    For i = 1 To 10
        '山に対象のカードが残っているか判定
        If Cards(i) > 0 Then
            NextIsSoftHand = IsSoftHand
            NextHand = Hand + i
            If i = 1 And NextIsSoftHand = False Then
                'ハードハンドをソフトハンドに変更
                NextIsSoftHand = True
                NextHand = NextHand + 10
            End If
            If NextHand > 21 And NextIsSoftHand = True Then
                'ソフトハンドをハードハンドに変更
                NextIsSoftHand = False
                NextHand = NextHand - 10
            End If
            
            Rate = HandRate * Cards(i) / CardsCount
            If NextIsSoftHand Then
                SoftHandRate(NextHand) = SoftHandRate(NextHand) + Rate
            Else
                HardHandRate(NextHand) = HardHandRate(NextHand) + Rate
            End If
            If NextHand < 17 Then
                '17未満なら次のカードを引く
                Call DecCard(i) 'iのカードを山から1枚減らす
                Call SetHandRate(NextHand, Rate, NextIsSoftHand)
                Call IncCard(i) 'iのカードを山に戻す
            End If
        End If
    Next
End Sub

'*****************************************************************************
'[概要] カードを山から1枚減らす
'[引数] 対象のカード
'*****************************************************************************
Private Sub DecCard(ByVal Card As Long)
'    Exit Sub  '使用済みカードの出現率の減少を考慮しない場合
    Cards(Card) = Cards(Card) - 1
    CardsCount = CardsCount - 1
End Sub

'*****************************************************************************
'[概要] カードを山に戻す
'[引数] 対象のカード
'*****************************************************************************
Private Sub IncCard(ByVal Card As Long)
'    Exit Sub  '使用済みカードの出現率の減少を考慮しない場合
    Cards(Card) = Cards(Card) + 1
    CardsCount = CardsCount + 1
End Sub
