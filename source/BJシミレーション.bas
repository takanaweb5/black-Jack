Attribute VB_Name = "BJシミレーション"
Option Explicit

Private Cards() As Long '山のカードの配列
Private CardPoint As Long '山の何枚目のカードか
Private HardHandRate(1 To 26) As Double 'ハードハンドの確率の計算結果
Const LOOP回数 = 10000

'*****************************************************************************
'[概要] 親のハンドの確率をシュミレーションし配列数式で返す
'[引数] OpenCard:親の初手(0の時は、初手の出現確率も含めて計算する)
'       Decks:トランプの組数
'[戻値] 配列数式
'*****************************************************************************
Public Function SimulateHandRate(ByVal OpenCard As Long, ByVal Decks As Double) As Variant
    Dim Hands(17 To 22) As Long
    Call SetHands(OpenCard, Decks, Hands())
    
    Dim Result(17 To 22)
    Dim i As Long
    For i = 17 To 22
        Result(i) = Hands(i) / LOOP回数
    Next
    SimulateHandRate = Result
End Function

'*****************************************************************************
'[概要] LOOP回数試行した各ハンドの出現回数を設定する
'[引数] OpenCard:初手(0の時は、初手を決めずに試行する)
'       Decks:トランプの組数
'       Result:各ハンドの出現回数
'*****************************************************************************
Private Sub SetHands(ByVal OpenCard As Long, ByVal Decks As Double, ByRef Result() As Long)
    Call Initilize(Decks)
    Call Shuffle(OpenCard)
    
    Dim i As Long
    Dim Hand As Long
    For i = 1 To LOOP回数
        '毎回シャッフルすると処理が重いためカードの山を半分まで使用するとシャッフルする
        If CardPoint >= UBound(Cards) * 0.5 Then
            Call Shuffle(OpenCard)
        End If
        Hand = Deal(OpenCard)
        Result(Hand) = Result(Hand) + 1
    Next
End Sub

'*****************************************************************************
'[概要] カードの山を作成する
'[引数] Decks:トランプの組数
'*****************************************************************************
Private Sub Initilize(ByVal Decks As Double)
    ReDim Cards(1 To Decks * 52) As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    For i = 1 To Decks * 4
        For j = 1 To 13
            k = k + 1
            Cards(k) = j
        Next
    Next
End Sub

'*****************************************************************************
'[概要] カードをシャッフルする
'[引数] OpenCard:1枚目のカード(0の時は、1枚目も含めてシャッフル)
'*****************************************************************************
Private Sub Shuffle(ByVal FirstCard As Long)
    Dim i As Long
    If FirstCard = 0 Then
        CardPoint = 1
    Else
        CardPoint = 2
        '1枚目のカードをOpenCardに固定する
        For i = 1 To UBound(Cards)
            If Cards(i) = FirstCard Then
                '1枚目とi枚目を交換
                Cards(i) = Cards(1)
                Cards(1) = FirstCard
                Exit For
            End If
        Next
    End If
    
    Dim Swap As Long
    Dim j As Long
    For i = CardPoint To UBound(Cards)
        j = WorksheetFunction.RandBetween(CardPoint, UBound(Cards))
        Swap = Cards(i)
        Cards(i) = Cards(j)
        Cards(j) = Swap
    Next
End Sub

'*****************************************************************************
'[概要] 17以上になるまでカードを引く
'[引数] OpenCard:オープンカード(0の時は、初手を決めずにカードを引く)
'[戻値] 17～22のいずれか（22はバースト）
'*****************************************************************************
Private Function Deal(ByVal OpenCard As Long) As Long
    Dim Hand As Long
    Dim Card As Long
    Dim IsSoftHand As Boolean
    
    If OpenCard = 1 Then
        IsSoftHand = True
        Hand = 11
    Else
        IsSoftHand = False
        Hand = OpenCard
    End If
    
    Do While (Hand < 17)
        Card = Hit()
        Hand = Hand + Card
        If Card = 1 And IsSoftHand = False Then
            IsSoftHand = True
            Hand = Hand + 10
        End If
        If Hand > 21 And IsSoftHand = True Then
            IsSoftHand = False
            Hand = Hand - 10
        End If
    Loop
    
    If Hand > 21 Then
        Deal = 22
    Else
        Deal = Hand
    End If
End Function

'*****************************************************************************
'[概要] カードを1枚引く
'[引数] なし
'[戻値] 引いたカード
'*****************************************************************************
Private Function Hit() As Long
    Hit = Cards(CardPoint)
    If Hit > 10 Then
        Hit = 10
    End If
    CardPoint = CardPoint + 1
End Function




