Attribute VB_Name = "BJシミレーション"
Option Explicit

Private Cards() As Long '山のカードの配列
Private CardPoint As Long '山の何枚目のカードか
Private HardHandRate(1 To 26) As Double 'ハードハンドの確率の計算結果

'*****************************************************************************
'[概要] 親のハンドの確率をシュミレーションし配列数式で返す
'[引数] LopCnt:ループ回数
'       Decks:トランプの組数
'[戻値] 配列数式
'*****************************************************************************
Public Function SimulateHandRate(ByVal LopCnt As Long, ByVal Decks As Double) As Variant
    Dim Result(1 To 10, 17 To 22) As Long
    Call SetHands(LopCnt, Decks, Result())
    SimulateHandRate = Result
End Function

'*****************************************************************************
'[概要] LOOP回数試行した各ハンドの出現回数を設定する
'[引数] LopCnt:ループ回数
'       Decks:トランプの組数
'       Result:各ハンドの出現回数
'*****************************************************************************
Private Sub SetHands(ByVal LopCnt As Long, ByVal Decks As Double, ByRef Result() As Long)
    Call Initilize(Decks)
    Call Shuffle
    
    Dim i As Long
    Dim Hand As Long
    For i = 1 To LopCnt
        '毎回シャッフルすると処理が重いためカードの山を4分の3まで使用するとシャッフルする
        '
'        If CardPoint >= UBound(Cards) * 0.75 Then
            Call Shuffle
'        End If
        
        Dim OpenCard As Long
        OpenCard = Hit()
        Hand = Deal(OpenCard)
        Result(OpenCard, Hand) = Result(OpenCard, Hand) + 1
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
'[引数] なし
'*****************************************************************************
Private Sub Shuffle()
    CardPoint = 0
    Dim i As Long
    Dim j As Long
    Dim Swap As Long
    For i = 1 To UBound(Cards)
        j = WorksheetFunction.RandBetween(1, UBound(Cards))
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
    CardPoint = CardPoint + 1
    Hit = Cards(CardPoint)
    If Hit > 10 Then
        Hit = 10
    End If
End Function
