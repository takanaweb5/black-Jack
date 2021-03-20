Attribute VB_Name = "BJ�V�~���[�V����"
Option Explicit

Private Cards() As Long '�R�̃J�[�h�̔z��
Private CardPoint As Long '�R�̉����ڂ̃J�[�h��
Private HardHandRate(1 To 26) As Double '�n�[�h�n���h�̊m���̌v�Z����
Const LOOP�� = 10000

'*****************************************************************************
'[�T�v] �e�̃n���h�̊m�����V���~���[�V�������z�񐔎��ŕԂ�
'[����] OpenCard:�e�̏���(0�̎��́A����̏o���m�����܂߂Čv�Z����)
'       Decks:�g�����v�̑g��
'[�ߒl] �z�񐔎�
'*****************************************************************************
Public Function SimulateHandRate(ByVal OpenCard As Long, ByVal Decks As Double) As Variant
    Dim Hands(17 To 22) As Long
    Call SetHands(OpenCard, Decks, Hands())
    
    Dim Result(17 To 22)
    Dim i As Long
    For i = 17 To 22
        Result(i) = Hands(i) / LOOP��
    Next
    SimulateHandRate = Result
End Function

'*****************************************************************************
'[�T�v] LOOP�񐔎��s�����e�n���h�̏o���񐔂�ݒ肷��
'[����] OpenCard:����(0�̎��́A��������߂��Ɏ��s����)
'       Decks:�g�����v�̑g��
'       Result:�e�n���h�̏o����
'*****************************************************************************
Private Sub SetHands(ByVal OpenCard As Long, ByVal Decks As Double, ByRef Result() As Long)
    Call Initilize(Decks)
    Call Shuffle(OpenCard)
    
    Dim i As Long
    Dim Hand As Long
    For i = 1 To LOOP��
        '����V���b�t������Ə������d�����߃J�[�h�̎R�𔼕��܂Ŏg�p����ƃV���b�t������
        If CardPoint >= UBound(Cards) * 0.5 Then
            Call Shuffle(OpenCard)
        End If
        Hand = Deal(OpenCard)
        Result(Hand) = Result(Hand) + 1
    Next
End Sub

'*****************************************************************************
'[�T�v] �J�[�h�̎R���쐬����
'[����] Decks:�g�����v�̑g��
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
'[�T�v] �J�[�h���V���b�t������
'[����] OpenCard:1���ڂ̃J�[�h(0�̎��́A1���ڂ��܂߂ăV���b�t��)
'*****************************************************************************
Private Sub Shuffle(ByVal FirstCard As Long)
    Dim i As Long
    If FirstCard = 0 Then
        CardPoint = 1
    Else
        CardPoint = 2
        '1���ڂ̃J�[�h��OpenCard�ɌŒ肷��
        For i = 1 To UBound(Cards)
            If Cards(i) = FirstCard Then
                '1���ڂ�i���ڂ�����
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
'[�T�v] 17�ȏ�ɂȂ�܂ŃJ�[�h������
'[����] OpenCard:�I�[�v���J�[�h(0�̎��́A��������߂��ɃJ�[�h������)
'[�ߒl] 17�`22�̂����ꂩ�i22�̓o�[�X�g�j
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
'[�T�v] �J�[�h��1������
'[����] �Ȃ�
'[�ߒl] �������J�[�h
'*****************************************************************************
Private Function Hit() As Long
    Hit = Cards(CardPoint)
    If Hit > 10 Then
        Hit = 10
    End If
    CardPoint = CardPoint + 1
End Function




