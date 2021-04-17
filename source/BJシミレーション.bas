Attribute VB_Name = "BJ�V�~���[�V����"
Option Explicit

Private Cards() As Long '�R�̃J�[�h�̔z��
Private CardPoint As Long '�R�̉����ڂ̃J�[�h��
Private HardHandRate(1 To 26) As Double '�n�[�h�n���h�̊m���̌v�Z����

'*****************************************************************************
'[�T�v] �e�̃n���h�̊m�����V���~���[�V�������z�񐔎��ŕԂ�
'[����] LopCnt:���[�v��
'       Decks:�g�����v�̑g��
'[�ߒl] �z�񐔎�
'*****************************************************************************
Public Function SimulateHandRate(ByVal LopCnt As Long, ByVal Decks As Double) As Variant
    Dim Result(1 To 10, 17 To 22) As Long
    Call SetHands(LopCnt, Decks, Result())
    SimulateHandRate = Result
End Function

'*****************************************************************************
'[�T�v] LOOP�񐔎��s�����e�n���h�̏o���񐔂�ݒ肷��
'[����] LopCnt:���[�v��
'       Decks:�g�����v�̑g��
'       Result:�e�n���h�̏o����
'*****************************************************************************
Private Sub SetHands(ByVal LopCnt As Long, ByVal Decks As Double, ByRef Result() As Long)
    Call Initilize(Decks)
    Call Shuffle
    
    Dim i As Long
    Dim Hand As Long
    For i = 1 To LopCnt
        '����V���b�t������Ə������d�����߃J�[�h�̎R��4����3�܂Ŏg�p����ƃV���b�t������
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
'[����] �Ȃ�
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
    CardPoint = CardPoint + 1
    Hit = Cards(CardPoint)
    If Hit > 10 Then
        Hit = 10
    End If
End Function
