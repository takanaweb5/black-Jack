Attribute VB_Name = "BJ���_�l�v�Z"
Option Explicit

Private Cards(1 To 10) As Long 'A�`10�J�[�h�̎R�̖���
Private CardsCount As Long '�J�[�h�̃g�[�^������
Private HardHandRate(1 To 26) As Double '�n�[�h�n���h�̊m���̌v�Z����
Private SoftHandRate(1 To 26) As Double '�\�t�g�n���h�̊m���̌v�Z����

'*****************************************************************************
'[�T�v] �e�̃n���h�̊m����z�񐔎��ŕԂ�
'[����] OpenCard:�e�̏���(0�̎��́A����̏o���m�����܂߂Čv�Z����)
'       Decks:�g�����v�̑g��
'[�ߒl] �z�񐔎�
'*****************************************************************************
Public Function CalcHandRate(ByVal OpenCard As Long, ByVal Decks As Double) As Variant
    Call Initilize(Decks)
    Call SetHand(OpenCard)
    
    Dim i As Long
    Dim Result(17 To 22)
    For i = 17 To 21
        Result(i) = HardHandRate(i) + SoftHandRate(i)
    Next
    '�o�[�X�g�̊m����22�ɐݒ�
    Result(22) = HardHandRate(22) + SoftHandRate(22) _
               + HardHandRate(23) + SoftHandRate(23) _
               + HardHandRate(24) + SoftHandRate(24) _
               + HardHandRate(25) + SoftHandRate(25) _
               + HardHandRate(26) + SoftHandRate(26)
    CalcHandRate = Result
End Function

'*****************************************************************************
'[�T�v] �n�[�h�n���h�̊m����z�񐔎��ŕԂ�
'[����] OpenCard:�e�̏���(0�̎��́A����̏o���m�����܂߂Čv�Z����)
'       Decks:�g�����v�̑g��
'[�ߒl] �z�񐔎�
'*****************************************************************************
Public Function CalcHardHandRate(ByVal OpenCard As Long, ByVal Decks As Double) As Variant
    Call Initilize(Decks)
    Call SetHand(OpenCard)
    
    Dim i As Long
    Dim Result(1 To 22)
    For i = 1 To 21
        Result(i) = HardHandRate(i)
    Next
    '�o�[�X�g�̊m����22�ɐݒ�
    Result(22) = HardHandRate(22) + HardHandRate(23) + _
                 HardHandRate(24) + HardHandRate(25) + HardHandRate(26)
    CalcHardHandRate = Result
End Function

'*****************************************************************************
'[�T�v] �\�t�g�n���h�̊m����z�񐔎��ŕԂ�
'[����] OpenCard:�e�̏���(0�̎��́A����̏o���m�����܂߂Čv�Z����)
'       Decks:�g�����v�̑g��
'[�ߒl] �z�񐔎�
'*****************************************************************************
Public Function CalcSoftHandRate(ByVal OpenCard As Long, ByVal Decks As Double) As Variant
    Call Initilize(Decks)
    Call SetHand(OpenCard)
    
    Dim i As Long
    Dim Result(1 To 22)
    For i = 1 To 21
        Result(i) = SoftHandRate(i)
    Next
    '�o�[�X�g�̊m����22�ɐݒ�
    Result(22) = SoftHandRate(22) + SoftHandRate(23) + _
                 SoftHandRate(24) + SoftHandRate(25) + SoftHandRate(26)
    CalcSoftHandRate = Result
End Function

'*****************************************************************************
'[�T�v] �z��̏������Ȃ�
'[����] Decks:�g�����v�̑g��
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
'[�T�v] �e�n���h�̊m����ݒ肷��
'[����] �e�̏���(0�̎��́A����̏o���m�����܂߂Čv�Z����)
'*****************************************************************************
Private Sub SetHand(ByVal OpenCard As Long)
    If OpenCard = 0 Then
        '�I�[�v���J�[�h�̏o���m�����܂߂Ċm�����v�Z���鎞
        Dim i As Long
        Dim Rate As Double
        For i = 1 To 10
            Rate = Cards(i) / CardsCount 'i�̃J�[�h�̏o���m�����v�Z
            Call DecCard(i) 'i�̃J�[�h���R����1�����炷
            If i = 1 Then
                'A�̎��́A�\�t�g�n���h��11�Ƃ݂Ȃ��Čv�Z����
                Call SetHandRate(11, Rate, True)
            Else
                Call SetHandRate(i, Rate, False)
            End If
            Call IncCard(i) 'i�̃J�[�h���R�ɖ߂�
        Next
        Exit Sub
    End If
    
    Call DecCard(OpenCard) '�I�[�v���J�[�h���R����1�����炷
    If OpenCard = 1 Then
        'A�̎��́A�\�t�g�n���h��11�Ƃ݂Ȃ��Čv�Z����
        Call SetHandRate(11, 1, True)
    Else
        Call SetHandRate(OpenCard, 1, False)
    End If
End Sub

'*****************************************************************************
'[�T�v] �e�n���h�̊m����ݒ肷��(���d���[�v���ċA�֐��Ŏ�������)
'[����] Hand:���݂̎�,HandRate:���݂̎�̏o���m��,IsSoftHand:�\�t�g�n���h�̎�True
'*****************************************************************************
Private Sub SetHandRate(ByVal Hand As Long, ByVal HandRate As Double, ByVal IsSoftHand As Boolean)
    Dim i As Long
    Dim Rate As Double
    Dim NextHand As Long
    Dim NextIsSoftHand As Boolean

    For i = 1 To 10
        '�R�ɑΏۂ̃J�[�h���c���Ă��邩����
        If Cards(i) > 0 Then
            NextIsSoftHand = IsSoftHand
            NextHand = Hand + i
            If i = 1 And NextIsSoftHand = False Then
                '�n�[�h�n���h���\�t�g�n���h�ɕύX
                NextIsSoftHand = True
                NextHand = NextHand + 10
            End If
            If NextHand > 21 And NextIsSoftHand = True Then
                '�\�t�g�n���h���n�[�h�n���h�ɕύX
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
                '17�����Ȃ玟�̃J�[�h������
                Call DecCard(i) 'i�̃J�[�h���R����1�����炷
                Call SetHandRate(NextHand, Rate, NextIsSoftHand)
                Call IncCard(i) 'i�̃J�[�h���R�ɖ߂�
            End If
        End If
    Next
End Sub

'*****************************************************************************
'[�T�v] �J�[�h���R����1�����炷
'[����] �Ώۂ̃J�[�h
'*****************************************************************************
Private Sub DecCard(ByVal Card As Long)
'    Exit Sub  '�g�p�ς݃J�[�h�̏o�����̌������l�����Ȃ��ꍇ
    Cards(Card) = Cards(Card) - 1
    CardsCount = CardsCount - 1
End Sub

'*****************************************************************************
'[�T�v] �J�[�h���R�ɖ߂�
'[����] �Ώۂ̃J�[�h
'*****************************************************************************
Private Sub IncCard(ByVal Card As Long)
'    Exit Sub  '�g�p�ς݃J�[�h�̏o�����̌������l�����Ȃ��ꍇ
    Cards(Card) = Cards(Card) + 1
    CardsCount = CardsCount + 1
End Sub
