Attribute VB_Name = "BJ���_�l�v�Z�Q����"
Option Explicit
Private FOpenCard  As Long

'*****************************************************************************
'[�T�v] �I�[�v���J�[�h�̊m��
'[����] x:���l
'[�ߒl] OpenCard�̊m��
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
'[�T�v] x�̃J�[�h�������m��
'[����] x:���l
'[�ߒl] x�̃J�[�h�������m��
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
'[�T�v] �n�[�h�n���h�̊m��
'[����] x:���l
'[�ߒl] �n�[�h�n���hx�ɂȂ�m��
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
    
    'A���珇�Ԃ�Hit����z��
    For i = 1 To 10
        '�I�[�v���J�[�h����̊m��
        old = x - i
        If 2 <= old And old <= 10 Then
            If i <> 1 Then 'A�ȊO
                Result = Result + OpenRate(old) * HitRate(i)
            Else
                '�\�t�g�n���h
            End If
        End If
    
        '�n�[�h�n���h����̊m��
        old = x - i
        If 4 <= old And old <= 16 Then
            If i = 1 Then
                If old >= 11 Then
                    'A��1�Ƃ��ăJ�E���g
                    Result = Result + HRate(old) * HitRate(i)
                Else
                    '�\�t�g�n���h
                End If
            Else
                Result = Result + HRate(old) * HitRate(i)
            End If
        End If
        
        '�\�t�g�n���h����̊m��
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
'[�T�v] �\�t�g�n���h�̊m��
'[����] x:���l
'[�ߒl] �\�t�g�n���hx�ɂȂ�m��
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
    
    'A���珇�Ԃ�Hit����z��
    For i = 1 To 10
        '�I�[�v���J�[�h����̊m��
        old = x - i - 10
        If i = 1 Then '����A
            If 1 <= old And old <= 10 Then
                Result = Result + OpenRate(old) * HitRate(i)
            End If
        ElseIf old = 1 Then '�O��A
            Result = Result + OpenRate(old) * HitRate(i)
        End If
        
        '�n�[�h�n���h����̊m��
        old = x - i - 10
        If i = 1 Then
            If 4 <= old And old <= 10 Then
                Result = Result + HRate(old) * HitRate(i)
            End If
        End If
        
        '�\�t�g�n���h����̊m��
        old = x - i
        If 12 <= old And old <= 16 Then
            Result = Result + SRate(old) * HitRate(i)
        End If
    Next
    
    SRate = Result
End Function

Function �Q����(ByVal OpenCard As Long, x As Long, Optional blnSoft = False)
    FOpenCard = OpenCard
    If blnSoft Then
        �Q���� = SRate(x)
    Else
        �Q���� = HRate(x)
    End If
End Function
