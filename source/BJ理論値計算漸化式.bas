Attribute VB_Name = "BJ���_�l�v�Z�Q����"
Option Explicit
Private FOpenCard  As Long

'*****************************************************************************
'[�T�v] �I�[�v���J�[�h�̊m��
'[����] x:���l
'[�ߒl] OpenCard�̊m��
'*****************************************************************************
Public Function OpenRate(ByVal x As Long)
    If Not (1 <= x And x <= 10) Then
        Call Err.Raise(513, , "OpenRate�͈͊O")
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
'[�T�v] x�̃J�[�h�������m��
'[����] x:���l
'[�ߒl] x�̃J�[�h�������m��
'*****************************************************************************
Public Function HitRate(ByVal x As Long)
    If Not (1 <= x And x <= 10) Then
        Call Err.Raise(513, , "HitRate�͈͊O")
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
'[�T�v] �n�[�h�n���h�̊m��
'[����] x:���l
'[�ߒl] �n�[�h�n���hx�ɂȂ�m��
'*****************************************************************************
Public Function HRate(ByVal x As Long, Optional lngOpen = 0)
    If lngOpen <> 0 Then FOpenCard = lngOpen

    If Not (4 <= x And x <= 26) Then
        Call Err.Raise(513, , "HRate�͈͊O")
    End If
    
    Dim Result
    Dim i As Long
    Dim hit As Long
    
    '�I�[�v���J�[�h����̊m��
    For i = 2 To 10
        hit = x - i
        If 2 <= hit And hit <= 10 Then
            Result = Result + OpenRate(i) * HitRate(hit)
        End If
    Next
    
    '�n�[�h�n���h����̊m��
    For i = 4 To 16
        hit = x - i
        If hit = 1 Then
            If i + 11 > 21 Then
                'A��1�Ƃ��ăJ�E���g
                Result = Result + HRate(i) * HitRate(1)
            End If
        Else
            If 2 <= hit And hit <= 10 Then
                Result = Result + HRate(i) * HitRate(hit)
            End If
        End If
    Next
        
    '�\�t�g�n���h����̊m��
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
'[�T�v] �\�t�g�n���h�̊m��
'[����] x:���l
'[�ߒl] �\�t�g�n���hx�ɂȂ�m��
'*****************************************************************************
Public Function SRate(ByVal x As Long, Optional lngOpen = 0)
    If lngOpen <> 0 Then FOpenCard = lngOpen
    
    If Not (12 <= x And x <= 21) Then
        Call Err.Raise(513, , "SRate�͈͊O")
    End If
    
    Dim Result
    Dim i As Long
    Dim hit As Long
    
    '�I�[�v���J�[�h����̊m��
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
    
    '�n�[�h�n���h����̊m��
    For i = 4 To 10
        hit = x - i
        If hit = 11 Then
            Result = Result + HRate(i) * HitRate(1)
        End If
    Next
        
    '�\�t�g�n���h����̊m��
    For i = 12 To 16
        hit = x - i
        If 1 <= hit And hit <= 10 Then
            Result = Result + SRate(i) * HitRate(hit)
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
