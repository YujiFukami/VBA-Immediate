Attribute VB_Name = "ModImmediate"
Option Explicit
'�C�~�f�B�G�C�g�E�B���h�E���p�p�̃v���V�[�W��
Sub DPHTest()

    Dim HairetuDummy
    HairetuDummy = Array(Array(1, 2, 3, 4, 5), _
                   Array("A", "AA", "AAA", "AAAA", "AAAAA"), _
                   Array("��", "������", "||||||", "������", "��"))
    HairetuDummy = Application.Transpose(Application.Transpose(HairetuDummy))
    
    Call DPH(HairetuDummy, 3, "�e�X�g1")
    
    Call DPH(HairetuDummy, , "�e�X�g2")

End Sub

Sub DPH(ByVal Hairetu, Optional HyoujiMaxNagasa%, Optional HairetuName$)
    '20210428�ǉ�
    '���͍������p�ɍ쐬
    
    Call DebugPrintHairetu(Hairetu, HyoujiMaxNagasa, HairetuName)
End Sub
Sub DebugPrintHairetu(ByVal Hairetu, Optional HyoujiMaxNagasa%, Optional HairetuName$)
    '20201023�ǉ�
    '�񎟌��z����C�~�f�B�G�C�g�E�B���h�E�Ɍ��₷���\������
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim TateMin&, TateMax&, YokoMin&, YokoMax& '�z��̏c���C���f�b�N�X�ő�ŏ�
    Dim WithTableHairetu '�e�[�u���t�z��c�C�~�f�B�G�C�g�E�B���h�E�ɕ\������ۂɃC���f�b�N�X�ԍ���\�������e�[�u����ǉ������z��
    Dim NagasaList, MaxNagasaList '�e�����̕����񒷂����i�[�A�e��ł̕����񒷂��̍ő�l���i�[
    Dim NagasaOnajiList '" "�i���p�X�y�[�X�j�𕶎���ɒǉ����Ċe��ŕ����񒷂��𓯂��ɂ�����������i�[
    Dim OutputList '�C�~�f�B�G�C�g�E�B���h�E�ɕ\�����镶������i�[
    Const SikiriMoji$ = "|" '�C�~�f�B�G�C�g�E�B���h�E�ɕ\�����鎞�Ɋe��̊Ԃɕ\������u�d�؂蕶���v
    
    '������������������������������������������������������
    '���͈����̏���
    Dim Jigen2%
    On Error Resume Next
    Jigen2 = UBound(Hairetu, 2)
    On Error GoTo 0
    If Jigen2 = 0 Then '1�����z���2�����z��ɂ���
        Hairetu = Application.Transpose(Hairetu)
    End If
    
    TateMin = LBound(Hairetu, 1) '�z��̏c�ԍ��i�C���f�b�N�X�j�̍ŏ�
    TateMax = UBound(Hairetu, 1) '�z��̏c�ԍ��i�C���f�b�N�X�j�̍ő�
    YokoMin = LBound(Hairetu, 2) '�z��̉��ԍ��i�C���f�b�N�X�j�̍ŏ�
    YokoMax = UBound(Hairetu, 2) '�z��̉��ԍ��i�C���f�b�N�X�j�̍ő�
    
    '�e�[�u���t���z��̍쐬
    ReDim WithTableHairetu(1 To TateMax - TateMin + 1 + 1, 1 To YokoMax - YokoMin + 1 + 1) '�e�[�u���ǉ��̕���"+1"����B
    '�uTateMax -TateMin + 1�v�͓��͂����uHairetu�v�̏c�C���f�b�N�X��
    '�uYokoMax -YokoMin + 1�v�͓��͂����uHairetu�v�̉��C���f�b�N�X��
    
    For I = 1 To TateMax - TateMin + 1
        WithTableHairetu(I + 1, 1) = TateMin + I - 1 '�c�e�[�u���iHairetu�̏c�C���f�b�N�X�ԍ��j
        For J = 1 To YokoMax - YokoMin + 1
            WithTableHairetu(1, J + 1) = YokoMin + J - 1 '���e�[�u���iHairetu�̉��C���f�b�N�X�ԍ��j
            WithTableHairetu(I + 1, J + 1) = Hairetu(I - 1 + TateMin, J - 1 + YokoMin) 'Hairetu�̒��̒l
        Next J
    Next I
    
    '������������������������������������������������������
    '�C�~�f�B�G�C�g�E�B���h�E�ɕ\������Ƃ��Ɋe��̕��𓯂��ɐ����邽�߂�
    '�����񒷂��Ƃ��̊e��̍ő�l���v�Z����B
    '�ȉ��ł́uHairetu�v�͈��킸�A�uWithTableHairetu�v�������B
    N = UBound(WithTableHairetu, 1) '�uWithTableHairetu�v�̏c�C���f�b�N�X���i�s���j
    M = UBound(WithTableHairetu, 2) '�uWithTableHairetu�v�̉��C���f�b�N�X���i�񐔁j
    ReDim NagasaList(1 To N, 1 To M)
    ReDim MaxNagasaList(1 To M)
    
    Dim TmpStr$
    For J = 1 To M
        For I = 1 To N
        
            If J > 1 And HyoujiMaxNagasa <> 0 Then
                '�ő�\���������w�肳��Ă���ꍇ�B
                '1��ڂ̃e�[�u���͂��̂܂܂ɂ���B
                TmpStr = WithTableHairetu(I, J)
                WithTableHairetu(I, J) = ��������w��o�C�g���������ɏȗ�(TmpStr, HyoujiMaxNagasa)
            End If
            
            NagasaList(I, J) = LenB(StrConv(WithTableHairetu(I, J), vbFromUnicode)) '�S�p�Ɣ��p����ʂ��Ē������v�Z����B
            MaxNagasaList(J) = WorksheetFunction.Max(MaxNagasaList(J), NagasaList(I, J))
            
        Next I
    Next J
    
    '������������������������������������������������������
    '�C�~�f�B�G�C�g�E�B���h�E�ɕ\�����邽�߂�" "(���p�X�y�[�X)��ǉ�����
    '�����񒷂��𓯂��ɂ���B
    ReDim NagasaOnajiList(1 To N, 1 To M)
    Dim TmpMaxNagasa&
    
    For J = 1 To M
        TmpMaxNagasa = MaxNagasaList(J) '���̗�̍ő啶���񒷂�
        For I = 1 To N
            'Rept�c�w�蕶������w����A�����ĂȂ�����������o�͂���B
            '�i�ő啶����-�������j�̕�" "�i���p�X�y�[�X�j�����ɂ�������B
            NagasaOnajiList(I, J) = WithTableHairetu(I, J) & WorksheetFunction.Rept(" ", TmpMaxNagasa - NagasaList(I, J))
       
        Next I
    Next J
    
    '������������������������������������������������������
    '�C�~�f�B�G�C�g�E�B���h�E�ɕ\�����镶������쐬
    ReDim OutputList(1 To N)
    For I = 1 To N
        For J = 1 To M
            If J = 1 Then
                OutputList(I) = NagasaOnajiList(I, J)
            Else
                OutputList(I) = OutputList(I) & SikiriMoji & NagasaOnajiList(I, J)
            End If
        Next J
    Next I
    
    ''������������������������������������������������������
    '�C�~�f�B�G�C�g�E�B���h�E�ɕ\��
    Debug.Print HairetuName
    For I = 1 To N
        Debug.Print OutputList(I)
    Next I
    
End Sub
Function ��������w��o�C�g���������ɏȗ�(Mojiretu$, ByteNum%)
    '20201023�ǉ�
    '��������w��ȗ��o�C�g�������܂ł̒����ŏȗ�����B
    '�ȗ����ꂽ������̍Ō�̕�����"."�ɕύX����B
    '��FMojiretu = "鳖���" , ByteNum = 6 �c �o�� = "鳖�.."
    '��FMojiretu = "鳖���" , ByteNum = 7 �c �o�� = "鳖��."
    '��FMojiretu = "鳖�XX�" , ByteNum = 6 �c �o�� = "鳖�X."
    '��FMojiretu = "鳖�XX�" , ByteNum = 7 �c �o�� = "鳖�XX."
    
    Dim OriginByte% '���͂���������uMojiretu�v�̃o�C�g������
    Dim Output '�o�͂���ϐ����i�[
    
    '�uMojiretu�v�̃o�C�g�������v�Z
    OriginByte = LenB(StrConv(Mojiretu, vbFromUnicode))
    
    If OriginByte <= ByteNum Then
        '�uMojiretu�v�̃o�C�g�������v�Z���ȗ�����o�C�g�������ȉ��Ȃ�
        '�ȗ��͂��Ȃ�
        Output = Mojiretu
    Else
    
        Dim RuikeiByteList, BunkaiMojiretu
        RuikeiByteList = ������̊e�����݌v�o�C�g���v�Z(Mojiretu)
        BunkaiMojiretu = �����񕪉�(Mojiretu)
        
        Dim AddMoji$
        AddMoji = "."
        
        Dim I&, N&
        N = Len(Mojiretu)
        
        For I = 1 To N
            If RuikeiByteList(I) < ByteNum Then
                Output = Output & BunkaiMojiretu(I)
                
            ElseIf RuikeiByteList(I) = ByteNum Then
                If LenB(StrConv(BunkaiMojiretu(I), vbFromUnicode)) = 1 Then
                    '��FMojiretu = "鳖���" , ByteNum = 6 ,RuikeiByteList(3) = 6
                    'Output = "鳖�.."
                    Output = Output & AddMoji
                Else
                    '��FMojiretu = "鳖�XX�" , ByteNum = 6 ,RuikeiByteList(4) = 6
                    'Output = "鳖�X."
                    Output = Output & AddMoji & AddMoji
                End If
                
                Exit For
                
            ElseIf RuikeiByteList(I) > ByteNum Then
                '��FMojiretu = "鳖���" , ByteNum = 7 ,RuikeiByteList(4) = 8
                'Output = "鳖��."
                Output = Output & AddMoji
                Exit For
            End If
        Next I
        
    End If
        
    ��������w��o�C�g���������ɏȗ� = Output

    
End Function
Function ������̊e�����݌v�o�C�g���v�Z(Mojiretu$)
    '20201023�ǉ�

    '�������1�������ɕ������āA�e�����̃o�C�g���������v�Z���A
    '���̗݌v�l���v�Z����B
    '��FMojiretu="�V�^EK���S��"
    '�o�́�Output = (2,4,5,6,7,10,12)
    
    Dim MojiKosu%
    MojiKosu = Len(Mojiretu)
    
    Dim Output
    ReDim Output(1 To MojiKosu)
    
    Dim I&
    Dim TmpMoji$
    
    For I = 1 To MojiKosu
        TmpMoji = Mid(Mojiretu, I, 1)
        If I = 1 Then
            Output(I) = LenB(StrConv(TmpMoji, vbFromUnicode))
        Else
            Output(I) = LenB(StrConv(TmpMoji, vbFromUnicode)) + Output(I - 1)
        End If
    Next I
    
    ������̊e�����݌v�o�C�g���v�Z = Output
    
End Function
Function �����񕪉�(Mojiretu$)
    '20201023�ǉ�

    '�������1�������������Ĕz��Ɋi�[
    Dim I&, N&
    Dim Output
    
    N = Len(Mojiretu)
    ReDim Output(1 To N)
    For I = 1 To N
        Output(I) = Mid(Mojiretu, I, 1)
    Next I
    
    �����񕪉� = Output
    
End Function

