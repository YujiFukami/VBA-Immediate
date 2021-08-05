Attribute VB_Name = "ModImmediate"
Option Explicit
'イミディエイトウィンドウ活用用のプロシージャ
Sub DPHTest()

    Dim HairetuDummy
    HairetuDummy = Array(Array(1, 2, 3, 4, 5), _
                   Array("A", "AA", "AAA", "AAAA", "AAAAA"), _
                   Array("あ", "あああ", "||||||", "ううう", "あ"))
    HairetuDummy = Application.Transpose(Application.Transpose(HairetuDummy))
    
    Call DPH(HairetuDummy, , "テスト1") '実行テスト1
    
    Call DPH(HairetuDummy, 3, "テスト2") '実行テスト2(文字の長さを3以内に指定)

End Sub

Sub DPH(ByVal Hairetu, Optional HyoujiMaxNagasa%, Optional HairetuName$)
    '20210428追加
    '入力高速化用に作成
    
    Call DebugPrintHairetu(Hairetu, HyoujiMaxNagasa, HairetuName)
End Sub

Sub DebugPrintHairetu(ByVal Hairetu, Optional HyoujiMaxNagasa%, Optional HairetuName$)
    '20201023追加
    '二次元配列をイミディエイトウィンドウに見やすく表示する
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim TateMin&, TateMax&, YokoMin&, YokoMax& '配列の縦横インデックス最大最小
    Dim WithTableHairetu 'テーブル付配列…イミディエイトウィンドウに表示する際にインデックス番号を表示したテーブルを追加した配列
    Dim NagasaList, MaxNagasaList '各文字の文字列長さを格納、各列での文字列長さの最大値を格納
    Dim NagasaOnajiList '" "（半角スペース）を文字列に追加して各列で文字列長さを同じにした文字列を格納
    Dim OutputList 'イミディエイトウィンドウに表示する文字列を格納
    Const SikiriMoji$ = "|" 'イミディエイトウィンドウに表示する時に各列の間に表示する「仕切り文字」
    
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力引数の処理
    Dim Jigen2%
    On Error Resume Next
    Jigen2 = UBound(Hairetu, 2)
    On Error GoTo 0
    If Jigen2 = 0 Then '1次元配列は2次元配列にする
        Hairetu = Application.Transpose(Hairetu)
    End If
    
    TateMin = LBound(Hairetu, 1) '配列の縦番号（インデックス）の最小
    TateMax = UBound(Hairetu, 1) '配列の縦番号（インデックス）の最大
    YokoMin = LBound(Hairetu, 2) '配列の横番号（インデックス）の最小
    YokoMax = UBound(Hairetu, 2) '配列の横番号（インデックス）の最大
    
    'テーブル付き配列の作成
    ReDim WithTableHairetu(1 To TateMax - TateMin + 1 + 1, 1 To YokoMax - YokoMin + 1 + 1) 'テーブル追加の分で"+1"する。
    '「TateMax -TateMin + 1」は入力した「Hairetu」の縦インデックス数
    '「YokoMax -YokoMin + 1」は入力した「Hairetu」の横インデックス数
    
    For I = 1 To TateMax - TateMin + 1
        WithTableHairetu(I + 1, 1) = TateMin + I - 1 '縦テーブル（Hairetuの縦インデックス番号）
        For J = 1 To YokoMax - YokoMin + 1
            WithTableHairetu(1, J + 1) = YokoMin + J - 1 '横テーブル（Hairetuの横インデックス番号）
            WithTableHairetu(I + 1, J + 1) = Hairetu(I - 1 + TateMin, J - 1 + YokoMin) 'Hairetuの中の値
        Next J
    Next I
    
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    'イミディエイトウィンドウに表示するときに各列の幅を同じに整えるために
    '文字列長さとその各列の最大値を計算する。
    '以下では「Hairetu」は扱わず、「WithTableHairetu」を扱う。
    N = UBound(WithTableHairetu, 1) '「WithTableHairetu」の縦インデックス数（行数）
    M = UBound(WithTableHairetu, 2) '「WithTableHairetu」の横インデックス数（列数）
    ReDim NagasaList(1 To N, 1 To M)
    ReDim MaxNagasaList(1 To M)
    
    Dim TmpStr$
    For J = 1 To M
        For I = 1 To N
        
            If J > 1 And HyoujiMaxNagasa <> 0 Then
                '最大表示長さが指定されている場合。
                '1列目のテーブルはそのままにする。
                TmpStr = WithTableHairetu(I, J)
                WithTableHairetu(I, J) = 文字列を指定バイト数文字数に省略(TmpStr, HyoujiMaxNagasa)
            End If
            
            NagasaList(I, J) = LenB(StrConv(WithTableHairetu(I, J), vbFromUnicode)) '全角と半角を区別して長さを計算する。
            MaxNagasaList(J) = WorksheetFunction.Max(MaxNagasaList(J), NagasaList(I, J))
            
        Next I
    Next J
    
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    'イミディエイトウィンドウに表示するために" "(半角スペース)を追加して
    '文字列長さを同じにする。
    ReDim NagasaOnajiList(1 To N, 1 To M)
    Dim TmpMaxNagasa&
    
    For J = 1 To M
        TmpMaxNagasa = MaxNagasaList(J) 'その列の最大文字列長さ
        For I = 1 To N
            'Rept…指定文字列を指定個数連続してつなげた文字列を出力する。
            '（最大文字数-文字数）の分" "（半角スペース）を後ろにくっつける。
            NagasaOnajiList(I, J) = WithTableHairetu(I, J) & WorksheetFunction.Rept(" ", TmpMaxNagasa - NagasaList(I, J))
       
        Next I
    Next J
    
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    'イミディエイトウィンドウに表示する文字列を作成
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
    
    ''※※※※※※※※※※※※※※※※※※※※※※※※※※※
    'イミディエイトウィンドウに表示
    Debug.Print HairetuName
    For I = 1 To N
        Debug.Print OutputList(I)
    Next I
    
End Sub

Function 文字列を指定バイト数文字数に省略(Mojiretu$, ByteNum%)
    '20201023追加
    '文字列を指定省略バイト文字数までの長さで省略する。
    '省略された文字列の最後の文字は"."に変更する。
    '例：Mojiretu = "魑魅魍魎" , ByteNum = 6 … 出力 = "魑魅.."
    '例：Mojiretu = "魑魅魍魎" , ByteNum = 7 … 出力 = "魑魅魍."
    '例：Mojiretu = "魑魅XX魎" , ByteNum = 6 … 出力 = "魑魅X."
    '例：Mojiretu = "魑魅XX魎" , ByteNum = 7 … 出力 = "魑魅XX."
    
    Dim OriginByte% '入力した文字列「Mojiretu」のバイト文字数
    Dim Output '出力する変数を格納
    
    '「Mojiretu」のバイト文字数計算
    OriginByte = LenB(StrConv(Mojiretu, vbFromUnicode))
    
    If OriginByte <= ByteNum Then
        '「Mojiretu」のバイト文字数計算が省略するバイト文字数以下なら
        '省略はしない
        Output = Mojiretu
    Else
    
        Dim RuikeiByteList, BunkaiMojiretu
        RuikeiByteList = 文字列の各文字累計バイト数計算(Mojiretu)
        BunkaiMojiretu = 文字列分解(Mojiretu)
        
        Dim AddMoji$
        AddMoji = "."
        
        Dim I&, N&
        N = Len(Mojiretu)
        
        For I = 1 To N
            If RuikeiByteList(I) < ByteNum Then
                Output = Output & BunkaiMojiretu(I)
                
            ElseIf RuikeiByteList(I) = ByteNum Then
                If LenB(StrConv(BunkaiMojiretu(I), vbFromUnicode)) = 1 Then
                    '例：Mojiretu = "魑魅魍魎" , ByteNum = 6 ,RuikeiByteList(3) = 6
                    'Output = "魑魅.."
                    Output = Output & AddMoji
                Else
                    '例：Mojiretu = "魑魅XX魎" , ByteNum = 6 ,RuikeiByteList(4) = 6
                    'Output = "魑魅X."
                    Output = Output & AddMoji & AddMoji
                End If
                
                Exit For
                
            ElseIf RuikeiByteList(I) > ByteNum Then
                '例：Mojiretu = "魑魅魍魎" , ByteNum = 7 ,RuikeiByteList(4) = 8
                'Output = "魑魅魍."
                Output = Output & AddMoji
                Exit For
            End If
        Next I
        
    End If
        
    文字列を指定バイト数文字数に省略 = Output

    
End Function

Function 文字列の各文字累計バイト数計算(Mojiretu$)
    '20201023追加

    '文字列を1文字ずつに分解して、各文字のバイト文字長を計算し、
    'その累計値を計算する。
    '例：Mojiretu="新型EKワゴン"
    '出力→Output = (2,4,5,6,7,10,12)
    
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
    
    文字列の各文字累計バイト数計算 = Output
    
End Function

Function 文字列分解(Mojiretu$)
    '20201023追加

    '文字列を1文字ずつ分解して配列に格納
    Dim I&, N&
    Dim Output
    
    N = Len(Mojiretu)
    ReDim Output(1 To N)
    For I = 1 To N
        Output(I) = Mid(Mojiretu, I, 1)
    Next I
    
    文字列分解 = Output
    
End Function

