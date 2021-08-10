# VBA-配列をイミディエイトウィンドウに見やすく表示する
License: The MIT license

Copyright (c) 2021 YujiFukami

開発テスト環境 Excel: Microsoft® Excel® 2019 32bit 

開発テスト環境 OS: Windows 10 Pro


# 使い方
モジュール「ModImmediate.bas」をVBEにインポートするか、コードをそのままVBEにコピー&ペーストする。

イミディエイトウィンドウに表示したい配列（1次元または2次元配列）をプロシージャ「DPH」で引数としてCallするか、
イミディエイトウィンドウ上で「DPH(表示したい配列)」と入力する。

モジュール「ModImmediate.bas」内にテスト実行用のプロシージャ「DPHTest」が入っているので実行して確かめてみるべし

    Sub DPHTest()

        Dim HairetuDummy
        HairetuDummy = Array(Array(1, 2, 3, 4, 5), _
                       Array("A", "AA", "AAA", "AAAA", "AAAAA"), _
                       Array("あ", "あああ", "||||||", "ううう", "あ"))
        HairetuDummy = Application.Transpose(Application.Transpose(HairetuDummy))

        Call DPH(HairetuDummy, 3, "テスト1")

        Call DPH(HairetuDummy, , "テスト2")

    End Sub

