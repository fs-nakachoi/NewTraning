Module TableInsert

    Sub Main()

        ' 変数宣言
        Dim ansYN As String     ' Y/N

        'アセンブリ名を取得
        Dim assemblyName As String = My.Application.Info.AssemblyName

        ' 入力画面クラスのインスタンス作成
        Dim inp As insDisplayProc = New insDisplayProc()
        Dim tbl As insTableProc = New insTableProc()
        Dim chk As insCheckProc = New insCheckProc()

        ' 入力処理呼び出し
        inp.InputData()

        ' 入力情報の表示
        'Console.WriteLine("{0} {1} {2} {3}", inp.inNo, inp.inName, inp.inAddress, inp.inTel)

        chk.CheckData(inp)
        If chk.ChkResult.Equals("OK") Then

            Console.Write("この入力データを登録します。宜しいですか？(y/n):")
            ansYN = Console.ReadLine()

            If ansYN.Equals("Y") Or ansYN.Equals("y") Then
                ' テーブルへデータ挿入
                tbl.InsertTable(inp, assemblyName)
            Else
                Console.WriteLine("処理がキャンセルされました。")
            End If

        End If

    End Sub

End Module
