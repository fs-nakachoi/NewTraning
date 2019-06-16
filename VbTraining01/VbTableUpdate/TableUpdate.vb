Imports VbTableSelect

Module TableUpdate

    Sub Main()

        ' 変数宣言
        Dim ansYN As String     ' Y/N

        'アセンブリ名を取得
        Dim assemblyName As String = My.Application.Info.AssemblyName

        ' 入力画面クラスのインスタンス作成
        Dim inp As updDisplayProc = New updDisplayProc()
        Dim tbl As updTableProc = New updTableProc()
        Dim chk As updCheckProc = New updCheckProc()

        Dim EmpList As New List(Of EmployeeModel)()

        ' 入力処理呼び出し
        inp.InputNo()

        ' 入力チェック（社員番号チェック）
        chk.CheckNo(inp.inNo, EmpList)

        If chk.ChkResult.Equals("OK") Then

            ' 取得データを標準出力
            inp.DispEmployee(inp.inNo, EmpList)

            ' 更新値の入力処理
            inp.InputData()

            ' 更新値のチェック処理
            chk.CheckData(inp)

            If chk.ChkResult.Equals("OK") Then

                Console.Write("このデータを更新します。宜しいですか？(y/n):")
                ansYN = Console.ReadLine()

                If ansYN.Equals("Y") Or ansYN.Equals("y") Then
                    ' データを更新
                    tbl.UpdateTable(inp, assemblyName)
                Else
                    Console.WriteLine("処理がキャンセルされました。")
                End If

            End If

        End If

    End Sub

End Module
