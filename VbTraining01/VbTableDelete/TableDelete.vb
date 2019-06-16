Imports VbTableSelect

Module TableDelete

    Sub Main()

        ' 変数宣言
        Dim ansYN As String     ' Y/N

        'アセンブリ名を取得
        Dim assemblyName As String = My.Application.Info.AssemblyName

        ' 入力画面クラスのインスタンス作成
        Dim inp As delDisplayProc = New delDisplayProc()
        Dim tbl As delTableProc = New delTableProc()
        Dim chk As delCheckProc = New delCheckProc()

        Dim EmpList As New List(Of EmployeeModel)()

        ' 入力処理呼び出し
        inp.InputNo()

        ' 入力チェック（社員番号チェック）
        chk.CheckNo(inp.inNo)

        If chk.ChkResult.Equals("OK") Then

            ' テーブル取得処理
            tbl.SelectTable(inp.inNo, EmpList)

            If EmpList.Count > 0 Then
                ' 取得データを標準出力
                inp.DispEmployee(inp.inNo, EmpList)
                Console.Write("このデータを削除します。宜しいですか？(y/n):")

                ansYN = Console.ReadLine()

                If ansYN.Equals("Y") Or ansYN.Equals("y") Then
                    ' データを削除
                    tbl.DeleteTable(inp.inNo)
                Else
                    Console.WriteLine("処理がキャンセルされました。")
                End If

            Else
                Console.WriteLine("入力されたNoは登録されていません。")
            End If

        End If

    End Sub

End Module
