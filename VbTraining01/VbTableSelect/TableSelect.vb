Module TableSelect

    Sub Main()

        ' 入力画面クラスのインスタンス作成
        Dim inp As selDisplayProc = New selDisplayProc()
        Dim tbl As selTableProc = New selTableProc()
        Dim chk As selCheckProc = New selCheckProc()
        'Dim emp As EmployeeModel = New EmployeeModel()

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
            Else
                Console.WriteLine("入力されたNoは登録されていません。")
            End If

        End If

    End Sub

End Module
