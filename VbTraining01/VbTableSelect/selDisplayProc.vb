Public Class selDisplayProc

    ' 変数宣言
    Public inNo As String           ' 社員番号

    '====================
    ' コンストラクタ
    '====================
    Public Sub New()

    End Sub

    '====================
    ' 入力処理
    '====================
    Public Sub InputNo()

        ' 社員番号
        Console.Write("No:")
        inNo = Console.ReadLine()

    End Sub

    '====================
    ' データ標準出力処理
    '====================
    Public Sub DispEmployee(ByVal no As String, ByVal list As List(Of EmployeeModel))

        ' 見出し行出力
        Console.WriteLine(" No 名前                 住所                                               電話番号    性別")

        ' 取得したデータリストを全て標準出力する。
        For Each data As EmployeeModel In list
            Dim namLen = System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(data.name)
            Dim addLen = System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(data.address)
            Dim telLen = System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(data.tel)
            Dim sexVal = ""
            Select Case data.sex
                Case 1
                    sexVal = "男性"
                Case 2
                    sexVal = "女性"
                Case Else
                    sexVal = "その他"
            End Select

            Console.WriteLine("{0,3} {1} {2} {3} {4}",
                              data.no,
                              data.name + Space(20 - namLen),
                              data.address + Space(50 - addLen),
                              data.tel + Space(11 - telLen),
                              sexVal)
        Next

        ' 入力された社員番号が空値の場合、件数を標準出力する。
        If String.IsNullOrEmpty(no) Then
            Console.WriteLine("{0}件のデータが登録されています。", list.Count)
        End If

    End Sub

End Class
