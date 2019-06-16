Imports VbTableSelect

Public Class updDisplayProc

    ' 変数宣言
    Public inNo As String           ' 社員番号
    Public inName As String         ' 氏名
    Public inAddress As String      ' 住所
    Public inTel As String          ' 電話番号
    Public insexuality As String   ' 性別
    Public dbsex As String        ' 性別（テーブル値）

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
        Console.Write("変更前No:")
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

    End Sub

    '====================
    ' 入力処理
    '====================
    Public Sub InputData()

        '---
        ' 入力項目毎に項目名を標準出力し入力値を受付する。
        '---

        ' 名前
        Console.Write("名前:")
        inName = Console.ReadLine()

        ' 住所
        Console.Write("住所:")
        inAddress = Console.ReadLine()

        ' 電話番号
        Console.Write("電話番号:")
        inTel = Console.ReadLine()

        ' 性別(M:男性,W:女性)
        Console.Write("性別(M:男性,W:女性):")
        insexuality = Console.ReadLine()

    End Sub

End Class
