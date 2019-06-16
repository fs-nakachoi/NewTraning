Public Class insDisplayProc

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
    Public Sub InputData()

        '---
        ' 入力項目毎に項目名を標準出力し入力値を受付する。
        '---

        ' 社員番号
        Console.Write("No:")
        inNo = Console.ReadLine()

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
