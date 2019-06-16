Public Class delCheckProc

    Public ChkResult As String

    '====================
    ' コンストラクタ
    '====================   
    Public Sub New()
        Me.ChkResult = Nothing
    End Sub

    Public Sub CheckNo(ByVal no As String)

        If String.IsNullOrEmpty(no) Then
            ' 空値の場合エラー
            Console.WriteLine("Noが入力されていません。")
            Me.ChkResult = "ERR"
            Exit Sub
        End If

        If Not IsNumeric(no) Then
            ' 数値以外の場合エラー
            Console.WriteLine("Noは数値を入力してください。")
            Me.ChkResult = "ERR"
            Exit Sub
        End If

        ' チェック結果：正常
        Me.ChkResult = "OK"

    End Sub

End Class
