Public Class selCheckProc

    Public ChkResult As String

    '====================
    ' コンストラクタ
    '====================   
    Public Sub New()
        Me.ChkResult = Nothing
    End Sub

    Public Sub CheckNo(ByVal no As String)

        If Not String.IsNullOrEmpty(no) And Not IsNumeric(no) Then
            ' 空値以外且つ数値以外の場合エラー
            Console.WriteLine("Noは数値を入力してください。")
            Me.ChkResult = "ERR"
            Exit Sub
        End If

        ' チェック結果：正常
        Me.ChkResult = "OK"

    End Sub

End Class
