Imports VbTableSelect

Public Class updCheckProc

    Public ChkResult As String

    '====================
    ' コンストラクタ
    '====================   
    Public Sub New()
        Me.ChkResult = Nothing
    End Sub

    Public Sub CheckNo(ByVal no As String, ByRef list As List(Of EmployeeModel))

        Dim tbl As updTableProc = New updTableProc()

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

        ' 社員番号存在チェック
        tbl.SelectTable(no, list)
        ' 社員番号が存在しない場合、エラー（処理を抜ける）
        If list.Count = 0 Then
            Console.WriteLine("入力されたNoは登録されていません。")
            Me.ChkResult = "ERR"
            Exit Sub
        End If

        ' チェック結果：正常
        Me.ChkResult = "OK"

    End Sub

    '====================
    ' 項目入力値チェック処理
    '====================
    Public Sub CheckData(ByVal inp As updDisplayProc)

        Dim tbl As updTableProc = New updTableProc()

        Dim SJIS As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")

        Me.ChkResult = Nothing

        ' 名前
        If String.IsNullOrEmpty(inp.inName) Then
            ' 入力値が空の場合エラー
            Console.WriteLine("警告：名前が未入力です。")
        ElseIf SJIS.GetByteCount(inp.inName) > 20 Then
            ' サイズオーバー
            Console.WriteLine("名前のサイズがオーバーしています。")
            Me.ChkResult = "ERR"
            Exit Sub
        End If

        ' 住所
        If String.IsNullOrEmpty(inp.inAddress) Then
            ' 入力値が空の場合エラー
            Console.WriteLine("警告：住所が未入力です。")
        ElseIf SJIS.GetByteCount(inp.inAddress) > 50 Then
            ' サイズオーバー
            Console.WriteLine("住所のサイズがオーバーしています。")
            Me.ChkResult = "ERR"
            Exit Sub
        End If

        ' 電話番号
        If String.IsNullOrEmpty(inp.inTel) Then
            ' 入力値が空の場合エラー
            Console.WriteLine("警告：電話番号が未入力です。")
        ElseIf Not IsNumeric(inp.inTel) Then
            ' 数値以外の場合エラー
            Console.WriteLine("電話番号は数値を入力してください。")
            Me.ChkResult = "ERR"
            Exit Sub
        ElseIf SJIS.GetByteCount(inp.inTel) > 11 Then
            ' サイズオーバー
            Console.WriteLine("電話番号のサイズがオーバーしています。")
            Me.ChkResult = "ERR"
            Exit Sub
        End If

        ' 性別
        If String.IsNullOrEmpty(inp.insexuality) Then
            ' 入力値が空の場合、その他
            inp.dbsex = Int(tbl.SexCode.Other).ToString
        ElseIf SJIS.GetByteCount(inp.insexuality) > 1 Then
            ' 入力誤り
            Console.WriteLine("性別値が範囲外です。")
            Me.ChkResult = "ERR"
            Exit Sub
        ElseIf inp.insexuality.Equals("M") Or inp.insexuality.Equals("m") Then
            inp.dbsex = Int(tbl.SexCode.Man).ToString

        ElseIf inp.insexuality.Equals("W") Or inp.insexuality.Equals("w") Then
            inp.dbsex = Int(tbl.SexCode.Woman).ToString
        Else
            inp.dbsex = Int(tbl.SexCode.Other).ToString
        End If

        ' チェック結果：正常
        Me.ChkResult = "OK"

    End Sub

End Class
