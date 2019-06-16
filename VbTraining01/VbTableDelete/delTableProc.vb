Imports MySql.Data.MySqlClient
Imports VbTableSelect

Public Class delTableProc

    Private DbConnVal As String

    '====================
    ' コンストラクタ
    '====================   
    Public Sub New()

        ' 接続文字列の設定
        Me.DbConnVal = "Database=fsdb01;Data Source=localhost;User Id=fsadmin;Password=fspas2019; sqlservermode=True;"

    End Sub

    '====================
    ' データ取得処理
    '====================   
    Public Sub SelectTable(ByVal no As String, ByRef list As List(Of EmployeeModel))

        Using conn As New MySqlConnection(Me.DbConnVal)

            Try
                Dim trn As MySqlTransaction = Nothing

                ' データベース・オープン
                conn.Open()

                ' Insert SQL文定義
                Dim strSQL As String
                strSQL = "SELECT NO, NAME, ADDRESS, TEL, SEXUALITY FROM FS_EMPLOYEE_T WHERE no = @no"

                ' SQLコマンド生成
                Dim cmd As MySqlCommand = New MySqlCommand(strSQL, conn)

                ' 項目値パラメータの設定
                Dim p1 As MySqlParameter = New MySqlParameter("@no", Integer.Parse(no))
                cmd.Parameters.Add(p1)

                Dim da As MySqlDataAdapter = New MySqlDataAdapter(cmd)
                Dim dt As DataTable = New DataTable()
                da.Fill(dt)

                For Each row As DataRow In dt.Rows
                    Dim data As EmployeeModel = New EmployeeModel()
                    data.no = row("no")
                    data.name = row("name")
                    data.address = row("address")
                    data.tel = row("tel")
                    data.sex = row("sexuality")
                    list.Add(data)
                Next

                ' 例外処理
            Catch ex As Exception

                Console.WriteLine("システムエラー：取得に失敗しました。")

                ' 終了処理
            Finally

                ' データベース・クローズ
                If Not conn.State = ConnectionState.Closed Then
                    conn.Close()
                End If

            End Try

        End Using

    End Sub

    '====================
    ' データ削除処理
    '====================   
    Public Sub DeleteTable(ByVal no As String)

        Using conn As New MySqlConnection(Me.DbConnVal)

            Dim trn As MySqlTransaction = Nothing

            ' データベース・オープン
            conn.Open()

            ' Insert SQL文定義
            Dim strSQL As String
            strSQL = "DELETE FROM FS_EMPLOYEE_T WHERE no = @no "
            ' SQLコマンド生成
            Dim cmd As MySqlCommand = New MySqlCommand(strSQL, conn)

            ' トランザクション開始
            trn = cmd.Connection.BeginTransaction()

            Try

                ' 項目値パラメータの設定
                Dim p1 As MySqlParameter = New MySqlParameter("@no", Integer.Parse(no))
                cmd.Parameters.Add(p1)

                ' SQL実行
                cmd.ExecuteNonQuery()

                ' SQL実行が成功した場合、コミット
                trn.Commit()

                ' 削除成功
                Console.WriteLine("データを削除しました。")

                ' 例外処理
            Catch ex As Exception

                ' SQL実行が失敗した場合、ロールバック
                If trn IsNot Nothing Then
                    trn.Rollback()
                End If
                Console.WriteLine("システムエラー：削除に失敗しました。")

                ' 終了処理
            Finally

                ' データベース・クローズ
                If Not conn.State = ConnectionState.Closed Then
                    conn.Close()
                End If

            End Try

        End Using

    End Sub

End Class
