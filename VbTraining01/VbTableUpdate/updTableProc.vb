Imports MySql.Data.MySqlClient
Imports VbTableSelect

Public Class updTableProc

    Private DbConnVal As String

    Public Enum SexCode As Integer
        Man = 1
        Woman = 2
        Other = 9
    End Enum

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
    ' データ更新処理
    '====================   
    Public Sub UpdateTable(ByVal dat As updDisplayProc, ByVal pgmid As String)

        Using conn As New MySqlConnection(Me.DbConnVal)

            Dim trn As MySqlTransaction = Nothing

            ' データベース・オープン
            conn.Open()

            ' Update SQL文定義
            Dim strSQL As String
            strSQL = "
UPDATE FS_EMPLOYEE_T SET
  NAME = @name,
  ADDRESS = @address,
  TEL = @tel,
  SEXUALITY = @sex,
UPDATE_PGM = @upgm,
UPDATE_DATE = CURRENT_TIMESTAMP
WHERE no = @no"
            ' SQLコマンド生成
            Dim cmd As MySqlCommand = New MySqlCommand(strSQL, conn)

            ' トランザクション開始
            trn = cmd.Connection.BeginTransaction()

            Try

                ' 項目値パラメータの設定
                Dim p1 As MySqlParameter = New MySqlParameter("@name", dat.inName)
                cmd.Parameters.Add(p1)
                Dim p2 As MySqlParameter = New MySqlParameter("@address", dat.inAddress)
                cmd.Parameters.Add(p2)
                Dim p3 As MySqlParameter = New MySqlParameter("@tel", dat.inTel)
                cmd.Parameters.Add(p3)
                Dim p4 As MySqlParameter = New MySqlParameter("@sex", dat.dbsex)
                cmd.Parameters.Add(p4)
                Dim p5 As MySqlParameter = New MySqlParameter("@upgm", pgmid)
                cmd.Parameters.Add(p5)
                Dim p6 As MySqlParameter = New MySqlParameter("@no", Integer.Parse(dat.inNo))
                cmd.Parameters.Add(p6)

                ' SQL実行
                cmd.ExecuteNonQuery()

                ' SQL実行が成功した場合、コミット
                trn.Commit()

                ' 登録成功
                Console.WriteLine("データを更新しました。")

                ' 例外処理
            Catch ex As Exception

                ' SQL実行が失敗した場合、ロールバック
                If trn IsNot Nothing Then
                    trn.Rollback()
                End If
                Console.WriteLine("システムエラー：更新に失敗しました。")

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
