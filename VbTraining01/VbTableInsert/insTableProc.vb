Imports MySql.data.MySqlClient

Public Class insTableProc

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
    ' データ登録処理
    '====================   
    Public Sub InsertTable(ByVal dat As insDisplayProc, ByVal pgmid As String)

        Using conn As New MySqlConnection(Me.DbConnVal)

            Dim trn As MySqlTransaction = Nothing

            ' データベース・オープン
            conn.Open()

            ' Insert SQL文定義
            Dim strSQL As String
            strSQL = "
            INSERT INTO FS_EMPLOYEE_T 
              (NO,NAME,ADDRESS,TEL,SEXUALITY,CREATE_PGM,CREATE_DATE,UPDATE_PGM,UPDATE_DATE)
              VALUES(@no,@name,@address,@tel,@sex,@cpgm,CURRENT_TIMESTAMP,@upgm,CURRENT_TIMESTAMP);
            "
            ' SQLコマンド生成
            Dim cmd As MySqlCommand = New MySqlCommand(strSQL, conn)

            ' トランザクション開始
            trn = cmd.Connection.BeginTransaction()

            Try

                ' 項目値パラメータの設定
                Dim p1 As MySqlParameter = New MySqlParameter("@no", Integer.Parse(dat.inNo))
                cmd.Parameters.Add(p1)
                Dim p2 As MySqlParameter = New MySqlParameter("@name", dat.inName)
                cmd.Parameters.Add(p2)
                Dim p3 As MySqlParameter = New MySqlParameter("@address", dat.inAddress)
                cmd.Parameters.Add(p3)
                Dim p4 As MySqlParameter = New MySqlParameter("@tel", dat.inTel)
                cmd.Parameters.Add(p4)
                Dim p5 As MySqlParameter = New MySqlParameter("@sex", dat.dbsex)
                cmd.Parameters.Add(p5)
                Dim p6 As MySqlParameter = New MySqlParameter("@cpgm", pgmid)
                cmd.Parameters.Add(p6)
                Dim p7 As MySqlParameter = New MySqlParameter("@upgm", pgmid)
                cmd.Parameters.Add(p7)

                ' SQL実行
                cmd.ExecuteNonQuery()

                ' SQL実行が成功した場合、コミット
                trn.Commit()

                ' 登録成功
                Console.WriteLine("登録に成功しました。")

                ' 例外処理
            Catch ex As Exception

                ' SQL実行が失敗した場合、ロールバック
                If trn IsNot Nothing Then
                    trn.Rollback()
                End If
                Console.WriteLine("システムエラー：登録に失敗しました。")

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
    ' 社員番号存在チェック処理
    '====================   
    Public Sub SelectTable(ByVal no As String, ByRef cnt As Integer)

        cnt = -1

        Using conn As New MySqlConnection(Me.DbConnVal)

            Try

                ' データベース・オープン
                conn.Open()

                ' Insert SQL文定義
                Dim strSQL As String
                strSQL = "SELECT COUNT(*) FROM FS_EMPLOYEE_T WHERE no = @no;"
                ' SQLコマンド生成
                Dim cmd As MySqlCommand = New MySqlCommand(strSQL, conn)

                ' 項目値パラメータの設定
                Dim p1 As MySqlParameter = New MySqlParameter("@no", Integer.Parse(no))
                cmd.Parameters.Add(p1)

                ' SQL実行（件数取得）
                cnt = cmd.ExecuteScalar()
                ' 入力された社員番号が存在した場合、エラー
                If cnt > 0 Then
                    Console.WriteLine("入力されたNoは既に使用されているため、登録できません。")
                End If

                ' 例外処理
            Catch ex As Exception

                Console.WriteLine("システムエラー：テーブル取得に失敗しました。(SelectTable)")

                '終了処理
            Finally

                ' データベース・クローズ
                If Not conn.State = ConnectionState.Closed Then
                    conn.Close()
                End If

            End Try

        End Using

    End Sub

End Class