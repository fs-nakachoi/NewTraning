Imports MySql.Data.MySqlClient

Public Class selTableProc

    Private DbConnVal As String

    ' 性別列挙型
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
    Public Sub SelectTable(ByVal no As String, ByRef list As List(Of EmployeeModel))

        Using conn As New MySqlConnection(Me.DbConnVal)

            Try
                Dim trn As MySqlTransaction = Nothing

                ' データベース・オープン
                conn.Open()

                ' Insert SQL文定義
                Dim strSQL As String
                strSQL = "SELECT NO, NAME, ADDRESS, TEL, SEXUALITY FROM FS_EMPLOYEE_T "
                If Not String.IsNullOrEmpty(no) Then
                    strSQL = strSQL & "WHERE no = @no "
                End If

                ' SQLコマンド生成
                Dim cmd As MySqlCommand = New MySqlCommand(strSQL, conn)

                ' 項目値パラメータの設定
                If Not String.IsNullOrEmpty(no) Then
                    Dim p1 As MySqlParameter = New MySqlParameter("@no", Integer.Parse(no))
                    cmd.Parameters.Add(p1)
                End If

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

End Class