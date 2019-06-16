'
' 社員テーブル・モデル定義
'
Public Class EmployeeModel

    Public no As Integer
    Public name As String
    Public address As String
    Public tel As String
    Public sex As String

    '====================
    ' コンストラクタ
    '====================   
    Public Sub New()

        no = Nothing
        name = Nothing
        address = Nothing
        tel = Nothing
        sex = Nothing

    End Sub

End Class
