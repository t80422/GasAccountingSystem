Imports System.Configuration
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports MySql.Data.MySqlClient
'MySQL相關
Module modMySQL
    Friend conn As MySqlConnection
    Private title = "MySQL"

    '初始化
    Sub New()
        Dim ip = ConfigurationManager.AppSettings("ServerIP")
        Dim uid = ConfigurationManager.AppSettings("UserID")
        Dim psw = ConfigurationManager.AppSettings("Password")
        Dim db = ConfigurationManager.AppSettings("Database")
        conn = New MySqlConnection($"server={ip};uid={uid};pwd={psw};database={db};charset=utf8")
        '測試連線
        Try
            conn.Open()
        Catch ex As Exception
            '使用3306Port 如果開不起來就是mysql卡住 要到工作管理員結束工作後重開
            MsgBox("資料庫連線失敗", Title:=title)
        End Try
        conn.Close()
    End Sub

    ''' <summary>
    ''' 查詢資料表
    ''' </summary>
    ''' <returns></returns>
    Friend Function SelectTable(sSQL As String) As DataTable
        Dim dt As New DataTable()
        Try
            conn.Open()
            Using cmd As New MySqlCommand(sSQL, conn)
                Dim adapter As New MySqlDataAdapter(cmd)
                adapter.Fill(dt)
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        conn.Close()
        Return dt
    End Function

    ''' <summary>
    ''' 新增資料至資料表
    ''' </summary>
    ''' <param name="sTable"></param>
    ''' <param name="dicData"></param>
    ''' <returns></returns>
    Public Function InserTable(sTable As String, dicData As Dictionary(Of String, String)) As Boolean
        Dim result As Boolean
        Dim cmd As New MySqlCommand($"INSERT INTO {sTable} ({String.Join(",", dicData.Keys)}) VALUES ({String.Join(",", dicData.Keys.Select(Function(key) $"@{key}"))})", conn)
        Try
            conn.Open()
            For Each kvp In dicData
                cmd.Parameters.AddWithValue($"@{kvp.Key}", Trim(kvp.Value))
            Next
            If cmd.ExecuteNonQuery() > 0 Then result = True
        Catch ex As Exception
            MsgBox(ex.Message, Title:=title)
        End Try
        conn.Close()
        Return result
    End Function

    ''' <summary>
    ''' 更新表格
    ''' </summary>
    ''' <param name="table">表格名稱</param>
    ''' <param name="dicFields">更新對象集合</param>
    ''' <param name="condition">Where</param>
    Public Function UpdateTable(table As String, dicFields As Dictionary(Of String, String), condition As String) As Boolean
        Dim result As Boolean = False

        Try
            conn.Open()
            Dim sql = $"UPDATE {table} SET "
            Dim lst As New List(Of String)

            For Each kvp In dicFields
                lst.Add($"{kvp.Key} = @{kvp.Key}")
            Next

            sql += String.Join(",", lst) + $" WHERE {condition}"
            Dim cmd As New MySqlCommand(sql, conn)

            For Each kvp In dicFields
                cmd.Parameters.AddWithValue($"@{kvp.Key}", Trim(kvp.Value))
            Next

            If cmd.ExecuteNonQuery() > 0 Then
                result = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, Title:=MethodBase.GetCurrentMethod.Name)
        Finally
            conn.Close()
        End Try

        Return result
    End Function

    ''' <summary>
    ''' MySQL Delete
    ''' </summary>
    ''' <param name="sTable">資料表</param>
    ''' <param name="sWhere">條件</param>
    ''' <returns></returns>
    Public Function DeleteData(sTable As String, sWhere As String) As Boolean
        Dim rowsAffected As Integer
        Dim cmd As New MySqlCommand($"DELETE FROM {sTable} WHERE {sWhere}", conn)
        Try
            conn.Open()
            rowsAffected = cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message, Title:="警告")
        End Try
        conn.Close()
        Return rowsAffected > 0
    End Function

    ''' <summary>
    ''' 取得SQL語句中的表格名稱
    ''' </summary>
    ''' <param name="query"></param>
    ''' <returns></returns>
    Public Function GetTableNamesFromQuery(query As String) As List(Of String)
        Dim tableNames As New List(Of String)

        ' 使用正則表達式搜尋 FROM 和 JOIN 子句中的表名
        Dim regex As New Regex("(?:FROM|JOIN)\s+(\w+)", RegexOptions.IgnoreCase)
        Dim matches As MatchCollection = regex.Matches(query)

        ' 迭代匹配的結果，並將表名加入列表
        For Each match As Match In matches
            Dim tableName As String = match.Groups(1).Value
            tableNames.Add(tableName)
        Next

        Return tableNames
    End Function

    ''' <summary>
    ''' 檢查是否重複新增
    ''' </summary>
    ''' <param name="selectFrom">SQL前半段</param>
    ''' <param name="dic">條件,key:欄位 value:值</param>
    ''' <param name="dgv">欲顯示的DataGridView</param>
    ''' <returns></returns>
    Public Function CheckDuplication(selectFrom As String, dic As Dictionary(Of String, String), Optional dgv As DataGridView = Nothing) As Boolean
        '修正參數List>Dictionary,如果遇到DateTimePicker就會遇到可能取的值不是想要的
        Dim lst As List(Of String) = dic.Select(Function(kvp) $"{kvp.Key} = '{kvp.Value}'").ToList
        Dim sql = selectFrom + $" WHERE {String.Join(" AND ", lst)}"
        Dim dt = SelectTable(sql)
        If dt.Rows.Count > 0 Then
            MsgBox("重複資料")
            If dgv IsNot Nothing Then dgv.DataSource = dt
            Return False
        End If
        Return True
    End Function
End Module
