Module modUtility
    ''' <summary>
    ''' 清空指定控制項內其他控制項
    ''' </summary>
    ''' <param name="ctrls">控制項的集合</param>
    Public Sub ClearControls(ctrls As Control, Optional exception As List(Of String) = Nothing)
        For Each ctrl As Control In ctrls.Controls
            If exception IsNot Nothing Then
                If exception.Contains(ctrls.Name) Or exception.Contains(ctrls.Text) Then Continue For
            End If
            If TypeOf ctrl Is GroupBox Then
                Dim grp = CType(ctrl, GroupBox)
                ClearControls(grp)
            ElseIf TypeOf ctrl Is TabControl Then
                For Each tp As TabPage In CType(ctrl, TabControl).Controls
                    ClearControls(ctrls)
                Next
            End If
            If TypeOf ctrl Is TextBox Then
                ctrl.Text = ""
            ElseIf TypeOf ctrl Is CheckBox Then
                CType(ctrl, CheckBox).Checked = False
            ElseIf TypeOf ctrl Is RadioButton Then
                CType(ctrl, RadioButton).Checked = False
            ElseIf TypeOf ctrl Is ComboBox Then
                CType(ctrl, ComboBox).SelectedIndex = -1
            End If
        Next
    End Sub

    ''' <summary>
    ''' 將取得的資料傳至各控制項(控制項的Tag必須寫上表格欄位名稱)
    ''' </summary>
    ''' <param name="ctrls">父容器</param>
    ''' <param name="row"></param>
    Public Sub GetDataToControls(ctrls As Control, row As Object)
        For Each ctrl In ctrls.Controls.Cast(Of Control).Where(Function(c) Not String.IsNullOrEmpty(c.Tag))
            Dim value = GetCellData(row, ctrl.Tag.ToString)
            Select Case ctrl.GetType.Name
                Case "TextBox"
                    ctrl.Text = value
                Case "DateTimePicker"
                    Dim dtp As DateTimePicker = ctrl
                    dtp.Value = value
                Case "ComboBox"
                    Dim cmb As ComboBox = ctrl
                    cmb.SelectedIndex = cmb.FindStringExact(value)
                Case "GroupBox"
                    Dim grp As GroupBox = ctrl
                    For Each c In grp.Controls
                        If TypeOf c Is CheckBox Then
                            Dim chk As CheckBox = c
                            Dim b As Boolean
                            If Boolean.TryParse(value, b) Then
                                chk.Checked = value
                            Else
                                chk.Checked = value.Contains(chk.Text)
                            End If
                        ElseIf TypeOf c Is RadioButton Then
                            Dim rdo As RadioButton = c
                            rdo.Checked = rdo.Text = value
                        End If
                    Next
                    GetDataToControls(ctrl, row)
                Case "CheckBox"
                    Dim chk As CheckBox = ctrl
                    If Boolean.Parse(value) Then
                        chk.Checked = value
                    Else
                        chk.Checked = value.Contains(chk.Text)
                    End If
                Case Else
            End Select
        Next
    End Sub

    ''' <summary>
    ''' 取得儲存格的內容
    ''' </summary>
    ''' <param name="row">DataRow、DataGridViewRow</param>
    ''' <param name="colName"></param>
    ''' <returns></returns>
    Private Function GetCellData(row As Object, colName As String) As String
        Select Case row.GetType.Name
            Case "DataRow"
                Dim r As DataRow = row
                Return r(colName).ToString
            Case "DataGridViewRow"
                Dim r As DataGridViewRow = row
                Return r.Cells(colName).Value.ToString
            Case Else
                Return ""
        End Select
    End Function

    ''' <summary>
    ''' 檢查必填欄位
    ''' </summary>
    ''' <param name="required">填入key:欄位名稱 value:控制項</param>
    ''' <returns></returns>
    Public Function CheckRequiredCol(required As Dictionary(Of String, Object)) As Boolean
        For Each kvp In required
            If String.IsNullOrWhiteSpace(kvp.Value.Text) Then
                MsgBox(kvp.Key + " 不能空白")
                kvp.Value.Focus()
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' 檢查是否重複新增
    ''' </summary>
    ''' <param name="selectFrom">SQL前半段</param>
    ''' <param name="list">條件,輸入控制項會自動取得Tag(欄位名稱),Text(值)</param>
    ''' <param name="dgv"></param>
    ''' <returns></returns>
    Public Function CheckDuplication(selectFrom As String, list As List(Of Object), dgv As DataGridView) As Boolean
        Dim sql = selectFrom + $" WHERE {String.Join(" AND ", list.Select(Function(x) $"{x.tag} = '{x.text}'"))}"
        If SelectTable(sql).Rows.Count > 0 Then
            MsgBox("重複資料")
            '列出重複的資料
            GetDataToDgv(sql, dgv)
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' 將資料放到DataGridView
    ''' </summary>
    ''' <param name="sql"></param>
    ''' <param name="dgv"></param>
    Public Sub GetDataToDgv(sql As String, dgv As DataGridView)
        With dgv
            .DataSource = SelectTable(sql)
            Dim lstTableNames = GetTableNamesFromQuery(sql)
            '條件式
            Dim conditions = String.Join(" OR ", lstTableNames.Select(Function(x) $"Table_name = '{x}'"))
            '用table欄位的備註將dgv的欄位改名
            Dim tableCol = SelectTable($"SELECT COLUMN_NAME, COLUMN_COMMENT FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'gas_accounting_system' AND {conditions}")
            For Each col As DataGridViewColumn In .Columns
                Dim row = tableCol.AsEnumerable().FirstOrDefault(Function(x) x("COLUMN_NAME").ToString() = col.Name)
                If row IsNot Nothing Then
                    col.HeaderText = row("COLUMN_COMMENT").ToString()
                End If
            Next
            .AutoResizeColumnHeadersHeight()
        End With
    End Sub

    ''' <summary>
    ''' 設定DataGridView的樣式屬性
    ''' </summary>
    ''' <param name="ctrl">父容器</param>
    Public Sub SetDataGridViewStyle(ctrl As Control)
        For Each dgv In GetControlInParent(Of DataGridView)(ctrl)
            With dgv
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                .ColumnHeadersDefaultCellStyle.Font = New Font("標楷體", 12, FontStyle.Bold)
                .DefaultCellStyle.Font = New Font("標楷體", 12, FontStyle.Bold)
                .AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(224, 224, 224)
                .EnableHeadersVisualStyles = False
                .ColumnHeadersDefaultCellStyle.BackColor = Color.MediumTurquoise
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .ReadOnly = True
                .Anchor = AnchorStyles.Bottom And AnchorStyles.Left
            End With
        Next
    End Sub

    ''' <summary>
    ''' 取得指定控制項內所有的目標控制項
    ''' </summary>
    ''' <typeparam name="T">目標控制項</typeparam>
    ''' <param name="parent">父控制項</param>
    ''' <returns></returns>
    Public Function GetControlInParent(Of T As Control)(parent As Control) As List(Of T)
        Dim lst As New List(Of T)
        If parent.Controls.Count > 0 Then
            For Each ctrl In parent.Controls
                If TypeOf ctrl Is T Then lst.Add(ctrl)
                lst.AddRange(GetControlInParent(Of T)(ctrl))
            Next
        End If
        Return lst
    End Function

    ''' <summary>
    ''' 檢查TextBox裡是否為正整數
    ''' </summary>
    ''' <param name="txt"></param>
    ''' <returns></returns>
    Public Function CheckPositiveInteger(txt As TextBox) As Boolean
        If Not IsNumeric(txt.Text) Then
            MsgBox(txt.Tag + " 不為數字!")
            txt.Focus()
            Return False
        End If
        If Val(txt.Text) < 0 Then
            MsgBox(txt.Tag + " 不能為負數!")
            txt.Focus()
            Return False
        End If
        Return True
    End Function
End Module
