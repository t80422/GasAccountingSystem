Public Class Form1
    '自定義TabPage索引標籤、文字顏色
    Private Sub TabControl1_DrawItem(sender As TabControl, e As DrawItemEventArgs) Handles TabControl1.DrawItem
        '檢查當前索引標籤是否為選中狀態
        Dim isSelected As Boolean = (e.State And DrawItemState.Selected) = DrawItemState.Selected
        '繪製索引標籤的背景
        Dim backColor As Color = If(isSelected, Color.LightBlue, Color.WhiteSmoke)
        e.Graphics.FillRectangle(New SolidBrush(backColor), e.Bounds)
        '繪製索引標籤的文字
        Dim text As String = sender.TabPages(e.Index).Text
        Dim textColor As Color = Color.Black
        Dim font As Font = sender.Font
        e.Graphics.DrawString(text, font, New SolidBrush(textColor), e.Bounds.Location)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '自定義索引標籤、文字顏色
        TabControl1.DrawMode = DrawMode.OwnerDrawFixed

        SetDataGridViewStyle(Me)
        Dim dgvs = GetControlInParent(Of DataGridView)(Me)
        dgvs.ForEach(Sub(dgv) AddHandler dgv.ColumnWidthChanged, AddressOf SaveDataGridWidth)
        '初始化各TabPage
        btnCancel_emp.PerformClick()
        btnCancel_perm_Click(btnCancel_perm, e)
        btnCancel_manu_Click(btnCancel_manu, e)
        btnCancel_cus_Click(btnCancel_cus, e)
        ReadDataGridWidth(dgvs)
    End Sub

    Private Sub tpLogOut_Enter(sender As Object, e As EventArgs) Handles tpLogOut.Enter
        If MessageBox.Show("確定要登出嗎??", "登出", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = System.Windows.Forms.DialogResult.OK Then Close()
    End Sub

    '員工管理-取消
    Private Sub btnCancel_emp_Click(sender As Object, e As EventArgs) Handles btnCancel_emp.Click
        Dim tp = sender.Parent
        ClearControls(tp)
        GetDataToDgv("SELECT emp_id,  emp_name,  emp_phone,  emp_address,  emp_ecp,  emp_ect,  emp_acc,  emp_psw, emp_identity_number, DATE_FORMAT(emp_birthday,'%Y/%m/%d') emp_birthday FROM employee", dgvEmployee)
        btnInsert_emp.Enabled = True
        btnModify_emp.Enabled = False
        btnDelete_emp.Enabled = False
        '新刪修後要更新權限表格
        btnCancel_perm.PerformClick()
    End Sub

    '員工管理-新增
    Private Sub btnInsert_emp_Click(sender As Object, e As EventArgs) Handles btnInsert_emp.Click
        Dim dic As Dictionary(Of String, String) = CheckEmployee(sender)
        If dic Is Nothing Then Exit Sub

        Dim dicDup As New Dictionary(Of String, String) From {
            {txtEmpName.Tag, txtEmpName.Text},
            {txtEmpPhone.Tag, txtEmpPhone.Text}
        }
        If Not CheckDuplication($"SELECT * FROM employee", dicDup, dgvEmployee) Then Exit Sub

        If InserTable("employee", dic) Then
            btnCancel_emp.PerformClick()
            btnCancel_perm_Click(btnCancel_perm, e)
            MsgBox("新增成功")
        End If
    End Sub

    '員工管理-dgv點擊
    Private Sub dgvEmployee_CellMouseClick(sender As DataGridView, e As DataGridViewCellMouseEventArgs) Handles dgvEmployee.CellMouseClick
        Dim tp = sender.Parent
        ClearControls(tp)
        GetDataToControls(tp, sender.SelectedRows(0))
        btnInsert_emp.Enabled = False
        btnModify_emp.Enabled = True
        btnDelete_emp.Enabled = True
    End Sub

    '員工管理-修改
    Private Sub btnModify_emp_Click(sender As Object, e As EventArgs) Handles btnModify_emp.Click
        Dim dic As Dictionary(Of String, String) = CheckEmployee(sender)
        If dic Is Nothing Then Exit Sub
        If UpdateTable("employee", dic, $"{txtID_emp.Tag} = '{txtID_emp.Text}'") Then
            btnCancel_emp.PerformClick()
            MsgBox("修改成功")
        End If
    End Sub

    Private Function CheckEmployee(sender As Button) As Dictionary(Of String, String)
        Dim dicReq As New Dictionary(Of String, Object) From {
             {"姓名", txtEmpName},
             {"連絡電話", txtEmpPhone},
             {"帳號", txtAcc},
             {"密碼", txtPsw}
         }
        If Not CheckRequiredCol(dicReq) Then Return Nothing

        '輸入格式防呆
        Dim d As Date
        If Not String.IsNullOrWhiteSpace(txtBirthday.Text) AndAlso Not Date.TryParse(txtBirthday.Text, d) Then
            MsgBox("生日 格式錯誤")
            txtBirthday.Focus()
            Return Nothing
        End If

        Dim tp As TabPage = sender.Parent
        Dim dic As New Dictionary(Of String, String)
        dic = tp.Controls.OfType(Of Control).Where(Function(ctrl) Not String.IsNullOrEmpty(ctrl.Tag) AndAlso ctrl.Tag <> "emp_id" AndAlso Not String.IsNullOrWhiteSpace(ctrl.Text)).
            ToDictionary(Function(ctrl) ctrl.Tag.ToString, Function(ctrl) ctrl.Text)
        Return dic
    End Function

    '員工管理-刪除
    Private Sub btnDelete_emp_Click(sender As Button, e As EventArgs) Handles btnDelete_emp.Click
        Dim tp As TabPage = sender.Parent
        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub
        If DeleteData("employee", $"{txtID_emp.Tag} = '{txtID_emp.Text}'") Then
            btnCancel_emp.PerformClick()
            btnCancel_perm_Click(btnCancel_perm, e)
            MsgBox("刪除成功")
        End If
    End Sub

    '員工管理-查詢
    Private Sub btnQuery_cus_Click(sender As Object, e As EventArgs) Handles btnQuery_emp.Click
        Cursor = Cursors.WaitCursor
        GetDataToDgv($"SELECT * FROM employee WHERE emp_name LIKE '%{txtQuery_emp.Text}%' OR emp_phone LIKE '%{txtQuery_emp.Text}%'", dgvEmployee)
        dgvEmployee.Columns("emp_perm").Visible = False
        Cursor = Cursors.Default
    End Sub

    '搜尋欄位按下"Enter"即可搜尋
    Private Sub txtQuery_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtQuery_emp.KeyPress, txtQuery_perm.KeyPress, txtQuery_manu.KeyPress, txtQuery_cus.KeyPress
        If e.KeyChar = vbCr Then
            Dim btn As Button = CType(sender, TextBox).Parent.Controls.OfType(Of Button).FirstOrDefault(Function(x) x.Text = "查詢")
            btn.PerformClick()
        End If
    End Sub

    '權限管理-取消
    Private Sub btnCancel_perm_Click(sender As Object, e As EventArgs) Handles btnCancel_perm.Click
        Dim tp = sender.Parent
        ClearControls(tp)
        GetDataToDgv("SELECT emp_id, emp_name, emp_acc, emp_perm FROM employee", dgvPermissions)
        dgvPermissions.Columns("emp_perm").Visible = False
        btnModify_perm.Enabled = False
    End Sub

    '權限管理-dgv點擊
    Private Sub dgvPermissions_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvPermissions.CellMouseClick
        Dim tp As TabPage = sender.Parent
        ClearControls(tp)
        Dim row As DataGridViewRow = sender.SelectedRows(0)
        GetDataToControls(tp, row)
        Dim perm = row.Cells("emp_perm").Value
        If IsDBNull(perm) Then Exit Sub
        Dim empPermList = Split(perm, ",")
        tp.Controls.OfType(Of CheckBox).ToList().ForEach(Sub(chk)
                                                             chk.Checked = empPermList.Contains(chk.Text)
                                                         End Sub)
        'For i As Integer = 0 To empPermList.Length - 1
        '    Dim value = empPermList(i).Trim()
        '    Dim checkBox = DirectCast(flpPermissions.Controls(i), CheckBox)
        '    checkBox.Checked = (value = "1")
        'Next
        btnModify_perm.Enabled = True
    End Sub

    '權限管理-修改
    Private Sub btnModify_perm_Click(sender As Object, e As EventArgs) Handles btnModify_perm.Click
        Dim perms = String.Join(",", tpPermissions.Controls.OfType(Of CheckBox).Where(Function(chk) chk.Checked).Select(Function(chk) chk.Text))

        Dim dic As New Dictionary(Of String, String) From {
            {"emp_perm", perms}
        }
        If UpdateTable("employee", dic, $"{txtID_perm.Tag} = '{txtID_perm.Text}'") Then
            btnCancel_perm.PerformClick()
            MsgBox("修改成功")
        End If
    End Sub

    '權限管理-查詢
    Private Sub btnQuery_perm_Click(sender As Object, e As EventArgs) Handles btnQuery_perm.Click
        Cursor = Cursors.WaitCursor
        GetDataToDgv($"SELECT emp_id, emp_name, emp_acc, emp_perm FROM employee WHERE emp_name LIKE '%{txtQuery_perm.Text}%' OR emp_acc LIKE '%{txtQuery_perm.Text}%'", dgvPermissions)
        dgvPermissions.Columns("emp_perm").Visible = False
        Cursor = Cursors.Default
    End Sub

    '廠商管理-取消
    Private Sub btnCancel_manu_Click(sender As Object, e As EventArgs) Handles btnCancel_manu.Click
        Dim tp = sender.Parent
        ClearControls(tp)
        GetDataToDgv("SELECT * FROM manufacturer ORDER BY manu_code", dgvManufacturer)
        btnInsert_manu.Enabled = True
        btnModify_manu.Enabled = False
        btnDelete_manu.Enabled = False
    End Sub

    '廠商管理-新增
    Private Sub btnInsert_menu_Click(sender As Object, e As EventArgs) Handles btnInsert_manu.Click
        Dim dic As Dictionary(Of String, String) = CheckManufacturer(sender)
        If dic Is Nothing Then Exit Sub

        Dim dicDup As New Dictionary(Of String, String) From {
            {txtName_menu.Tag, txtName_menu.Text},
            {txtphone1_menu.Tag, txtphone1_menu.Text}
        }
        If Not CheckDuplication($"SELECT * FROM manufacturer", dicDup, dgvManufacturer) Then Exit Sub

        If InserTable("manufacturer", dic) Then
            btnCancel_manu.PerformClick()
            MsgBox("新增成功")
        End If
    End Sub

    '廠商管理-dgv點擊
    Private Sub dgvManufacturer_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvManufacturer.CellMouseClick
        Dim tp = sender.Parent
        ClearControls(tp)
        GetDataToControls(tp, sender.SelectedRows(0))
        btnInsert_manu.Enabled = False
        btnModify_manu.Enabled = True
        btnDelete_manu.Enabled = True
    End Sub

    '廠商管理-修改
    Private Sub btnModify_manu_Click(sender As Object, e As EventArgs) Handles btnModify_manu.Click
        Dim dic As Dictionary(Of String, String) = CheckManufacturer(sender)
        If dic Is Nothing Then Exit Sub
        If UpdateTable("manufacturer", dic, $"{txtNo_manu.Tag} = '{txtNo_manu.Text}'") Then
            btnCancel_manu.PerformClick()
            MsgBox("修改成功")
        End If
    End Sub

    '廠商管理-刪除
    Private Sub btnDelete_manu_Click(sender As Object, e As EventArgs) Handles btnDelete_manu.Click
        Dim tp As TabPage = sender.Parent
        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub
        If DeleteData("manufacturer", $"{txtNo_manu.Tag} = '{txtNo_manu.Text}'") Then
            btnCancel_manu.PerformClick()
            MsgBox("刪除成功")
        End If
    End Sub

    '廠商管理-查詢
    Private Sub btnQuery_manu_Click(sender As Object, e As EventArgs) Handles btnQuery_manu.Click
        Cursor = Cursors.WaitCursor
        GetDataToDgv("SELECT * FROM manufacturer " +
                                $"WHERE manu_name LIKE '%{txtQuery_manu.Text}%' OR manu_phone1 LIKE '%{txtQuery_manu.Text}%' " +
                                $"OR manu_phone2 LIKE '%{txtQuery_manu.Text}%' OR manu_code LIKE '%{txtQuery_manu.Text}%'", dgvManufacturer)
        Cursor = Cursors.Default
    End Sub

    Private Function CheckManufacturer(sender As Button)
        Dim dicReq As New Dictionary(Of String, Object) From {
             {"名稱", txtName_menu},
             {"聯絡人", txtContact_manu},
             {"代號", txtCode_manu},
             {"電話1", txtphone1_menu}
         }
        If Not CheckRequiredCol(dicReq) Then Return Nothing

        Dim tp As TabPage = sender.Parent
        Dim dic As New Dictionary(Of String, String)
        dic = tp.Controls.OfType(Of Control).Where(Function(ctrl) Not String.IsNullOrEmpty(ctrl.Tag) AndAlso ctrl.Tag <> "manu_id" AndAlso Not String.IsNullOrWhiteSpace(ctrl.Text)).
            ToDictionary(Function(ctrl) ctrl.Tag.ToString, Function(ctrl) ctrl.Text)
        Return dic
    End Function

    '客戶管理-取消
    Private Sub btnCancel_cus_Click(sender As Object, e As EventArgs) Handles btnCancel_cus.Click
        ClearControls(sender.Parent)
        GetDataToDgv("SELECT * FROM customer ORDER BY cus_code", dgvCustomer)
        btnInsert_cus.Enabled = True
        btnModify_cus.Enabled = False
        btnDelete_cus.Enabled = False
    End Sub

    '客戶管理-新增
    Private Sub btnInsert_cus_Click(sender As Object, e As EventArgs) Handles btnInsert_cus.Click
        Dim dic As Dictionary(Of String, String) = CheckCustomer(sender)
        If dic Is Nothing Then Exit Sub

        Dim dicDup As New Dictionary(Of String, String) From {
            {txtName_cus.Tag, txtName_cus.Text},
            {txtPhone1_cus.Tag, txtPhone1_cus.Text}
        }
        If Not CheckDuplication($"SELECT * FROM customer", dicDup, dgvCustomer) Then Exit Sub

        If InserTable("customer", dic) Then
            btnCancel_cus.PerformClick()
            MsgBox("新增成功")
        End If
    End Sub

    Private Function CheckCustomer(sender As Button)
        Dim dicReq As New Dictionary(Of String, Object) From {
             {"名稱", txtName_cus},
             {"聯絡人", txtContact_cus},
             {"代號", txtCode_cus},
             {"電話1", txtPhone1_cus}
         }
        If Not CheckRequiredCol(dicReq) Then Return Nothing

        Dim tp As TabPage = sender.Parent
        Dim dic As New Dictionary(Of String, String)
        dic = tp.Controls.OfType(Of Control).Where(Function(ctrl) Not String.IsNullOrEmpty(ctrl.Tag) AndAlso ctrl.Tag <> "cus_id" AndAlso Not String.IsNullOrWhiteSpace(ctrl.Text)).
            ToDictionary(Function(ctrl) ctrl.Tag.ToString, Function(ctrl) ctrl.Text)
        Return dic
    End Function

    '客戶管理-dgv點擊
    Private Sub dgvCustomer_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvCustomer.CellMouseClick
        Dim tp = sender.Parent
        ClearControls(tp)
        GetDataToControls(tp, sender.SelectedRows(0))
        btnInsert_cus.Enabled = False
        btnModify_cus.Enabled = True
        btnDelete_cus.Enabled = True
    End Sub

    '客戶管理-修改
    Private Sub btnModify_cus_Click(sender As Object, e As EventArgs) Handles btnModify_cus.Click
        Dim dic As Dictionary(Of String, String) = CheckCustomer(sender)
        If dic Is Nothing Then Exit Sub
        If UpdateTable("customer", dic, $"{txtID_cus.Tag} = '{txtID_cus.Text}'") Then
            btnCancel_cus.PerformClick()
            MsgBox("修改成功")
        End If
    End Sub

    '客戶管理-刪除
    Private Sub btnDelete_cus_Click(sender As Object, e As EventArgs) Handles btnDelete_cus.Click
        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub
        If DeleteData("customer", $"{txtID_cus.Tag} = '{txtID_cus.Text}'") Then
            btnCancel_cus.PerformClick()
            MsgBox("刪除成功")
        End If
    End Sub

    '客戶管理-查詢
    Private Sub btnQuery_cus_Click_1(sender As Object, e As EventArgs) Handles btnQuery_cus.Click
        Cursor = Cursors.WaitCursor
        GetDataToDgv("SELECT * FROM customer " +
                                $"WHERE cus_name LIKE '%{txtQuery_cus.Text}%' OR cus_phone1 LIKE '%{txtQuery_cus.Text}%' " +
                                $"OR cus_phone2 LIKE '%{txtQuery_cus.Text}%' OR cus_code LIKE '%{txtQuery_cus.Text}%'", dgvCustomer)
        Cursor = Cursors.Default
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        LoginForm1.Show()
    End Sub
End Class
