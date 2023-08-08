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

        '初始化各TabPage
        btnCancel_emp.PerformClick()
        btnCancel_perm_Click(btnCancel_perm, e)
        btnCancel_manu_Click(btnCancel_manu, e)

    End Sub

    Private Sub tpLogOut_Enter(sender As Object, e As EventArgs) Handles tpLogOut.Enter
        If MessageBox.Show("確定要登出嗎??", "登出", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = System.Windows.Forms.DialogResult.OK Then
            LoginForm1.Show()
            Close()
        End If
    End Sub

    '員工管理-取消
    Private Sub btnCancel_emp_Click(sender As Object, e As EventArgs) Handles btnCancel_emp.Click
        Dim tp = sender.Parent
        ClearControls(tp)
        GetDataToDgv("SELECT emp_id,  emp_name,  emp_phone,  emp_address,  emp_ecp,  emp_ect,  emp_acc,  emp_psw, emp_identity_number, emp_birthday FROM employee", dgvEmployee)
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
        If InserTable("employee", dic) Then
            btnCancel_emp.PerformClick()
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
            MsgBox("刪除成功")
        End If
    End Sub

    '員工管理-查詢
    Private Sub btnQuery_cus_Click(sender As Object, e As EventArgs) Handles btnQuery_cus.Click
        GetDataToDgv($"SELECT * FROM employee WHERE emp_name LIKE '%{txtQuery_cus.Text}%' OR emp_phone LIKE '%{txtQuery_cus.Text}%'", dgvEmployee)
        dgvEmployee.Columns("emp_perm").Visible = False
        MsgBox("查詢完畢")
    End Sub

    '搜尋欄位按下"Enter"即可搜尋
    Private Sub txtQuery_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtQuery_cus.KeyPress, txtQuery_perm.KeyPress, txtQuery_manu.KeyPress
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
        'Dim perms = String.Join(",", flpPermissions.Controls.OfType(Of CheckBox).Select(Function(chk) If(chk.Checked, "1", "0")))

        Dim dic As New Dictionary(Of String, String) From {{"emp_perm", perms}}
        If UpdateTable("employee", dic, $"{txtID_perm.Tag} = '{txtID_perm.Text}'") Then
            btnCancel_perm.PerformClick()
            MsgBox("修改成功")
        End If
    End Sub

    '權限管理-查詢
    Private Sub btnQuery_perm_Click(sender As Object, e As EventArgs) Handles btnQuery_perm.Click
        GetDataToDgv($"SELECT emp_id, emp_name, emp_acc, emp_perm FROM employee WHERE emp_name LIKE '%{txtQuery_perm.Text}%' OR emp_acc LIKE '%{txtQuery_perm.Text}%'", dgvPermissions)
        dgvPermissions.Columns("emp_perm").Visible = False
        MsgBox("查詢完畢")
    End Sub

    '廠商管理-取消
    Private Sub btnCancel_manu_Click(sender As Object, e As EventArgs) Handles btnCancel_manu.Click
        Dim tp = sender.Parent
        ClearControls(tp)
        GetDataToDgv("SELECT * FROM manufacturer", dgvManufacturer)
        btnInsert_manu.Enabled = True
        btnModify_manu.Enabled = False
        btnDelete_manu.Enabled = False
    End Sub

    '廠商管理-新增
    Private Sub btnInsert_menu_Click(sender As Object, e As EventArgs) Handles btnInsert_manu.Click
        Dim dic As Dictionary(Of String, String) = CheckManufacturer(sender)
        If dic Is Nothing Then Exit Sub
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
        GetDataToDgv($"SELECT * FROM manufacturer WHERE manu_name LIKE '%{txtQuery_manu.Text}%' OR manu_phone1 LIKE '%{txtQuery_manu.Text}%'", dgvManufacturer)
        MsgBox("查詢完畢")
    End Sub

    Private Function CheckManufacturer(sender As Button)
        Dim dicReq As New Dictionary(Of String, Object) From {
             {"名稱", txtName_menu},
             {"聯絡人", txtContact},
             {"電話1", txtphone1_menu}
         }
        If Not CheckRequiredCol(dicReq) Then Return Nothing

        Dim tp As TabPage = sender.Parent
        Dim dic As New Dictionary(Of String, String)
        dic = tp.Controls.OfType(Of Control).Where(Function(ctrl) Not String.IsNullOrEmpty(ctrl.Tag) AndAlso ctrl.Tag <> "manu_id" AndAlso Not String.IsNullOrWhiteSpace(ctrl.Text)).
            ToDictionary(Function(ctrl) ctrl.Tag.ToString, Function(ctrl) ctrl.Text)
        Return dic
    End Function

    Private Sub btnCancel_cus_Click(sender As Object, e As EventArgs) Handles btnCancel_cus.Click

    End Sub
End Class
