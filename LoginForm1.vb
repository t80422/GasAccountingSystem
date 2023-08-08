Public Class LoginForm1
    Private Sub OK_Click(sender As Object, e As EventArgs) Handles OK.Click
        Dim rows = SelectTable($"SELECT emp_acc, emp_psw, emp_perm FROM employee WHERE emp_acc = '{txtUsername.Text}' AND emp_psw = '{txtPassword.Text}'").Rows
        If rows.Count > 0 Then
            Form1.Show()
            Form1.TabControl1.Controls.OfType(Of TabPage).Where(Function(x) Not x.Text = "µn  ¥X" AndAlso Not Split(rows(0)("emp_perm"), ",").Contains(x.Text)).ToList.
                ForEach(Sub(y) y.Parent = Nothing)
            Form1.TabControl1.SelectedIndex = 0
            Hide()
        Else
            MsgBox("±b¸¹±K½X¿ù»~")
        End If
        txtUsername.Clear()
        txtPassword.Clear()
    End Sub

    Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles Cancel.Click
        Close()
    End Sub
End Class
