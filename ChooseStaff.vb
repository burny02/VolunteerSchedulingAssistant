Public Class ChooseStaff

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        For Each row In Form1.DataGridView11.Rows

            Form1.DataGridView11.Item("StaffID", row.index).Value = Me.ComboBox1.SelectedValue

        Next

        Me.Close()

    End Sub

    Private Sub ChooseStaff_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.ComboBox1.DataSource = OverClass.TempDataTable("SELECT StaffID, FName & ' ' & SName AS FullName " & _
                                                         "FROM STAFF WHERE Hidden=False ORDER BY FName ASC")
        Me.ComboBox1.ValueMember = "StaffID"
        Me.ComboBox1.DisplayMember = "FullName"

    End Sub
End Class