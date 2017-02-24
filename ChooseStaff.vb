Public Class ChooseStaff

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        For Each row In Form1.DataGridView11.Rows

            Form1.DataGridView11.Item("SharepointID", row.index).Value = Me.ComboBox1.SelectedValue

        Next

        Me.Close()

    End Sub

    Private Sub ChooseStaff_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.ComboBox1.DataSource = OverClass.TempDataTable("SELECT SharepointID, FullName " &
                                                         "FROM STAFF WHERE Hidden=False ORDER BY FullName ASC")
        Me.ComboBox1.ValueMember = "SharepointID"
        Me.ComboBox1.DisplayMember = "FullName"

    End Sub
End Class