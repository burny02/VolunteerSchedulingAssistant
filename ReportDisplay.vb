Public Class ReportDisplay

    Private Sub ReportViewer1_Load_1(sender As Object, e As EventArgs) Handles ReportViewer1.Load

        Me.WindowState = FormWindowState.Maximized
        Me.Text = SolutionName

    End Sub
End Class