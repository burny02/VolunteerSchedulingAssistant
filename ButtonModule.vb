Imports Microsoft.Reporting.WinForms

Module ButtonModule

    Public Sub ButtonSpecifics(sender As Object, e As EventArgs)

        Dim ctl As Object = Nothing

        Select Case sender.name.ToString

            Case "Button1"
                Call Saver(Form1.DataGridView1)
            Case "Button3"
                Call Saver(Form1.DataGridView2)
            Case "Button4"
                Call Saver(Form1.DataGridView3)
            Case "Button6"
                Call Saver(Form1.DataGridView5)
            Case "Button7"
                Call Saver(Form1.DataGridView6)
            Case "Button8"
                Call Saver(Form1.DataGridView7)
            Case "Button9"
                Call Saver(Form1.DataGridView8)
            Case "Button10"
                Call Generator()
            Case "Button11"
                Call Saver(Form1.DataGridView9)
            Case "Button12"
                Call Saver(Form1.DataGridView10)
            Case "Button13"
                Call Saver(Form1.DataGridView11)
            Case "Button5"
                Call Saver(Form1.DataGridView4)
            Case "Button14"
                Call Saver(Form1.DataGridView4)
            Case "Button15"
                'If CheckCohort() = True Then
                Dim OK As New ReportDisplay
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.ProcessingMode = ProcessingMode.Local
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "VolunteerSchedulingAssistant.VolunteerReport.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                            OverClass.TempDataTable("SELECT * FROM VolReport")))

                OK.ReportViewer1.RefreshReport()


            Case "Button16"
                'If CheckDates() = True Then
                Dim OK As New ReportDisplay
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.ProcessingMode = ProcessingMode.Local
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "VolunteerSchedulingAssistant.StaffReport.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                            OverClass.TempDataTable("SELECT * FROM StaffReport")))

                OK.ReportViewer1.RefreshReport()

            Case "Button17"
                'If CheckDates() = True Then
            Case "Button18"
                'If CheckDates() = True Then
            Case "Button19"
                'If CheckStudy() = True Then

        End Select



    End Sub

    Private Function CheckDates() As Boolean



    End Function

    Private Function CheckCohort() As Boolean



    End Function

    Private Function CheckStudy() As Boolean



    End Function

End Module
