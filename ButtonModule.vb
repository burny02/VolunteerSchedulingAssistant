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
                Call Saver(Form1.DataGridView12)
            Case "Button15"
                Dim OK As New ReportDisplay
                If CheckDates() = True Then
                    OK.ReportViewer1.ProcessingMode = ProcessingMode.Local
                    OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "VolunteerSchedulingAssistant.VolunteerReport.rdlc"
                    OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                                OverClass.TempDataTable("SELECT * FROM VolReport " & _
                                                                                        "WHERE CohortID=" & Form1.ComboBox18.SelectedValue & _
                                                                                        " AND CalcDate BETWEEN " & OverClass.SQLDate(Form1.DateTimePicker1.Value) & _
                                                                                            " AND " & OverClass.SQLDate(Form1.DateTimePicker2.Value))))

                    OK.ReportViewer1.RefreshReport()


                    Dim NextID As Long
                    Try
                        NextID = OverClass.TempDataTable("SELECT max(ArchiveID) FROM Reportarchive").Rows(0).Item(0) + 1

                    Catch ex As Exception
                        NextID = 1
                    End Try
                    Dim ArchiveType As String = "VolunteerReport"
                    Dim Criteria As String = "Dates: " & Format(Form1.DateTimePicker1.Value, "dd-MMM-yyyy HH:mm") _
                                             & " -> " & Format(Form1.DateTimePicker2.Value, "dd-MMM-yyyy HH:mm") _
                                             & vbNewLine & "Study: " & Form1.ComboBox17.Text _
                                             & vbNewLine & "Cohort: " & Form1.ComboBox18.Text

                    Dim pdfContent As Byte() = OK.ReportViewer1.LocalReport.Render("PDF", Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                    Dim pdfPath As String = ReportPath & NextID & ".pdf"
                    Dim pdfFile As New System.IO.FileStream(pdfPath, System.IO.FileMode.Create)
                    pdfFile.Write(pdfContent, 0, pdfContent.Length)
                    pdfFile.Close()

                    OverClass.ExecuteSQL("INSERT INTO ReportArchive (ArchiveID, ArchivePath, ArchiveUser, ArchiveType, ArchiveCriteria) " & _
                                         "VALUES (" & NextID & ", '" & pdfPath & "', '" & OverClass.GetUserName & "', '" & ArchiveType & "', '" _
                                         & Criteria & "')")



                    OK.Visible = True
                    OK.ReportViewer1.Visible = True
                End If


            Case "Button16"
                If CheckDates() = True Then
                    Dim OK As New ReportDisplay

                    OK.ReportViewer1.ProcessingMode = ProcessingMode.Local
                    OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "VolunteerSchedulingAssistant.StaffReport.rdlc"
                    OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                                OverClass.TempDataTable("SELECT * FROM StaffReport " & _
                                                                                        "WHERE CalcDate BETWEEN " & OverClass.SQLDate(Form1.DateTimePicker1.Value) & _
                                                                                        " AND " & OverClass.SQLDate(Form1.DateTimePicker2.Value))))
                    OK.ReportViewer1.RefreshReport()

                    'Create PDF file on disk

                    Dim NextID As Long
                    Try
                        NextID = OverClass.TempDataTable("SELECT max(ArchiveID) FROM Reportarchive").Rows(0).Item(0) + 1
                    Catch ex As Exception
                        NextID = 1
                    End Try

                    Dim ArchiveType As String = "StaffReport"
                    Dim Criteria As String = "Dates: " & Format(Form1.DateTimePicker1.Value, "dd-MMM-yyyy HH:mm") & " -> " _
                                             & Format(Form1.DateTimePicker2.Value, "dd-MMM-yyyy HH:mm")


                    Dim pdfContent As Byte() = OK.ReportViewer1.LocalReport.Render("PDF", Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                    Dim pdfPath As String = ReportPath & NextID & ".pdf"
                    Dim pdfFile As New System.IO.FileStream(pdfPath, System.IO.FileMode.Create)
                    pdfFile.Write(pdfContent, 0, pdfContent.Length)
                    pdfFile.Close()

                    OverClass.ExecuteSQL("INSERT INTO ReportArchive (ArchiveID, ArchivePath, ArchiveUser, ArchiveType, ArchiveCriteria) " & _
                                     "VALUES (" & NextID & ", '" & pdfPath & "', '" & OverClass.GetUserName & "', '" & ArchiveType & "', '" _
                                     & Criteria & "')")



                    OK.Visible = True
                    OK.ReportViewer1.Visible = True


                End If

            Case "Button17"
                If CheckDates() = True Then
                    Dim OK As New ReportDisplay

                    OK.ReportViewer1.ProcessingMode = ProcessingMode.Local
                    OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "VolunteerSchedulingAssistant.MasterReport.rdlc"
                    OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                                OverClass.TempDataTable("SELECT * FROM StaffReport " & _
                                                                                        "WHERE CalcDate BETWEEN " & OverClass.SQLDate(Form1.DateTimePicker1.Value) & _
                                                                                        " AND " & OverClass.SQLDate(Form1.DateTimePicker2.Value))))

                    OK.ReportViewer1.RefreshReport()

                    Dim NextID As Long
                    Try
                        NextID = OverClass.TempDataTable("SELECT max(ArchiveID) FROM Reportarchive").Rows(0).Item(0) + 1
                    Catch ex As Exception
                        NextID = 1
                    End Try

                    Dim ArchiveType As String = "MasterReport"
                    Dim Criteria As String = "Dates: " & Format(Form1.DateTimePicker1.Value, "dd-MMM-yyyy HH:mm") _
                                             & " -> " & Format(Form1.DateTimePicker2.Value, "dd-MMM-yyyy HH:mm")

                    Dim pdfContent As Byte() = OK.ReportViewer1.LocalReport.Render("PDF", Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                    Dim pdfPath As String = ReportPath & NextID & ".pdf"
                    Dim pdfFile As New System.IO.FileStream(pdfPath, System.IO.FileMode.Create)
                    pdfFile.Write(pdfContent, 0, pdfContent.Length)
                    pdfFile.Close()

                    OverClass.ExecuteSQL("INSERT INTO ReportArchive (ArchiveID, ArchivePath, ArchiveUser, ArchiveType, ArchiveCriteria) " & _
                                     "VALUES (" & NextID & ", '" & pdfPath & "', '" & OverClass.GetUserName & "', '" & ArchiveType & "', '" _
                                     & Criteria & "')")



                    OK.Visible = True
                    OK.ReportViewer1.Visible = True

                End If

            Case "Button18"

            Case "Button19"

            Case "Button20"


        End Select



    End Sub

    Private Function CheckDates() As Boolean

        Dim dater1, dater2 As Date
        dater1 = Form1.DateTimePicker1.Value
        dater2 = Form1.DateTimePicker2.Value

        If dater1 >= dater2 Then
            MsgBox("'Date To' must be greater than 'Date From'")
            CheckDates = False
        Else
            CheckDates = True
        End If

    End Function

End Module
