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
            Case "Button14"
                Call Saver(Form1.DataGridView12)
                Call Saver(Form1.DataGridView12)
                Call Saver(Form1.DataGridView12)
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

                Dim dt As DataTable
                dt = OverClass.TempDataTable("SELECT StudyTimepointID, TimepointName " & _
                                             "FROM StudyTimepoint WHERE StudyID=" & Form1.ComboBox17.SelectedValue.ToString)
                dt.Columns.Add("TimepointDateTime", System.Type.GetType("System.DateTime"))


                Dim Accepted As Boolean = False
                Dim Temp As String
                Dim TempDate As Date

                For Each row In dt.Rows

                    Accepted = False
                    Dim TimepointName As String = row.item("TimepointName")

                    Do While Accepted = False

                        Temp = InputBox("Input " & TimepointName & " Date", , "01-Jan-2010 10:00")

                        Try
                            TempDate = CDate(Temp)
                            If Format(TempDate, "HH:mm") = "00:00" Then Throw New System.Exception
                            row.item("TimepointDateTime") = TempDate

                        Catch ex As Exception
                            MsgBox("Must enter a valid Date/Time to continue")
                            Continue Do

                        End Try

                        Accepted = True

                    Loop

                Next

                Dim dt2 As DataTable
                dt2 = OverClass.TempDataTable("SELECT * FROM SchedulePreview " & _
                                              "WHERE StudyID=" & Form1.ComboBox17.SelectedValue.ToString)


                Dim ReportData =
                    From a In dt.AsEnumerable()
                    Join b In dt2.AsEnumerable()
                    On
                       a.Field(Of Int32)("StudyTimepointID") Equals b.Field(Of Int32)("StudyTimepointID")
                    Select Order = b.Field(Of Int32)("ProcOrd"), TimepointName = a.Field(Of String)("TimepointName"),
                    TimepointDateTime = a.Field(Of Date)("TimepointDateTime"),
                    StudyCode = b.Field(Of String)("StudyCode"), ProcName = b.Field(Of String)("ProcName"),
                    CalcDate = If(b.Field(Of String)("Approx") = "Set Time",
                CDate(DateValue(DateAdd("d", b.Field(Of Int32)("DaysPost"), a.Field(Of Date)("TimepointDateTime"))) + " " + TimeValue(b.Field(Of Date)("SetTime"))),
                DateAdd("n", b.Field(Of Int16)("MinsPost"), DateAdd("h", b.Field(Of Int16)("HoursPost"),
                DateAdd("d", b.Field(Of Int32)("DaysPost"), a.Field(Of Date)("TimepointDateTime")))))


                Dim OK As New ReportDisplay

                OK.ReportViewer1.ProcessingMode = ProcessingMode.Local
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "VolunteerSchedulingAssistant.SchedulePreview.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                            ReportData))

                OK.ReportViewer1.RefreshReport()

                Dim NextID As Long
                Try
                    NextID = OverClass.TempDataTable("SELECT max(ArchiveID) FROM Reportarchive").Rows(0).Item(0) + 1
                Catch ex As Exception
                    NextID = 1
                End Try

                Dim ArchiveType As String = "SchedulePreview"
                Dim Criteria As String = "Study: " & Form1.ComboBox17.Text

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

            Case "Button20"
                MsgBox("Please ensure to save changes to see up to date report")
                Dim dt As DataTable
                dt = OverClass.TempDataTable("SELECT StudyTimepointID, TimepointName " & _
                                             "FROM StudyTimepoint WHERE StudyID=" & Form1.ComboBox4.SelectedValue.ToString)
                dt.Columns.Add("TimepointDateTime", System.Type.GetType("System.DateTime"))


                Dim Accepted As Boolean = False
                Dim Temp As String
                Dim TempDate As Date

                For Each row In dt.Rows

                    Accepted = False
                    Dim TimepointName As String = row.item("TimepointName")

                    Do While Accepted = False

                        Temp = InputBox("Input " & TimepointName & " Date", , "01-Jan-2010 10:00")

                        Try
                            TempDate = CDate(Temp)
                            If Format(TempDate, "HH:mm") = "00:00" Then Throw New System.Exception
                            row.item("TimepointDateTime") = TempDate

                        Catch ex As Exception
                            MsgBox("Must enter a valid Date/Time to continue")
                            Continue Do

                        End Try

                        Accepted = True

                    Loop

                Next

                Dim dt2 As DataTable
                dt2 = OverClass.TempDataTable("SELECT * FROM SchedulePreview " & _
                                              "WHERE StudyID=" & Form1.ComboBox4.SelectedValue.ToString)


                Dim ReportData =
                    From a In dt.AsEnumerable()
                    Join b In dt2.AsEnumerable()
                    On
                       a.Field(Of Int32)("StudyTimepointID") Equals b.Field(Of Int32)("StudyTimepointID")
                    Select Order = b.Field(Of Int32)("ProcOrd"), TimepointName = a.Field(Of String)("TimepointName"),
                    TimepointDateTime = a.Field(Of Date)("TimepointDateTime"),
                    StudyCode = b.Field(Of String)("StudyCode"), ProcName = b.Field(Of String)("ProcName"),
                    CalcDate = If(b.Field(Of String)("Approx") = "Set Time",
                CDate(DateValue(DateAdd("d", b.Field(Of Int32)("DaysPost"), a.Field(Of Date)("TimepointDateTime"))) + " " + TimeValue(b.Field(Of Date)("SetTime"))),
                DateAdd("n", b.Field(Of Int16)("MinsPost"), DateAdd("h", b.Field(Of Int16)("HoursPost"),
                DateAdd("d", b.Field(Of Int32)("DaysPost"), a.Field(Of Date)("TimepointDateTime")))))


                Dim OK As New ReportDisplay

                OK.ReportViewer1.ProcessingMode = ProcessingMode.Local
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "VolunteerSchedulingAssistant.SchedulePreview.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                            ReportData))

                OK.ReportViewer1.RefreshReport()

                Dim NextID As Long
                Try
                    NextID = OverClass.TempDataTable("SELECT max(ArchiveID) FROM Reportarchive").Rows(0).Item(0) + 1
                Catch ex As Exception
                    NextID = 1
                End Try

                Dim ArchiveType As String = "SchedulePreview"
                Dim Criteria As String = "Study: " & Form1.ComboBox4.Text

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

            Case "Button5"

                Dim AssAll As New ChooseStaff

                AssAll.ShowDialog()
                

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
