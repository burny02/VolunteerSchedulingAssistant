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
                Dim StudyFilter As String = Form1.FilterCombo12.Text
                Dim CohortFilter As String = Form1.FilterCombo26.Text
                Dim Offset As Long
                Dim VolFilter As String = Form1.FilterCombo27.Text
                Dim ProcFilter As String = Form1.FilterCombo28.Text
                Dim DayFilter As String = Form1.FilterCombo29.Text

                Try
                    Offset = Form1.TextBox2.Text
                Catch ex As Exception
                    MsgBox("Error occured. An offset must be a number")
                    Exit Sub
                End Try

                Dim StudyCrit As String = ""
                    Dim CohortCrit As String = ""
                    Dim VolCrit As String = ""
                Dim ProcCrit As String = ""
                Dim DayCrit As String = ""


                If StudyFilter <> "" Then StudyCrit = "Cohort.StudyID=" & Form1.FilterCombo12.SelectedValue & " AND "
                If CohortFilter <> "" Then CohortCrit = "Cohort.CohortID=" & Form1.FilterCombo26.SelectedValue & " AND "
                If VolFilter <> "" Then VolCrit = "Volunteer.VolID=" & Form1.FilterCombo27.SelectedValue & " AND "
                If ProcFilter <> "" Then ProcCrit = "StudySchedule.ProcID=" & Form1.FilterCombo28.SelectedValue & " AND "
                If DayFilter <> "" Then DayCrit = "StudySchedule.DaysPost=" & Form1.FilterCombo29.SelectedValue & " AND "

                Dim Criteria As String = StudyCrit & CohortCrit & VolCrit & ProcCrit & DayCrit
                Criteria = Left(Criteria, Len(Criteria) - 5)
                    Criteria = " WHERE " & Criteria

                Dim NumAffected As Long = 0

                Dim INList As String = OverClass.CreateCSVString("SELECT VolunteerScheduleID FROM " &
                "(Cohort INNER JOIN Volunteer ON Cohort.CohortID = Volunteer.CohortID) INNER JOIN " &
                "(StudySchedule INNER JOIN VolunteerSchedule ON StudySchedule.StudyScheduleID = VolunteerSchedule.StudyScheduleID) " &
                "ON Volunteer.VolID = VolunteerSchedule.VolID" & Criteria)

                Dim Cnt As Long = 0
                For Each c As Char In INList
                    If c = "," Then cnt += 1
                Next

                NumAffected = Cnt

                If MsgBox("The system will now set an offset of " & Offset & " minutes to all procedures with criteria..." & vbNewLine & vbNewLine &
                          "Study:  " & StudyFilter & vbNewLine &
                          IIf(CohortFilter = "", "", "Cohort: " & CohortFilter & vbNewLine) &
                          IIf(VolFilter = "", "", "Vol: " & VolFilter & vbNewLine) &
                          IIf(ProcFilter = "", "", "Procedure: " & ProcFilter & vbNewLine) &
                          IIf(DayFilter = "", "", "Day: " & DayFilter & vbNewLine) &
                          vbNewLine &
                          NumAffected & " records will be affected. Do you want to proceed?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

                    OverClass.ExecuteSQL("UPDATE VolunteerSchedule " &
                    "SET ProcOffset=" & Offset & " WHERE VolunteerScheduleID IN (" & INList & ")")

                    MsgBox("Offsets updated")


                End If

            Case "Button10"
                Call Saver(Form1.DataGridView8)
            Case "Button11"
                Call Saver(Form1.DataGridView9)
            Case "Button12"
                Call Saver(Form1.DataGridView10)
            Case "Button13"
                Call Saver(Form1.DataGridView11)
            Case "Button14"
                Call Saver(Form1.DataGridView12)
            Case "Button15"
                Dim OK As New ReportDisplay
                If CheckDates() = True Then
                    OK.ReportViewer1.ProcessingMode = ProcessingMode.Local
                    OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "Resource_Scheduling_System.VolunteerReport.rdlc"
                    OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet",
                                                                OverClass.TempDataTable("Select * FROM VolReport " &
                                                                                        "WHERE CohortID=" & Form1.FilterCombo2.SelectedValue &
                                                                                        " And CalcDate BETWEEN " & OverClass.SQLDate(Form1.DateTimePicker1.Value) &
                                                                                            " And " & OverClass.SQLDate(Form1.DateTimePicker2.Value))))

                    OK.ReportViewer1.RefreshReport()


                    Dim NextID As Long
                    Try
                        NextID = OverClass.TempDataTable("Select max(ArchiveID) FROM Reportarchive").Rows(0).Item(0) + 1

                    Catch ex As Exception
                        NextID = 1
                    End Try
                    Dim ArchiveType As String = "VolunteerReport"
                    Dim Criteria As String = "Dates " & Format(Form1.DateTimePicker1.Value, "dd-MMM-yyyy HHmm") _
                                             & " -> " & Format(Form1.DateTimePicker2.Value, "dd-MMM-yyyy HH:mm") _
                                             & vbNewLine & "Study:   " & Form1.FilterCombo1.Text _
                                             & vbNewLine & "Cohort " & Form1.FilterCombo2.Text

                    Dim pdfContent As Byte() = OK.ReportViewer1.LocalReport.Render("PDF", Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                    Dim pdfPath As String = ReportPath & NextID & ".pdf"
                    Dim pdfFile As New System.IO.FileStream(pdfPath, System.IO.FileMode.Create)
                    pdfFile.Write(pdfContent, 0, pdfContent.Length)
                    pdfFile.Close()

                    OverClass.ExecuteSQL("INSERT INTO ReportArchive (ArchiveID, ArchivePath, ArchiveUser, ArchiveType, ArchiveCriteria) " &
                                         "VALUES (" & NextID & ", '" & pdfPath & "', '" & OverClass.GetUserName & "', '" & ArchiveType & "', '" _
                                         & Criteria & "')")



                    OK.Visible = True
                    OK.ReportViewer1.Visible = True
                End If


            Case "Button16"
                If CheckDates() = True Then
                    Dim OK As New ReportDisplay

                    OK.ReportViewer1.ProcessingMode = ProcessingMode.Local
                    OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "Resource_Scheduling_System.StaffReport.rdlc"
                    OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet",
                                                                OverClass.TempDataTable("SELECT * FROM StaffReport " &
                                                                                        "WHERE CalcDate BETWEEN " & OverClass.SQLDate(Form1.DateTimePicker1.Value) &
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

                    OverClass.ExecuteSQL("INSERT INTO ReportArchive (ArchiveID, ArchivePath, ArchiveUser, ArchiveType, ArchiveCriteria) " &
                                     "VALUES (" & NextID & ", '" & pdfPath & "', '" & OverClass.GetUserName & "', '" & ArchiveType & "', '" _
                                     & Criteria & "')")



                    OK.Visible = True
                    OK.ReportViewer1.Visible = True


                End If

            Case "Button17"
                If CheckDates() = True Then
                    Dim OK As New ReportDisplay

                    OK.ReportViewer1.ProcessingMode = ProcessingMode.Local
                    OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "Resource_Scheduling_System.MasterReport.rdlc"
                    OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet",
                                                                OverClass.TempDataTable("SELECT * FROM StaffReport " &
                                                                                        "WHERE CalcDate BETWEEN " & OverClass.SQLDate(Form1.DateTimePicker1.Value) &
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

                    OverClass.ExecuteSQL("INSERT INTO ReportArchive (ArchiveID, ArchivePath, ArchiveUser, ArchiveType, ArchiveCriteria) " &
                                     "VALUES (" & NextID & ", '" & pdfPath & "', '" & OverClass.GetUserName & "', '" & ArchiveType & "', '" _
                                     & Criteria & "')")



                    OK.Visible = True
                    OK.ReportViewer1.Visible = True

                End If

            Case "Button18"

                If CheckDates() = True Then


                    Dim OK As New SOAForm

                    Dim SQLString As String = "TRANSFORM First(Format(DateAdd('n',IIF(ISNULL([ProcOffSet]),0,[ProcOffSet]),IIf([Approx]='Set Time',DateValue(DateAdd('d',[DaysPost],[TimepointDateTime]))+TimeValue([ProcTime]),DateAdd('n',DateDiff('n',TimeValue([DefaultTime]),TimeValue([ProcTime])),DateAdd('d',[DaysPost],[TimepointDateTime])))),'dd-MMM') " &
    "& Chr(13) & Chr(10) & Format(DateAdd('n',IIF(ISNULL([ProcOffSet]),0,[ProcOffSet]),IIf([Approx]='Set Time',DateValue(DateAdd('d',[DaysPost],[TimepointDateTime]))+TimeValue([ProcTime]),DateAdd('n',DateDiff('n',TimeValue([DefaultTime]),TimeValue([ProcTime])),DateAdd('d',[DaysPost],[TimepointDateTime])))),'HH:mm') " &
    "& Chr(13) & Chr(10) & IIf(IsNull([FullName]),'',Left([FullName],1) & '-' & left(Right([FullName],Len([FullName])-InStr([FullName],' ')),1))) AS CalcDate " &
    "SELECT StudyTimepoint.StudyID, a.StudyScheduleID, e.ProcName, a.DaysPost, a.ProcTime, e.ProcOrd " &
    "FROM ((Study INNER JOIN (Cohort INNER JOIN Volunteer AS d ON Cohort.CohortID = d.CohortID) ON Study.StudyID = Cohort.StudyID) INNER JOIN StudyTimepoint ON Study.StudyID = StudyTimepoint.StudyID) INNER JOIN ((ProcTask AS e INNER JOIN (StudySchedule AS a INNER JOIN VolunteerTimepoint AS f ON a.StudyTimepointID = f.StudyTimepointID) ON e.ProcID = a.ProcID) " &
    "INNER JOIN (VolunteerSchedule AS c LEFT JOIN Staff ON c.SharepointID = Staff.SharepointID) ON a.StudyScheduleID = c.StudyScheduleID) ON (d.VolID = f.VolID) AND (d.VolID = c.VolID) AND (StudyTimepoint.StudyTimepointID = a.StudyTimepointID) " &
    "WHERE (((StudyTimepoint.StudyID)=" & Form1.FilterCombo1.SelectedValue.ToString & ") " &
    "AND ((DateAdd('n',IIF(ISNULL([ProcOffSet]),0,[ProcOffSet]),IIf([Approx]='Set Time',DateValue(DateAdd('d',[DaysPost],[TimepointDateTime]))+TimeValue([ProcTime]),DateAdd('n',DateDiff('n',TimeValue([DefaultTime]),TimeValue([ProcTime])),DateAdd('d',[DaysPost],[TimepointDateTime]))))) " &
    "BETWEEN " & OverClass.SQLDate(Form1.DateTimePicker1.Value) &
    " AND " & OverClass.SQLDate(Form1.DateTimePicker2.Value) & ")) " &
    "GROUP BY StudyTimepoint.StudyID, a.StudyScheduleID, e.ProcName, a.DaysPost, a.ProcTime, e.ProcOrd " &
    "ORDER BY a.DaysPost, a.ProcTime, e.ProcOrd, 'Room ' & [RoomNo] & Chr(13) & Chr(10) & [Initials] & Chr(13) & Chr(10) & [RVLNo] " &
    "PIVOT 'Room ' & [RoomNo] & Chr(13) & Chr(10) & [Initials] & Chr(13) & Chr(10) & [RVLNo]"



                    OverClass.CreateDataSet(SQLString, OK.BindingSource1, OK.DataGridView1)
                    OK.DataGridView1.Columns("StudyID").Visible = False
                    OK.DataGridView1.Columns("StudyScheduleID").Visible = False
                    OK.DataGridView1.Columns("DaysPost").Visible = False
                    OK.DataGridView1.Columns("ProcTime").Visible = False
                    OK.DataGridView1.Columns("ProcOrd").Visible = False
                    OK.DataGridView1.Columns("ProcName").DisplayIndex = 0
                    OK.DataGridView1.Columns("ProcName").HeaderText = "Procedure"
                    OK.DataGridView1.Columns("ProcName").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                    OK.DataGridView1.Columns("ProcName").Width = 200

                    For Each column In OK.DataGridView1.Columns
                        column.SortMode = DataGridViewColumnSortMode.NotSortable
                    Next

                    Dim i As Long = 0
                    Dim RoomNo As Long = 0
                    Dim DisplayNumber As Long = 1

                    Do While i < 200

                        For Each column In OK.DataGridView1.Columns

                            If column.headertext Like "Room*" Then

                                RoomNo = CInt(Trim(Replace(Left(column.headertext, InStr(column.headertext, vbNewLine)), "Room ", vbNullString)))

                                If i = RoomNo Then
                                    column.displayindex = DisplayNumber
                                    column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                                    column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                                    DisplayNumber = DisplayNumber + 1
                                End If

                            Else

                                Continue For

                            End If

                        Next
                        i = i + 1

                    Loop


                    OK.ShowDialog()

                End If


            Case "Button20"

                MsgBox("Please ensure to save changes to see up to date report")


                Dim dt2 As DataTable
                dt2 = OverClass.TempDataTable("SELECT * FROM SchedulePreview " &
                                              "WHERE StudyID=" & Form1.FilterCombo22.SelectedValue.ToString)

                Dim OK As New ReportDisplay

                OK.ReportViewer1.ProcessingMode = ProcessingMode.Local
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "Resource_Scheduling_System.SchedulePreview.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet",
                                                            dt2))

                OK.ReportViewer1.RefreshReport()

                Dim NextID As Long
                Try
                    NextID = OverClass.TempDataTable("SELECT max(ArchiveID) FROM Reportarchive").Rows(0).Item(0) + 1
                Catch ex As Exception
                    NextID = 1
                End Try

                Dim ArchiveType As String = "SchedulePreview"
                Dim Criteria As String = "Study: " & Form1.FilterCombo22.Text

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

            Case "Button21"

                Dim Dt As DataTable = OverClass.TempDataTable("SELECT DefaultTime FROM StudyTimepoint WHERE StudyTimepointID=" &
                                                              Form1.FilterCombo21.SelectedValue)
                Dim TempTime As String = ""

                Try
                    TempTime = Dt.Rows(0).Item(0).ToString
                Catch ex As Exception
                End Try

                InputBox("Default Time:", "Default Time", TempTime)

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
