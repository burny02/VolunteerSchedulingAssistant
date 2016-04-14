Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        Call StartUp()

        Try
            Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside" & vbNewLine & "Version: " & System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        Catch
            Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside"
        End Try

        Me.Text = SolutionName


    End Sub

    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If OverClass.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        OverClass.ResetCollection()

        Select Case e.TabPage.Text

            Case "Set Up"
                Me.TabControl2.SelectedIndex = 0
                Me.TabControl2_Selecting(Me.TabControl2, New TabControlCancelEventArgs(TabPage3, 0, False, TabControlAction.Selecting))
            Case "Studies"
                Me.TabControl3.SelectedIndex = 0
                Me.TabControl3_Selecting(Me.TabControl3, New TabControlCancelEventArgs(TabPage10, 0, False, TabControlAction.Selecting))
            Case "Volunteers"
                Me.TabControl4.SelectedIndex = 0
                Me.TabControl4_Selecting(Me.TabControl4, New TabControlCancelEventArgs(TabPage7, 0, False, TabControlAction.Selecting))
            Case "Staffing"
                Me.TabControl5.SelectedIndex = 0
                Me.TabControl5_Selecting(Me.TabControl5, New TabControlCancelEventArgs(TabPage18, 0, False, TabControlAction.Selecting))
            Case "Reports"
                FilterCombo1.AllowBlanks = False
                FilterCombo2.AllowBlanks = False
                FilterCombo1.SetAsExternalSource("StudyID", "StudyCode",
                                                 "Select StudyID, StudyCode FROM Study", OverClass)
                FilterCombo2.SetAsExternalSource("CohortID", "CohortName",
                                                 "Select CohortID, CohortName FROM Cohort WHERE StudyID=" &
                                                FilterCombo2.SetCmbPointer(FilterCombo1), OverClass)

            Case "Report Archive"
                ctl = Me.DataGridView13

            Case "Offsets"
                Me.TabControl6.SelectedIndex = 0
                Me.TabControl6_Selecting(Me.TabControl6, New TabControlCancelEventArgs(TabPage16, 0, False, TabControlAction.Selecting))

        End Select


        If Not IsNothing(ctl) Then Call Specifics(ctl)

    End Sub

    Public Sub Specifics(ctl As Object)

        If IsNothing(ctl) Then Exit Sub

        Dim SQLCode As String = vbNullString

        Select Case ctl.name

            Case "DataGridView1"
                ctl.Columns(0).Visible = False
                ctl.columns(1).headertext = "Procedure"
                ctl.columns(2).headertext = "Minutes Taken"
                ctl.columns(3).headertext = "Order"

            Case "DataGridView2"
                ctl.Columns(0).Visible = False
                ctl.columns(1).headertext = "Name"
                ctl.columns(2).headertext = "Surname"

            Case "DataGridView3"
                ctl.Columns(0).visible = False
                Dim cmb As New DataGridViewButtonColumn
                cmb.HeaderText = "Colour Picker"
                cmb.Name = "ColourPick"
                cmb.UseColumnTextForButtonValue = True
                cmb.Text = "Pick Colour"
                ctl.Columns("Colour").ReadOnly = True
                ctl.columns.add(cmb)

            Case "DataGridView5"

                SQLCode = "Select StudyID, StudyTimepointID, TimepointName FROM StudyTimepoint ORDER BY TimepointName ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)

                FilterCombo19.AllowBlanks = False
                FilterCombo19.SetAsExternalSource("StudyID", "StudyCode",
                           "Select StudyID, StudyCode FROM Study", OverClass)
                FilterCombo19.SetDGVDefault(ctl, "StudyID")

                ctl.Columns("StudyTimepointID").visible = False
                ctl.Columns("StudyID").visible = False
                ctl.columns("TimepointName").headertext = "Timepoint Name"

            Case "DataGridView6"

                ctl.autogeneratecolumns = True


                SQLCode = "Select StudyID, a.StudyTimepointID, StudyScheduleID, ProcID, DaysPost, Approx, ProcTime " &
                    "FROM StudySchedule a INNER JOIN StudyTimepoint b On a.StudyTimepointID= b.StudyTimepointID " &
                    "ORDER BY DaysPost ASC, CDate(format(ProcTime,'HH:mm')) ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)

                FilterCombo22.AllowBlanks = False
                FilterCombo22.SetAsExternalSource("StudyID", "StudyCode",
                           "SELECT b.StudyID, StudyCode FROM Study a INNER JOIN StudyTimepoint b " &
                           "ON a.studyid=b.studyid", OverClass)
                FilterCombo22.SetDGVDefault(ctl, "StudyID")

                FilterCombo21.AllowBlanks = False
                FilterCombo21.SetAsExternalSource("StudyTimepointID", "TimepointName",
                           "SELECT StudyTimepointID, TimepointName FROM StudyTimepoint WHERE StudyID=" &
                           FilterCombo21.SetCmbPointer(FilterCombo22), OverClass)
                FilterCombo21.SetDGVDefault(ctl, "StudyTimepointID")


                FilterCombo23.SetAsExternalSource("ProcID", "ProcName",
                          "SELECT ProcID, ProcName FROM ProcTask WHERE ProcID IN (" &
                FilterCombo23.SetCmbPointer(OverClass.CurrentDataSet.Tables(0).Columns("ProcID")) & ")", OverClass)



                ctl.Columns("StudyTimepointID").visible = False
                ctl.Columns("StudyID").visible = False
                ctl.Columns("StudyScheduleID").visible = False
                ctl.columns("ProcID").visible = False
                Dim cmb As New DataGridViewComboBoxColumn
                cmb.HeaderText = "Procedure"
                cmb.DataSource = OverClass.TempDataTable("SELECT ProcID, ProcName" &
                                                     " FROM ProcTask ORDER BY ProcName ASC")
                cmb.ValueMember = "ProcID"
                cmb.DisplayMember = "ProcName"
                cmb.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("ProcID").ToString
                ctl.columns.add(cmb)
                cmb.DisplayIndex = 1
                ctl.columns("DaysPost").headertext = "Days After Timepoint"
                ctl.columns("ProcTime").headertext = "Procedure Time"
                ctl.Columns("ProcTime").DefaultCellStyle.Format = "HH:mm"
                cmb.Name = "PickProc"
                Dim cmb2 As New DataGridViewComboBoxColumn
                cmb2.HeaderText = "Timepoint"
                cmb2.Items.Add("Approx")
                cmb2.Items.Add("Set Time")
                cmb2.Items.Add("Timed")
                cmb2.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("Approx").ToString
                ctl.columns.add(cmb2)
                ctl.columns("Approx").visible = False
                cmb2.DisplayIndex = 2
                cmb2.Name = "PickTimepoint"
                Dim cmb3 As New DataGridViewComboBoxColumn
                cmb3.DataSource = OverClass.TempDataTable("SELECT ProcID, MinsTaken " &
                                                          "FROM ProcTask")
                cmb3.ValueMember = "ProcID"
                cmb3.DisplayMember = "MinsTaken"
                cmb3.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("ProcID").ToString
                cmb3.Name = "MinsTaken"
                ctl.columns.add(cmb3)
                cmb3.Visible = False
                Dim cmb4 As New DataGridViewImageColumn
                ctl.autogeneratecolumns = False
                cmb4.DisplayIndex = 10
                cmb4.HeaderText = "Delete Procedure"
                cmb4.Image = My.Resources.Remove
                cmb4.ImageLayout = DataGridViewImageCellLayout.Zoom
                ctl.columns.add(cmb4)
                cmb4.Name = "DeleteButton"



            Case "DataGridView7"

                OverClass.ResetCollection()

                SQLCode = "SELECT StudyID, CohortID, CohortName, NumVols, Generated" &
                    " FROM Cohort ORDER BY CohortName ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)

                FilterCombo20.AllowBlanks = False
                FilterCombo20.SetAsExternalSource("StudyID", "StudyCode",
                           "SELECT StudyID, StudyCode FROM Study", OverClass)
                FilterCombo20.SetDGVDefault(ctl, "StudyID")

                ctl.Columns("CohortID").visible = False
                ctl.Columns("StudyID").visible = False
                ctl.Columns("Generated").readonly = True
                ctl.Columns("Generated").HeaderText = "Schedule Generated"
                ctl.columns("NumVols").HeaderText = "Number of volunteers"
                ctl.columns("CohortName").HeaderText = "Cohort Name"
                Dim cmb As New DataGridViewImageColumn
                cmb.HeaderText = "Add Volunteer"
                cmb.Image = My.Resources.Plus
                cmb.ImageLayout = DataGridViewImageCellLayout.Zoom
                ctl.columns.add(cmb)
                cmb.Name = "AddVolButton"

                Dim clm As New DataGridViewImageColumn
                clm.HeaderText = "Schedule/Re-Schedule"
                clm.Image = My.Resources.Calendar_icon
                clm.ImageLayout = DataGridViewImageCellLayout.Zoom
                clm.Name = "GenerateSchedule"
                ctl.columns.add(clm)


                Dim cmb2 As New DataGridViewImageColumn
                cmb2.HeaderText = "Delete Cohort"
                cmb2.Image = My.Resources.Remove
                cmb2.ImageLayout = DataGridViewImageCellLayout.Zoom
                ctl.columns.add(cmb2)
                cmb2.Name = "DeleteButton"

            Case "DataGridView8"

                OverClass.ResetCollection()

                SQLCode = "SELECT * FROM AlterOffset ORDER BY CalcDate ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)

                With DataGridView8
                    .Columns("VolunteerScheduleID").Visible = False
                    .Columns("CalcDate").Visible = False
                    .Columns("CohortID").Visible = False
                    .Columns("StudyID").Visible = False
                    .Columns("StudyCode").Visible = False
                    .Columns("VolID").Visible = False
                    .Columns("ProcName").ReadOnly = True
                    .Columns("ProcName").HeaderText = "Procedure"
                    .Columns("ProcOffSet").HeaderText = "Offset in Minutes"
                End With

                Dim clm As New DataGridViewTextBoxColumn
                clm.HeaderText = "Calculated Time"
                clm.ReadOnly = True
                DataGridView8.Columns.Add(clm)
                clm.Name = "ReCalc"
                clm.DefaultCellStyle.Format = "dd-MMM-yyyy HH:mm"

                DataGridView8.Columns("ProcOffSet").DisplayIndex = DataGridView8.Columns.Count - 1

                FilterCombo11.SetAsInternalSource("ProcName", "ProcName", OverClass)

                FilterCombo13.AllowBlanks = False
                FilterCombo13.SetAsExternalSource("StudyID", "StudyCode",
                "SELECT b.StudyID, StudyCode FROM Study a INNER JOIN Cohort b ON a.StudyID=b.StudyID", OverClass)

                FilterCombo24.AllowBlanks = False
                FilterCombo24.SetAsExternalSource("CohortID", "CohortName",
                "SELECT CohortID, CohortName FROM Cohort WHERE StudyID=" & FilterCombo24.SetCmbPointer(FilterCombo13), OverClass)

                FilterCombo25.AllowBlanks = False
                FilterCombo25.SetAsExternalSource("VolID", "Vol",
                "SELECT VolID, RVLNo & ' ' & Initials AS Vol FROM Volunteer " &
                "WHERE CohortID=" & FilterCombo25.SetCmbPointer(FilterCombo24), OverClass)

                DataGridView8.Columns("VolunteerScheduleID").Visible = False


            Case "DataGridView9"

                SQLCode = "Select CohortID, a.VolID, RVLNo, Initials, RoomNo, min(TimepointDateTime) As FirstDate " &
                    "FROM (Volunteer a INNER JOIN VolunteerTimepoint b On a.VolID=b.VolID) " &
                    "GROUP BY CohortID, a.VolID, RVLNo, Initials, RoomNo " &
                    "ORDER BY min(TimepointDateTime) ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.AllowUserToAddRows = False
                ctl.columns("VolID").visible = False
                ctl.columns("FirstDate").visible = False
                ctl.columns("CohortID").visible = False
                ctl.columns("RVLNo").HeaderText = "RVL Number"
                ctl.columns("RoomNo").HeaderText = "Room Number"
                Dim cmb2 As New DataGridViewImageColumn
                cmb2.HeaderText = "View Timepoints"
                cmb2.Image = My.Resources.Preview
                cmb2.ImageLayout = DataGridViewImageCellLayout.Zoom
                ctl.columns.add(cmb2)
                cmb2.Name = "Timepoints"
                Dim cmb As New DataGridViewImageColumn
                cmb.HeaderText = "Delete Volunteer"
                cmb.Image = My.Resources.Remove
                cmb.ImageLayout = DataGridViewImageCellLayout.Zoom
                ctl.columns.add(cmb)
                cmb.Name = "DeleteButton"

                FilterCombo14.AllowBlanks = False
                FilterCombo14.SetUpFilter(False, Nothing)
                FilterCombo14.SetAsExternalSource("StudyID", "StudyCode",
                "Select b.StudyID, StudyCode FROM Study a INNER JOIN Cohort b On a.StudyID= b.StudyID", OverClass)

                FilterCombo15.AllowBlanks = False
                FilterCombo15.SetAsExternalSource("CohortID", "CohortName",
                "Select CohortID, CohortName FROM Cohort WHERE StudyID=" & FilterCombo15.SetCmbPointer(FilterCombo14), OverClass)

                ctl.columns("CohortID").visible = False

            Case "DataGridView10"

                OverClass.ResetCollection()

                SQLCode = "Select CohortID, RVLNo, a.VolID, StudyID, VolunteerTimepointID, TimepointName, TimepointDateTime, DayNumber " &
                    "FROM (VolunteerTimepoint a INNER JOIN StudyTimepoint b " &
                    "On a.StudyTimepointID=b.StudyTimepointID) INNER JOIN Volunteer c On a.VolID=c.VolID " &
                    "ORDER BY TimepointDateTime ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.AllowUserToAddRows = False
                ctl.columns("VolunteerTimepointID").visible = False
                ctl.columns("TimepointName").Readonly = True
                ctl.columns("RVLNo").Readonly = True
                ctl.columns("StudyID").visible = False
                ctl.columns("VolID").visible = False
                ctl.columns("CohortID").visible = False
                ctl.columns("TimepointName").HeaderText = "Timepoint Name"
                ctl.columns("DayNumber").HeaderText = "Day Number"
                ctl.columns("TimepointDateTime").HeaderText = "Date/Time"
                ctl.columns("TimepointDateTime").DefaultCellStyle.Format = "dd-MMM-yyyy HH:mm"

                FilterCombo16.AllowBlanks = False
                FilterCombo16.SetAsExternalSource("StudyID", "StudyCode",
                "SELECT b.StudyID, StudyCode FROM Study a INNER JOIN Cohort b ON a.StudyID=b.StudyID", OverClass)

                FilterCombo17.AllowBlanks = False
                FilterCombo17.SetAsExternalSource("CohortID", "CohortName",
                "SELECT CohortID, CohortName FROM Cohort WHERE StudyID=" & FilterCombo17.SetCmbPointer(FilterCombo16), OverClass)

                FilterCombo18.AllowBlanks = False
                FilterCombo18.SetAsExternalSource("VolID", "Vol",
                "SELECT VolID, RVLNo & ' ' & Initials AS Vol FROM Volunteer " &
                "WHERE CohortID=" & FilterCombo18.SetCmbPointer(FilterCombo17), OverClass)


                ctl.columns("VolID").visible = False

            Case "DataGridView11"

                OverClass.ResetCollection()

                SQLCode = "Select * FROM Assign " &
                            "ORDER BY CalcDate ASC, ProcOrd ASC"

                If Me.CheckBox1.Checked = True Then
                    SQLCode = "Select * FROM Assign WHERE VolunteerScheduleID In " &
                        "(Select ID FROM " &
                        "(Select First(VolunteerScheduleID) As ID,  Min(CalcDate) " &
                        "FROM(" & SQLCode & ")" &
                        "GROUP BY ProcName, VOL)) " &
                        "ORDER BY CalcDate ASC, ProcOrd ASC"
                End If



                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.AllowUserToAddRows = False
                ctl.columns("VolunteerScheduleID").visible = False
                ctl.columns("StaffID").visible = False
                ctl.columns("CohortID").visible = False
                ctl.columns("ProcOrd").visible = False
                ctl.columns("FullName").visible = False
                ctl.columns("CalcDate").visible = False
                ctl.columns("EndFull").visible = False
                ctl.columns("CohortName").visible = False
                ctl.columns("StudyCode").visible = False
                ctl.columns("WhichDay").visible = False
                ctl.columns("Approx").visible = False
                ctl.columns("Vol").readonly = True
                ctl.columns("Approx").readonly = True
                ctl.columns("ProcName").readonly = True
                ctl.columns("CalcDate").readonly = True
                ctl.columns("DispTime").readonly = True
                ctl.columns("DispStudy").readonly = True
                ctl.columns("DispTime").HeaderText = "Date/Time"
                ctl.columns("DispStudy").HeaderText = "Study/Cohort"
                ctl.columns("Approx").HeaderText = "Timepoint"
                ctl.columns("ProcName").HeaderText = "Procedure"
                ctl.columns("StudyCode").HeaderText = "Study Code"
                ctl.columns("CohortName").HeaderText = "Cohort"
                ctl.columns("Vol").HeaderText = "Volunteer"
                ctl.columns("CalcDate").DefaultCellStyle.Format = "dd-MMM-yyyy HH:mm"
                ctl.columns("EndFull").DefaultCellStyle.Format = "HH:mm"
                Dim cmb As New DataGridViewComboBoxColumn
                cmb.DataSource = OverClass.TempDataTable("SELECT StaffID, FName & ' ' & SName AS Fullname, StaffID " &
                                                         "FROM STAFF WHERE Hidden=False ORDER BY FName ASC")
                ctl.columns.add(cmb)
                cmb.HeaderText = "Staff Member"
                cmb.ValueMember = "StaffID"
                cmb.DisplayMember = "Fullname"
                cmb.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("StaffID").ToString
                cmb.Name = "PICK"
                Dim cmb2 As New DataGridViewImageColumn
                cmb2.HeaderText = "Delete Procedure"
                cmb2.Image = My.Resources.Remove
                cmb2.ImageLayout = DataGridViewImageCellLayout.Zoom
                ctl.columns.add(cmb2)
                cmb2.Name = "DeleteButton"
                cmb2.Width = 60

                FilterCombo5.SetAsInternalSource("StudyCode", "StudyCode", OverClass)
                FilterCombo10.SetAsInternalSource("CohortName", "CohortName", OverClass)
                FilterCombo7.SetAsInternalSource("ProcName", "ProcName", OverClass)
                FilterCombo9.SetAsInternalSource("Vol", "Vol", OverClass)
                FilterCombo8.SetAsInternalSource("WhichDay", "WhichDay", OverClass)
                FilterCombo6.SetAsInternalSource("FullName", "FullName", OverClass)



            Case "DataGridView12"

                SQLCode = "SELECT StaffProcID, StaffID, ProcID, ProcDateTime " &
                    "FROM StaffProc " &
                    "WHERE ProcDateTime > Now() " &
                    " ORDER BY ProcDateTime ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)

                ctl.columns("StaffProcID").visible = False
                ctl.columns("StaffID").visible = False
                ctl.columns("ProcID").visible = False
                ctl.columns("ProcDateTime").HeaderText = "Date/Time"
                ctl.columns("ProcDateTime").DefaultCellStyle.Format = "dd-MMM-yyyy HH:mm"
                Dim cmb As New DataGridViewComboBoxColumn
                cmb.DataSource = OverClass.TempDataTable("SELECT StaffID, FName & ' ' & SName AS Fullname " &
                                                         "FROM STAFF WHERE Hidden=False ORDER BY FName ASC")
                ctl.columns.add(cmb)
                cmb.Name = "Pick"
                cmb.HeaderText = "Staff Member"
                cmb.ValueMember = "StaffID"
                cmb.DisplayMember = "Fullname"
                cmb.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("StaffID").ToString
                Dim cmb2 As New DataGridViewComboBoxColumn
                cmb2.DataSource = OverClass.TempDataTable("SELECT ProcID, ProcName " &
                                                         "FROM ProcTask ORDER BY ProcName ASC")
                ctl.columns.add(cmb2)
                cmb2.HeaderText = "Procedure"
                cmb2.ValueMember = "ProcID"
                cmb2.DisplayMember = "ProcName"
                cmb2.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("ProcID").ToString
                Dim cmb3 As New DataGridViewComboBoxColumn
                cmb3.DataSource = OverClass.TempDataTable("SELECT ProcID, MinsTaken " &
                                                          "FROM ProcTask")
                cmb3.ValueMember = "ProcID"
                cmb3.DisplayMember = "MinsTaken"
                cmb3.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("ProcID").ToString
                cmb3.Name = "MinsTaken"
                cmb2.Name = "ProcPick"
                ctl.columns.add(cmb3)
                ctl.columns("ProcDateTime").name = "CalcDate"
                cmb3.Visible = False
                Dim cmb4 As New DataGridViewImageColumn
                cmb4.HeaderText = "Delete Procedure"
                cmb4.Image = My.Resources.Remove
                cmb4.ImageLayout = DataGridViewImageCellLayout.Zoom
                ctl.columns.add(cmb4)
                cmb4.Name = "DeleteButton"
                cmb4.Width = 60

                FilterCombo4.SetAsExternalSource("ProcID", "ProcName",
                           "SELECT ProcID, ProcName FROM ProcTask " &
                          "WHERE ProcID IN (" &
                FilterCombo4.SetCmbPointer(OverClass.CurrentDataSet.Tables(0).Columns("ProcID")) & ")", OverClass)

                FilterCombo3.SetAsExternalSource("StaffID", "FullName",
                           "SELECT StaffID, FName & ' ' & SName AS FullName FROM Staff " &
                          "WHERE StaffID IN (" &
                FilterCombo3.SetCmbPointer(OverClass.CurrentDataSet.Tables(0).Columns("StaffID")) & ")", OverClass)

            Case "DataGridView13"
                SQLCode = "SELECT * FROM ReportArchive ORDER BY ArchiveID DESC, ArchiveDate DESC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.columns("ArchiveID").Visible = False
                ctl.columns("ArchivePath").visible = False
                ctl.Columns("ArchiveType").displayindex = 0
                ctl.columns("ArchiveType").HeaderText = "Report Type"
                ctl.columns("ArchiveDate").HeaderText = "Date Ran"
                ctl.columns("ArchiveDate").DefaultCellStyle.Format = "dd-MMM-yyyy HH:mm"
                ctl.columns("ArchiveUser").HeaderText = "User Ran"
                ctl.columns("ArchiveCriteria").HeaderText = "Report Criteria"
                ctl.Columns("ArchiveCriteria").DefaultCellStyle.WrapMode = DataGridViewTriState.True
                ctl.Columns("ArchiveType").DefaultCellStyle.Font = New Font("Arial", 10, FontStyle.Underline)
                ctl.Columns("ArchiveType").DefaultCellStyle.ForeColor = Color.Blue

            Case "FilterCombo12"

                FilterCombo12.AllowBlanks = False
                FilterCombo12.Filter = False
                FilterCombo12.SetAsExternalSource("StudyID", "StudyCode", "SELECT StudyCode, StudyID FROM Study", OverClass)

                FilterCombo26.Filter = False
                FilterCombo26.SetAsExternalSource("CohortID", "CohortName", "SELECT CohortID, CohortName FROM Cohort WHERE StudyID=" &
                                                  FilterCombo26.SetCmbPointer(FilterCombo12), OverClass)

                FilterCombo27.Filter = False
                FilterCombo27.SetAsExternalSource("VolID", "Vol",
                "SELECT VolID, RVLNo & ' ' & Initials AS Vol FROM Volunteer " &
                "WHERE CohortID=" & FilterCombo27.SetCmbPointer(FilterCombo26), OverClass)

                FilterCombo28.Filter = False
                FilterCombo28.SetAsExternalSource("ProcID", "ProcName", "SELECT ProcTask.ProcID, ProcTask.ProcName, Study.StudyID " &
                "FROM (Study INNER JOIN StudyTimepoint On Study.StudyID = StudyTimepoint.StudyID) INNER JOIN " &
                "(ProcTask INNER JOIN StudySchedule On ProcTask.ProcID = StudySchedule.ProcID) On StudyTimepoint.StudyTimepointID = StudySchedule.StudyTimepointID " &
                " WHERE Study.StudyID=" & FilterCombo28.SetCmbPointer(FilterCombo12), OverClass)

                FilterCombo29.Filter = False
                FilterCombo29.SetAsExternalSource("DaysPost", "DaysPost", "SELECT DaysPost " &
                "FROM StudyTimepoint INNER JOIN StudySchedule ON StudyTimepoint.StudyTimepointID = StudySchedule.StudyTimepointID " &
                "WHERE StudyID=" & FilterCombo29.SetCmbPointer(FilterCombo12), OverClass)

        End Select

    End Sub

    Private Sub TabControl2_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl2.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If OverClass.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        OverClass.ResetCollection()

        Select Case e.TabPage.Text

            Case "Procedures"
                ctl = Me.DataGridView1
                SQLCode = "Select ProcID, ProcName, MinsTaken, ProcOrd FROM ProcTask ORDER BY ProcName ASC"
                OverClass.CreateDataSet(SQLCode, Bind, ctl)

            Case "Staff"
                ctl = Me.DataGridView2
                SQLCode = "Select StaffID, FName, SName, Hidden FROM Staff ORDER BY Hidden DESC, FName ASC"
                OverClass.CreateDataSet(SQLCode, Bind, ctl)

        End Select


        If Not IsNothing(ctl) Then Call Specifics(ctl)

    End Sub

    Private Sub TabControl3_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl3.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If OverClass.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        OverClass.ResetCollection()

        Select Case e.TabPage.Text

            Case "StudyCodes"
                ctl = Me.DataGridView3
                SQLCode = "Select StudyID, StudyCode, Colour FROM Study ORDER BY StudyCode ASC"
                OverClass.CreateDataSet(SQLCode, Bind, ctl)

            Case "Timepoints"
                Call Specifics(DataGridView5)

            Case "Schedules"
                Call Specifics(DataGridView6)

            Case "Cohorts"
                Call Specifics(DataGridView7)

        End Select

    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick
        Dim senderGrid = DirectCast(sender, DataGridView)

        If TypeOf senderGrid.Columns(e.ColumnIndex) Is DataGridViewButtonColumn AndAlso
           e.RowIndex >= 0 Then
            Dim cDialog As New ColorDialog()
            If (cDialog.ShowDialog() = DialogResult.OK) Then
                Me.DataGridView3.Rows(e.RowIndex).Cells("Colour").Value = ColorTranslator.ToHtml(cDialog.Color)
                Me.DataGridView3.Rows(e.RowIndex).Cells("Colour").Style.BackColor = cDialog.Color
            End If
        End If

    End Sub

    Private Sub DataGridView3_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DataGridView3.RowPrePaint
        For Each row In Me.DataGridView3.Rows
            If Not IsNothing(row.Cells("Colour").value) And Not IsDBNull(row.Cells("Colour").value) Then
                Dim colourString As String = row.Cells("Colour").value
                row.Cells("Colour").Style.BackColor = ColorTranslator.FromHtml(colourString)
            End If
        Next
    End Sub

    Private Sub TabControl4_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl4.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If OverClass.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        OverClass.ResetCollection()

        Select Case e.TabPage.Text

            Case "Details"
                Specifics(DataGridView9)

            Case "Timepoints"
                Specifics(DataGridView10)

        End Select

    End Sub

    Private Sub TabControl5_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl5.Selecting


        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If OverClass.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        OverClass.ResetCollection()

        Select Case e.TabPage.Text

            Case "Volunteer Assign"
                Call Specifics(Me.DataGridView11)

            Case "Staff Procedures"
                Call Specifics(Me.DataGridView12)

        End Select

    End Sub

    Private Sub DataGridView11_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView11.CellEndEdit

        Dim returner As String = vbNullString

        If IsDBNull(sender.rows(e.RowIndex).cells("StaffID").value) Or
            IsNothing(sender.rows(e.RowIndex).cells("StaffID").value) Then Exit Sub
        If e.ColumnIndex <> sender.Columns("Pick").Index Then Exit Sub


        returner = CheckVolunteerOverlap(sender.rows(e.RowIndex).cells("StaffID").value, sender.rows(e.RowIndex).cells("VolunteerScheduleID").value,
                              sender.rows(e.RowIndex).cells("CalcDate").value, sender.rows(e.RowIndex).cells("EndFull").value, sender)

        If returner <> vbNullString Then
            MsgBox("Overlap found - " & vbNewLine & vbNewLine & returner)
        End If

    End Sub

    Private Sub DataGridView7_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView7.CellBeginEdit

        If IsDBNull(sender.item("CohortID", e.RowIndex).value) Or IsNothing(sender.item("CohortID", e.RowIndex).value) Then Exit Sub
        If sender.item("Generated", e.RowIndex).value = True And e.ColumnIndex = sender.columns("NumVols").index Then e.Cancel = True

    End Sub

    Private Sub DataGridView7_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView7.CellContentClick

        If IsDBNull(sender.item("CohortID", e.RowIndex).value) Then Exit Sub


        If e.ColumnIndex = sender.columns("GenerateSchedule").index Then
            If sender.item("Generated", e.RowIndex).value = False Then
                Call Generator(sender.item("CohortID", e.RowIndex).value)
            Else
                Call ReSchedule(sender.item("CohortID", e.RowIndex).value, sender.item("StudyID", e.RowIndex).value,
                                sender.item("CohortName", e.RowIndex).value, sender.item("NumVols", e.RowIndex).value)
            End If
        End If

        If e.ColumnIndex = sender.columns("DeleteButton").index Then
            Dim row As DataGridViewRow
            row = sender.rows(e.RowIndex)
            sender.rows.remove(row)
        End If


        If e.ColumnIndex = sender.columns("AddVolButton").index Then

            Me.ValidateChildren()
            If OverClass.UnloadData() = True Then Exit Sub

            Dim Response As Integer = MsgBox("Do you want To transfer the volunteer from another study?" _
                                             & vbNewLine & vbNewLine _
                                             & "Timepoints Of EXACTLY the same name will be transfered across", MsgBoxStyle.YesNoCancel)
            If Response = vbCancel Then Exit Sub



            'EXISTING VOLUNTEER
            If Response = vbYes Then

                PickCohort = sender.item("CohortID", e.RowIndex).value
                Dim PickVol As New AddVol

                PickVol.ShowDialog()

                'ON RETURN
                Me.TabControl3.SelectedIndex = 3
                Me.TabControl3_Selecting(Me.TabControl3, New TabControlCancelEventArgs(TabPage5, 0, False, TabControlAction.Selecting))
                Call Me.Specifics(Me.DataGridView7)



            End If


            'NEW VOLUNTEER
            If Response = vbNo Then

                Dim RVLNo, RoomNo, VolID, CohortID As Long
                Dim Initials = vbNullString
                Dim Temp As String = vbNullString
                Dim Accepted As Boolean

                CohortID = sender.item("CohortID", e.RowIndex).value

                'GET NEW VOL INFO
                Do While Accepted = False
                    Initials = InputBox("Input Volunteer Initials", "Volunteer Initials", "AAA")
                    If Initials = "" Then Exit Sub
                    If Len(Initials) <> 3 Then
                        MsgBox("Initials must be 3 characters Long")
                        Continue Do
                    End If
                    If Not Initials Like "[A-Z][A-Z][A-Z]" Then
                        MsgBox("Initials must be 3 text characters such As 'AAA'")
                Continue Do
                    End If
                    Accepted = True
                Loop

                Accepted = False
                Do While Accepted = False
                    Temp = InputBox("Input Volunteer RVL Number", "Volunteer RVL No", "123456")
                    If Temp = "" Then Exit Sub
                    Try
                        RVLNo = CLng(Temp)
                    Catch ex As Exception
                        MsgBox("RVL Number must be a number")
                        Continue Do
                    End Try
                    Accepted = True
                Loop

                Accepted = False
                Do While Accepted = False
                    Temp = InputBox("Input Volunteer Room Number", "Volunteer Room No", "10")
                    If Temp = "" Then Exit Sub
                    Try
                        RoomNo = CLng(Temp)
                    Catch ex As Exception
                        MsgBox("Room Number must be a number")
                        Continue Do
                    End Try
                    Accepted = True
                Loop


                'GET A NEW VOL ID
                Try
                    VolID = OverClass.TempDataTable("SELECT Max(VolID) FROM Volunteer").Rows(0).Item(0) + 1

                Catch ex As Exception
                    VolID = 1

                End Try


                'TRY AND INSERT VOLUNTEER
                Try
                    Dim InsertString As String
                    Dim cmdInsert As OleDb.OleDbCommand

                    InsertString = "INSERT INTO Volunteer " &
                                         "(VolID, RVLNo, Initials, CohortID, RoomNo) " &
                                     "VALUES (" & VolID & ", " & RVLNo & ", '" & Initials & "', " & CohortID & ", " & RoomNo & ")"


                    cmdInsert = New OleDb.OleDbCommand(InsertString)

                    OverClass.ExecuteSQL(cmdInsert)

                Catch ex As Exception
                    MsgBox(ex.Message)
                    Exit Sub

                End Try



                For Each row In OverClass.TempDataTable("SELECT TimepointName, StudyTimepointID FROM ((StudyTimepoint a " &
                                                        "INNER JOIN Study b ON a.StudyID=b.StudyID) " &
                                                        "INNER JOIN Cohort c ON b.StudyID=C.StudyID) " &
                                                        "WHERE CohortID=" & CohortID &
                                                        " GROUP BY TimepointName, StudyTimepointID").Rows
                    Accepted = False
                    Dim TempDate As Date
                    Dim InsertString, TimepointName As String
                    Dim cmdInsert As OleDb.OleDbCommand
                    Dim StudyTimepointID As Long

                    TimepointName = row.item("TimepointName")
                    StudyTimepointID = row.item("StudyTimepointID")

                    'INSERT INTO VOLUNTEER TIMEPOINT
                    Do While Accepted = False

                        Temp = InputBox("Input " & Initials & "(" & RVLNo & ") " & TimepointName & " Date/Time",
                                        TimepointName & " Date", "01-Jan-2010 10:00")

                        Try
                            TempDate = CDate(Temp)
                            If Format(TempDate, "HH:mm") = "00:00" Then Throw New System.Exception
                            InsertString = "INSERT INTO VolunteerTimepoint " &
                                         "(VolID, TimepointDateTime, StudyTimepointID) " &
                                     "VALUES (" & VolID & ", " & OverClass.SQLDate(TempDate) & ", " & StudyTimepointID & ")"


                            cmdInsert = New OleDb.OleDbCommand(InsertString)

                            OverClass.AddToMassSQL(cmdInsert)


                        Catch ex As Exception
                            MsgBox("Must enter a valid Date/Time to continue")
                            Continue Do

                        End Try

                        'INSERT ALL PROCEDURES
                        For Each row2 In OverClass.TempDataTable("SELECT StudyScheduleID FROM StudySchedule " &
                                                                "WHERE StudyTimepointID=" & StudyTimepointID).Rows

                            Dim CmdInsert2 As New OleDb.OleDbCommand("INSERT INTO VolunteerSchedule (StudyScheduleID,VolID) " &
                                                                     "VALUES (" & row2.item(0) & ", " & VolID & ")")

                            OverClass.AddToMassSQL(CmdInsert2)

                        Next


                        Accepted = True
                        Continue For

                    Loop

                Next

                OverClass.ExecuteMassSQL()




                'UPDATE COHORT TO GENERATED
                Dim UpdateString As String
                Dim cmdUpdate As OleDb.OleDbCommand

                UpdateString = "UPDATE Cohort SET Generated=TRUE, NumVols=NumVols+1 " &
                    "WHERE CohortID=" & CohortID & " AND Generated=TRUE"

                cmdUpdate = New OleDb.OleDbCommand(UpdateString)

                OverClass.ExecuteSQL(cmdUpdate)

                UpdateString = "UPDATE Cohort SET Generated=TRUE, NumVols=1 " &
                    "WHERE CohortID=" & CohortID & " AND Generated=False"

                cmdUpdate = New OleDb.OleDbCommand(UpdateString)

                OverClass.ExecuteSQL(cmdUpdate)

                MsgBox("Volunteer Added")

                'REFRESH SCREEN
                Me.TabControl3.SelectedIndex = 3
                Me.TabControl3_Selecting(Me.TabControl3, New TabControlCancelEventArgs(TabPage5, 0, False, TabControlAction.Selecting))
                Call Specifics(Me.DataGridView7)

            End If
        End If



    End Sub

    Private Sub DataGridView9_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView9.CellContentClick

        If e.ColumnIndex = sender.columns("DeleteButton").index Then
            If MsgBox("Are you sure you want to delete?" & vbNewLine & "Table must be saved to commit delete", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Dim row As DataGridViewRow
                row = sender.rows(e.RowIndex)
                sender.rows.remove(row)
            End If
        End If

        If e.ColumnIndex = sender.columns("Timepoints").index Then
            Dim dt = OverClass.TempDataTable("SELECT TimepointName, TimepointDateTime " &
                                             "FROM VolunteerTimepoint a INNER JOIN StudyTimepoint b ON a.StudyTimepointID=b.StudyTimepointID " &
                                             "WHERE VolID=" & sender.item(sender.Columns("VolID").Index, e.RowIndex).value &
                                             " ORDER BY TimepointDateTime ASC")

            Dim msg As String = vbNullString

            For Each row In dt.Rows
                msg = row.Item("TimepointName").ToString & " - " & row.Item("TimepointDateTime").ToString
                msg = msg & vbNewLine
            Next

            MsgBox(msg)

        End If


    End Sub

    Private Sub DataGridView11_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView11.CellContentClick

        If e.ColumnIndex <> sender.columns("DeleteButton").index Then Exit Sub

        If MsgBox("Are you sure you want to delete?" & vbNewLine & "Table must be saved to commit delete", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim row As DataGridViewRow
            row = sender.rows(e.RowIndex)
            sender.rows.remove(row)
        End If

    End Sub

    Private Sub DataGridView12_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView12.CellContentClick

        If e.ColumnIndex <> sender.columns("DeleteButton").index Then Exit Sub

        If MsgBox("Are you sure you want to delete?" & vbNewLine & "Table must be saved to commit delete", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim row As DataGridViewRow
            row = sender.rows(e.RowIndex)
            sender.rows.remove(row)
        End If

    End Sub

    Private Sub DataGridView13_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView13.CellContentClick

        If e.ColumnIndex = sender.columns("ArchiveType").index Then

            Dim FilePath As String = sender.item("ArchivePath", e.RowIndex).value

            Try

                Process.Start("explorer.exe", FilePath)


            Catch ex As Exception
                MsgBox(ex.Message)

            End Try

        End If

    End Sub

    Private Sub DataGridView6_CellEndEdit_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView6.CellEndEdit
        Dim Returner As String = vbNullString

        Dim DaysPost As Long
        Dim ProcTime As Date
        Dim MinsTaken As Long

        Try
            DaysPost = CLng(sender.rows(e.RowIndex).cells("Dayspost").value)
        Catch ex As Exception
            Exit Sub
        End Try
        Try
            ProcTime = CDate(sender.rows(e.RowIndex).cells("ProcTime").value)
        Catch ex As Exception
            Exit Sub
        End Try
        Try
            MinsTaken = CLng(sender.rows(e.RowIndex).cells("MinsTaken").formattedvalue)
        Catch ex As Exception
            Exit Sub
        End Try

        Returner = ScheduleOverlap(sender, e.RowIndex, sender.rows(e.RowIndex).cells("Dayspost").value,
                              sender.rows(e.RowIndex).cells("ProcTime").value,
                              sender.rows(e.RowIndex).cells("MinsTaken").formattedvalue)

        If Returner <> vbNullString Then MsgBox("Overlap found - " & vbNewLine & vbNewLine & Returner)


    End Sub

    Private Sub DataGridView6_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView6.CellContentClick

        If e.ColumnIndex <> sender.columns("DeleteButton").index Then Exit Sub

        If MsgBox("Are you sure you want To delete?" & vbNewLine & "Table must be saved To commit delete", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim row As DataGridViewRow
            row = sender.rows(e.RowIndex)
            sender.rows.remove(row)
        End If

    End Sub

    Private Sub DataGridView12_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView12.CellEndEdit

        Dim returner As String = vbNullString

        Dim StaffID As Long = 0
        Dim ProcID As Long = 0
        Dim CalcDate As Date = "01/01/2000"
        Dim Identifier As Long = 0
        Try
            StaffID = sender.rows(e.RowIndex).cells("StaffID").value
        Catch ex As Exception
        End Try
        Try
            ProcID = sender.rows(e.RowIndex).cells("ProcID").value
        Catch ex As Exception
        End Try
        Try
            CalcDate = sender.rows(e.RowIndex).cells("CalcDate").value
        Catch ex As Exception
        End Try
        Try
            Identifier = sender.rows(e.RowIndex).cells("StaffProcID").value
        Catch ex As Exception
        End Try

        If StaffID = 0 Or ProcID = 0 Or CalcDate = "01/01/2000" Then Exit Sub

        returner = CheckExtraOverlap(StaffID, Identifier,
                        CalcDate,
                        DateAdd(DateInterval.Minute, sender.rows(e.RowIndex).cells("MinsTaken").FormattedValue, CalcDate),
                        sender, e.RowIndex)

        If returner <> vbNullString Then
            MsgBox("Overlap found - " & vbNewLine & vbNewLine & returner)
        End If

    End Sub

    Private Sub CheckBox1_Click(sender As Object, e As EventArgs) Handles CheckBox1.Click
        If OverClass.UnloadData() = True Then
            RemoveHandler CheckBox1.Click, AddressOf CheckBox1_Click
            CheckBox1.Checked = Not CheckBox1.Checked
            AddHandler CheckBox1.Click, AddressOf CheckBox1_Click
            Exit Sub
        End If
        Call Specifics(DataGridView11)
    End Sub

    Private Sub FilterCombo21_SelectedIndexChanged(sender As Object, e As EventArgs) Handles FilterCombo21.SelectedIndexChanged

        Try
            Me.TextBox1.Clear()
            Dim dt As DataTable = OverClass.TempDataTable("SELECT DefaultTime FROM StudyTimepoint " &
                                                     "WHERE StudyTimepointID=" & Me.FilterCombo21.SelectedValue.ToString)

            If Not IsDBNull(dt.Rows(0).Item(0)) Then Me.TextBox1.Text = Format(dt.Rows(0).Item(0), "HH:mm")
        Catch ex As Exception
        End Try

    End Sub

    Private Sub TabControl6_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl6.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If OverClass.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        OverClass.ResetCollection()

        Select Case e.TabPage.Text

            Case "Bulk"
                ctl = FilterCombo12


            Case "Individual"
                ctl = Me.DataGridView8
                SQLCode = "SELECT StaffID, FName, SName, Hidden FROM Staff ORDER BY Hidden DESC, FName ASC"
                OverClass.CreateDataSet(SQLCode, Bind, ctl)

        End Select


        If Not IsNothing(ctl) Then Call Specifics(ctl)

    End Sub

    Private Sub DataGridView8_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DataGridView8.RowPrePaint

        On Error Resume Next

        Dim Offset As Long = DataGridView8.Item("ProcOffSet", e.RowIndex).Value
        Dim CalcDate As Date = DataGridView8.Item("CalcDate", e.RowIndex).Value
        DataGridView8.Item("ReCalc", e.RowIndex).Value = DateAdd(DateInterval.Minute, Offset, CalcDate)

    End Sub

End Class

