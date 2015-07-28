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

        Select Case e.TabPageIndex

            Case 1
                Me.TabControl2.SelectedIndex = 0
                Me.TabControl2_Selecting(Me.TabControl2, New TabControlCancelEventArgs(TabPage3, 0, False, TabControlAction.Selecting))
            Case 2
                Me.TabControl3.SelectedIndex = 0
                Me.TabControl3_Selecting(Me.TabControl3, New TabControlCancelEventArgs(TabPage5, 0, False, TabControlAction.Selecting))
            Case 3
                Me.TabControl4.SelectedIndex = 0
                Me.TabControl4_Selecting(Me.TabControl4, New TabControlCancelEventArgs(TabPage15, 0, False, TabControlAction.Selecting))
            Case 4
                Me.TabControl5.SelectedIndex = 0
                Me.TabControl5_Selecting(Me.TabControl5, New TabControlCancelEventArgs(TabPage18, 0, False, TabControlAction.Selecting))


        End Select


        Call Specifics(ctl)

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
                SQLCode = "SELECT StudyTimepointID, TimepointName FROM StudyTimepoint WHERE StudyID=" _
                    & Me.ComboBox2.SelectedValue.ToString & " ORDER BY TimepointName ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.Columns(0).visible = False
                ctl.columns(1).headertext = "Timepoint Name"

            Case "DataGridView6"
                SQLCode = "SELECT StudyScheduleID, ProcID, DaysPost, HoursPost, MinsPost, Approx, SetTime" & _
                    " FROM StudySchedule WHERE StudyTimepointID=" & Me.ComboBox3.SelectedValue.ToString & _
                    " ORDER BY DaysPost ASC, HoursPost ASC, MinsPost ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.Columns(0).visible = False
                ctl.columns(1).visible = False
                Dim cmb As New DataGridViewComboBoxColumn
                cmb.HeaderText = "Procedure"
                cmb.DataSource = OverClass.TempDataTable("SELECT ProcID, ProcName" & _
                                                     " FROM ProcTask ORDER BY ProcName ASC")
                cmb.ValueMember = "ProcID"
                cmb.DisplayMember = "ProcName"
                cmb.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("ProcID").ToString
                ctl.columns.add(cmb)
                cmb.DisplayIndex = 1
                ctl.columns(2).headertext = "Days"
                ctl.columns(3).headertext = "Hours"
                ctl.columns(4).headertext = "Minutes"
                ctl.columns(6).headertext = "Set Time"
                ctl.Columns("SetTime").DefaultCellStyle.Format = "HH:mm"
                Dim cmb2 As New DataGridViewComboBoxColumn
                cmb2.HeaderText = "Timepoint"
                cmb2.DataSource = OverClass.TempDataTable("SELECT Display FROM (SELECT 'Approx' As Display " & _
                                                     " FROM StudyTimepoint WHERE StudyID=" & Me.ComboBox4.SelectedValue.ToString & _
                                                     " UNION ALL " & _
                                                     " SELECT 'Timed' As Display " & _
                                                     " FROM StudyTimepoint WHERE StudyID=" & Me.ComboBox4.SelectedValue.ToString & _
                                                     " UNION ALL " & _
                                                     " SELECT 'Set Time' AS Display FROM StudyTimepoint) " & _
                                                     "GROUP BY Display ORDER BY Display ASC")
                cmb2.ValueMember = "Display"
                cmb2.DisplayMember = "Display"
                cmb2.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("Approx").ToString
                ctl.columns.add(cmb2)
                ctl.columns("SetTime").DefaultCellStyle.Format = "HH:mm"
                ctl.columns("Approx").visible = False
                cmb2.DisplayIndex = 2

            Case "DataGridView7"
                SQLCode = "SELECT CohortID, CohortName, NumVols" & _
                    " FROM Cohort WHERE StudyID=" & Me.ComboBox5.SelectedValue.ToString & _
                    " ORDER BY CohortName ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.Columns("CohortID").visible = False
                ctl.columns("NumVols").HeaderText = "Number of volunteers"
                ctl.columns("CohortName").HeaderText = "Cohort Name"

            Case "DataGridView8"
                SQLCode = "SELECT CohortTimePointID, StudyTimepointID, VolGap, TimepointDateTime" & _
                    " FROM CohortTimepoint " & _
                    " WHERE CohortID=" & Me.ComboBox7.SelectedValue.ToString & _
                    " ORDER BY TimepointDateTime ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.Columns("CohortTimePointID").visible = False
                ctl.Columns("StudyTimePointID").visible = False
                ctl.columns("TimepointDateTime").HeaderText = "Date/Time"
                ctl.columns("VolGap").HeaderText = "Interval (Minutes)"
                Dim cmb As New DataGridViewComboBoxColumn
                cmb.HeaderText = "Timepoint"
                cmb.DataSource = OverClass.TempDataTable("SELECT StudyTimepointID, TimepointName " & _
                                                    "FROM StudyTimepoint " & _
                                                    "WHERE StudyID=" & Me.ComboBox6.SelectedValue.ToString)
                cmb.ValueMember = "StudyTimepointID"
                cmb.DisplayMember = "TimepointName"
                cmb.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("StudyTimepointID").ToString
                ctl.columns.add(cmb)
                cmb.DisplayIndex = 0
                ctl.columns("TimepointDateTime").DefaultCellStyle.Format = "dd-MMM-yyyy HH:mm"

            Case "DataGridView9"
                SQLCode = "SELECT VolID, RVLNo, Initials, RoomNo " & _
                    "FROM Volunteer " & _
                    "WHERE CohortID=" & Me.ComboBox10.SelectedValue.ToString & _
                    " ORDER BY Initials ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.AllowUserToAddRows = False
                ctl.columns("VolID").visible = False
                ctl.columns("RVLNo").HeaderText = "RVL Number"
                ctl.columns("RoomNo").HeaderText = "Room Number"

            Case "DataGridView10"
                SQLCode = "SELECT VolunteerTimepointID, TimepointName, TimepointDateTime " & _
                    "FROM VolunteerTimepoint a INNER JOIN StudyTimepoint b " & _
                    "ON a.StudyTimepointID=b.StudyTimepointID " & _
                    "WHERE VolID=" & Me.ComboBox13.SelectedValue.ToString & _
                    " ORDER BY TimepointDateTime ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.AllowUserToAddRows = False
                ctl.columns("VolunteerTimepointID").visible = False
                ctl.columns("TimepointName").Readonly = True
                ctl.columns("TimepointName").HeaderText = "Timepoint Name"
                ctl.columns("TimepointDateTime").HeaderText = "Date/Time"
                ctl.columns("TimepointDateTime").DefaultCellStyle.Format = "dd-MMM-yyyy HH:mm"

            Case "DataGridView11"
                SQLCode = "SELECT VolunteerScheduleID, 'Vol: ' & RVLNo & ' - ' & Initials & ' - Room (' & RoomNo & ')' AS VOL, " & _
                "StaffID, Approx, ProcName, " & _
                    "iif(Approx='Set Time',dateadd('d',DaysPost, TimepointDateTime), dateadd('n',Minspost, " & _
                    "dateadd('h',HoursPost,dateadd('d',DaysPost, TimepointDateTime)))) AS CalcDate " & _
                    "FROM (((StudySchedule a " & _
                    "INNER JOIN VolunteerSchedule c ON a.StudyScheduleID=c.StudyScheduleID) " & _
                    "INNER JOIN Volunteer d ON c.VolID=d.VolID) " & _
                    "INNER JOIN ProcTask e ON a.ProcID=e.ProcID) " & _
                    "INNER JOIN VolunteerTimepoint f ON d.VolID=f.VolID " & _
                    "AND a.StudyTimepointID=f.StudyTimepointID " & _
                    "WHERE IsNull(StaffID) " & _
                    "AND iif(Approx='Set Time',dateadd('d',DaysPost, TimepointDateTime), dateadd('n',Minspost, " & _
                    "dateadd('h',HoursPost,dateadd('d',DaysPost, TimepointDateTime)))) > Now() " & _
                    "ORDER BY iif(Approx='Set Time',dateadd('d',DaysPost, TimepointDateTime), dateadd('n',Minspost, " & _
                    "dateadd('h',HoursPost,dateadd('d',DaysPost, TimepointDateTime)))) ASC, RVLNo ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.AllowUserToAddRows = False
                ctl.columns("VolunteerScheduleID").visible = False
                ctl.columns("StaffID").visible = False
                ctl.columns("Vol").readonly = True
                ctl.columns("Approx").readonly = True
                ctl.columns("ProcName").readonly = True
                ctl.columns("CalcDate").readonly = True
                ctl.columns("CalcDate").HeaderText = "Date/Time"
                ctl.columns("Approx").HeaderText = "Timepoint"
                ctl.columns("ProcName").HeaderText = "Procedure"
                ctl.columns("Vol").HeaderText = "Volunteer"
                ctl.columns("CalcDate").DefaultCellStyle.Format = "dd-MMM-yyyy HH:mm"
                Dim cmb As New DataGridViewComboBoxColumn
                cmb.DataSource = OverClass.TempDataTable("SELECT StaffID, FName & ' ' & SName AS Fullname " & _
                                                         "FROM STAFF")
                ctl.columns.add(cmb)
                cmb.HeaderText = "Staff Member"
                cmb.ValueMember = "StaffID"
                cmb.DisplayMember = "Fullname"
                cmb.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("StaffID").ToString

            Case "DataGridView4"
                SQLCode = "SELECT VolunteerScheduleID, 'Vol: ' & RVLNo & ' - ' & Initials & ' - Room(' & RoomNo & ')' AS VOL, " & _
                "StaffID, Approx, ProcName, " & _
                    "iif(Approx='Set Time',dateadd('d',DaysPost, TimepointDateTime), dateadd('n',Minspost, " & _
                    "dateadd('h',HoursPost,dateadd('d',DaysPost, TimepointDateTime)))) AS CalcDate " & _
                    "FROM (((StudySchedule a " & _
                    "INNER JOIN VolunteerSchedule c ON a.StudyScheduleID=c.StudyScheduleID) " & _
                    "INNER JOIN Volunteer d ON c.VolID=d.VolID) " & _
                    "INNER JOIN ProcTask e ON a.ProcID=e.ProcID) " & _
                    "INNER JOIN VolunteerTimepoint f ON d.VolID=f.VolID " & _
                    "AND a.StudyTimepointID=f.StudyTimepointID " & _
                    "WHERE NOT IsNull(StaffID) " & _
                    "AND iif(Approx='Set Time',dateadd('d',DaysPost, TimepointDateTime), dateadd('n',Minspost, " & _
                    "dateadd('h',HoursPost,dateadd('d',DaysPost, TimepointDateTime)))) > Now() " & _
                    "ORDER BY iif(Approx='Set Time',dateadd('d',DaysPost, TimepointDateTime), dateadd('n',Minspost, " & _
                    "dateadd('h',HoursPost,dateadd('d',DaysPost, TimepointDateTime)))) ASC, RVLNo ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.AllowUserToAddRows = False
                ctl.columns("VolunteerScheduleID").visible = False
                ctl.columns("StaffID").visible = False
                ctl.columns("Vol").readonly = True
                ctl.columns("Approx").readonly = True
                ctl.columns("ProcName").readonly = True
                ctl.columns("CalcDate").readonly = True
                ctl.columns("CalcDate").HeaderText = "Date/Time"
                ctl.columns("Approx").HeaderText = "Timepoint"
                ctl.columns("ProcName").HeaderText = "Procedure"
                ctl.columns("Vol").HeaderText = "Volunteer"
                ctl.columns("CalcDate").DefaultCellStyle.Format = "dd-MMM-yyyy HH:mm"
                Dim cmb As New DataGridViewComboBoxColumn
                cmb.DataSource = OverClass.TempDataTable("SELECT StaffID, FName & ' ' & SName AS Fullname " & _
                                                         "FROM STAFF ORDER BY SName ASC")
                ctl.columns.add(cmb)
                cmb.HeaderText = "Staff Member"
                cmb.ValueMember = "StaffID"
                cmb.DisplayMember = "Fullname"
                cmb.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("StaffID").ToString

            Case "DataGridView12"
                ctl.columns("StaffProcID").visible = False
                ctl.columns("StaffID").visible = False
                ctl.columns("ProcID").visible = False
                ctl.columns("ProcDateTime").HeaderText = "Date/Time"
                ctl.columns("ProcDateTime").DefaultCellStyle.Format = "dd-MMM-yyyy HH:mm"
                Dim cmb As New DataGridViewComboBoxColumn
                cmb.DataSource = OverClass.TempDataTable("SELECT StaffID, FName & ' ' & SName AS Fullname " & _
                                                         "FROM STAFF ORDER BY SName ASC")
                ctl.columns.add(cmb)
                cmb.HeaderText = "Staff Member"
                cmb.ValueMember = "StaffID"
                cmb.DisplayMember = "Fullname"
                cmb.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("StaffID").ToString
                Dim cmb2 As New DataGridViewComboBoxColumn
                cmb2.DataSource = OverClass.TempDataTable("SELECT ProcID, ProcName " & _
                                                         "FROM ProcTask ORDER BY ProcName ASC")
                ctl.columns.add(cmb2)
                cmb2.HeaderText = "Procedure"
                cmb2.ValueMember = "ProcID"
                cmb2.DisplayMember = "ProcName"
                cmb2.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("ProcID").ToString
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

        Select Case e.TabPageIndex

            Case 0
                ctl = Me.DataGridView1
                SQLCode = "SELECT ProcID, ProcName, MinsTaken, ProcOrd FROM ProcTask ORDER BY ProcName ASC"
                OverClass.CreateDataSet(SQLCode, Bind, ctl)

            Case 1
                ctl = Me.DataGridView2
                SQLCode = "SELECT StaffID, FName, SName FROM Staff ORDER BY SName ASC"
                OverClass.CreateDataSet(SQLCode, Bind, ctl)

        End Select


        Call Specifics(ctl)

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

        Select Case e.TabPageIndex

            Case 0
                ctl = Me.DataGridView3
                SQLCode = "SELECT StudyID, StudyCode, Colour FROM Study ORDER BY StudyCode ASC"
                OverClass.CreateDataSet(SQLCode, Bind, ctl)

            Case 1
                ctl = Me.DataGridView5
                Me.ComboBox2.DataSource = OverClass.TempDataTable("SELECT StudyID, " & _
                                                              "StudyCode FROM Study ORDER BY StudyCode ASC")
                Me.ComboBox2.ValueMember = "StudyID"
                Me.ComboBox2.DisplayMember = "StudyCode"
            Case 2
                ctl = Me.DataGridView6
                Me.ComboBox4.DataSource = OverClass.TempDataTable("SELECT DISTINCT a.StudyID, " & _
                                                              "StudyCode FROM Study a INNER JOIN StudyTimepoint b " & _
                                                              "ON a.StudyID=b.StudyID " & _
                                                              "ORDER BY StudyCode ASC")
                Me.ComboBox4.ValueMember = "StudyID"
                Me.ComboBox4.DisplayMember = "StudyCode"
                Me.ComboBox3.DataSource = OverClass.TempDataTable("SELECT StudyTimepointID, " & _
                                                              "TimepointName FROM StudyTimepoint WHERE StudyID=" _
                                                              & Me.ComboBox4.SelectedValue.ToString & _
                                                              " ORDER BY TimepointName ASC")
                Me.ComboBox3.ValueMember = "StudyTimepointID"
                Me.ComboBox3.DisplayMember = "TimepointName"
            Case 3
                ctl = Me.DataGridView7
                Me.ComboBox5.DataSource = OverClass.TempDataTable("SELECT StudyID, " & _
                                                              "StudyCode FROM Study ORDER BY StudyCode ASC")
                Me.ComboBox5.ValueMember = "StudyID"
                Me.ComboBox5.DisplayMember = "StudyCode"

        End Select


        Call Specifics(ctl)
    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick
        Dim senderGrid = DirectCast(sender, DataGridView)

        If TypeOf senderGrid.Columns(e.ColumnIndex) Is DataGridViewButtonColumn AndAlso
           e.RowIndex >= 0 Then
            Dim cDialog As New ColorDialog()
            If (cDialog.ShowDialog() = DialogResult.OK) Then
                Me.DataGridView3.Rows(e.RowIndex).Cells("Colour").Value = cDialog.Color.ToArgb
                Me.DataGridView3.Rows(e.RowIndex).Cells("Colour").Style.BackColor = cDialog.Color
            End If
        End If

    End Sub

    Private Sub DataGridView3_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DataGridView3.RowPrePaint
        For Each row In Me.DataGridView3.Rows
            If Not IsNothing(row.Cells("Colour").value) And Not IsDBNull(row.Cells("Colour").value) Then
                Dim colourlong As Long = row.Cells("Colour").value
                row.Cells("Colour").Style.BackColor = Color.FromArgb(row.Cells("Colour").value)
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

        Select Case e.TabPageIndex

            Case 0
                ctl = Me.DataGridView8
                Me.ComboBox6.DataSource = OverClass.TempDataTable("SELECT DISTINCT a.StudyID, " & _
                                                              "StudyCode FROM (Study a INNER JOIN StudyTimePoint b " & _
                                                              "ON a.StudyID=b.StudyID) " & _
                                                              "INNER JOIN Cohort c ON a.StudyID=c.StudyID " & _
                                                              "ORDER BY StudyCode ASC")
                Me.ComboBox6.ValueMember = "StudyID"
                Me.ComboBox6.DisplayMember = "StudyCode"
                Me.ComboBox7.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                              "CohortName FROM Cohort WHERE StudyID=" _
                                                              & Me.ComboBox6.SelectedValue.ToString & _
                                                              " ORDER BY CohortName ASC")
                Me.ComboBox7.ValueMember = "CohortID"
                Me.ComboBox7.DisplayMember = "CohortName"

            Case 1
                Me.ComboBox8.DataSource = OverClass.TempDataTable("SELECT a.CohortID, StudyCode & ' - ' & CohortName AS Display " & _
                                                              "FROM (SELECT StudyCode, CohortName, CohortID, " & _
                                                              "Count(StudyTimepointID) as NumTimepoint " & _
                                                              "FROM (Study a INNER JOIN StudyTimePoint b " & _
                                                              "ON a.StudyID=b.StudyID) " & _
                                                              "INNER JOIN Cohort c ON a.StudyID=c.StudyID " & _
                                                              "GROUP BY StudyCode, CohortName, CohortID) as a " & _
                                                              "INNER JOIN " & _
                                                              "(SELECT c.CohortID, Count(CohortTimepointID) as NumTimepoint " & _
                                                              "FROM CohortTimepoint c INNER JOIN Cohort d " & _
                                                              "ON c.CohortID=d.CohortID WHERE Generated=False " & _
                                                              "GROUP BY c.CohortID) as b " & _
                                                              "ON a.CohortID=b.CohortID AND a.NumTimepoint=b.NumTimepoint")
                Me.ComboBox8.ValueMember = "CohortID"
                Me.ComboBox8.DisplayMember = "Display"

            Case 2
                ctl = Me.DataGridView9
                Me.ComboBox9.DataSource = OverClass.TempDataTable("SELECT DISTINCT a.StudyID, " & _
                                                              "StudyCode FROM (Study a INNER JOIN StudyTimePoint b " & _
                                                              "ON a.StudyID=b.StudyID) " & _
                                                              "INNER JOIN Cohort c ON a.StudyID=c.StudyID " & _
                                                              "WHERE Generated=True " & _
                                                              "ORDER BY StudyCode ASC")
                Me.ComboBox9.ValueMember = "StudyID"
                Me.ComboBox9.DisplayMember = "StudyCode"
                Me.ComboBox10.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                              "CohortName FROM Cohort WHERE StudyID=" _
                                                              & Me.ComboBox9.SelectedValue.ToString & _
                                                                " AND Generated=True" & _
                                                                " ORDER BY CohortName ASC")
                Me.ComboBox10.ValueMember = "CohortID"
                Me.ComboBox10.DisplayMember = "CohortName"

            Case 3
                ctl = Me.DataGridView10
                Me.ComboBox11.DataSource = OverClass.TempDataTable("SELECT DISTINCT a.StudyID, " & _
                                                              "StudyCode FROM (Study a INNER JOIN StudyTimePoint b " & _
                                                              "ON a.StudyID=b.StudyID) " & _
                                                              "INNER JOIN Cohort c ON a.StudyID=c.StudyID " & _
                                                              "WHERE Generated=True " & _
                                                              "ORDER BY StudyCode ASC")
                Me.ComboBox11.ValueMember = "StudyID"
                Me.ComboBox11.DisplayMember = "StudyCode"
                Me.ComboBox12.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                              "CohortName FROM Cohort WHERE StudyID=" _
                                                              & Me.ComboBox11.SelectedValue.ToString & _
                                                                " AND Generated=True" & _
                                                                " ORDER BY CohortName ASC")
                Me.ComboBox12.ValueMember = "CohortID"
                Me.ComboBox12.DisplayMember = "CohortName"
                Me.ComboBox13.DataSource = OverClass.TempDataTable("SELECT RVLNo & ' - ' & Initials AS Display, VolID " & _
                                                              "FROM Volunteer WHERE CohortID=" _
                                                              & Me.ComboBox12.SelectedValue.ToString & _
                                                                " ORDER BY Initials ASC")
                Me.ComboBox13.ValueMember = "VolID"
                Me.ComboBox13.DisplayMember = "Display"

        End Select


        Call Specifics(ctl)

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

        Select e.TabPageIndex

            Case 0
                ctl = Me.DataGridView11
                Me.ComboBox14.DataSource = OverClass.TempDataTable("SELECT DISTINCT a.StudyID, " & _
                                                              "StudyCode FROM (Study a INNER JOIN StudyTimePoint b " & _
                                                              "ON a.StudyID=b.StudyID) " & _
                                                              "INNER JOIN Cohort c ON a.StudyID=c.StudyID " & _
                                                              "WHERE Generated=True " & _
                                                              "ORDER BY StudyCode ASC")
                Me.ComboBox14.ValueMember = "StudyID"
                Me.ComboBox14.DisplayMember = "StudyCode"
                Me.ComboBox15.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                              "CohortName FROM Cohort WHERE StudyID=" _
                                                              & Me.ComboBox14.SelectedValue.ToString & _
                                                              " AND Generated=True " & _
                                                              " ORDER BY CohortName ASC")
                Me.ComboBox15.ValueMember = "CohortID"
                Me.ComboBox15.DisplayMember = "CohortName"
            Case 1
                ctl = Me.DataGridView4
                Me.ComboBox1.DataSource = OverClass.TempDataTable("SELECT DISTINCT a.StudyID, " & _
                                                              "StudyCode FROM (Study a INNER JOIN StudyTimePoint b " & _
                                                              "ON a.StudyID=b.StudyID) " & _
                                                              "INNER JOIN Cohort c ON a.StudyID=c.StudyID " & _
                                                              "WHERE Generated=True " & _
                                                              "ORDER BY StudyCode ASC")
                Me.ComboBox1.ValueMember = "StudyID"
                Me.ComboBox1.DisplayMember = "StudyCode"
                Me.ComboBox16.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                              "CohortName FROM Cohort WHERE StudyID=" _
                                                              & Me.ComboBox1.SelectedValue.ToString & _
                                                              " AND Generated=True " & _
                                                              " ORDER BY CohortName ASC")
                Me.ComboBox16.ValueMember = "CohortID"
                Me.ComboBox16.DisplayMember = "CohortName"

            Case 2
                ctl = Me.DataGridView12
                SQLCode = "SELECT StaffProcID, StaffID, ProcID, ProcDateTime " & _
                    "FROM StaffProc " & _
                    "WHERE ProcDateTime > Now() ORDER BY ProcDateTime ASC"
                OverClass.CreateDataSet(SQLCode, Bind, ctl)


        End Select

        Call Specifics(ctl)

    End Sub

End Class
