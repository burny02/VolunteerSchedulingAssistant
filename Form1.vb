﻿Imports Microsoft.Reporting.WinForms

Public Class Form1
    Private LastValue As Object

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
            Case 5
                StartCombo(Me.ComboBox17)
                StartCombo(Me.ComboBox18)
            Case 6
                ctl = Me.DataGridView13


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
                If IsNothing(Me.ComboBox2.SelectedValue) Then Exit Sub
                SQLCode = "SELECT StudyTimepointID, TimepointName FROM StudyTimepoint WHERE StudyID=" _
                    & Me.ComboBox2.SelectedValue.ToString & " ORDER BY TimepointName ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.Columns(0).visible = False
                ctl.columns(1).headertext = "Timepoint Name"

            Case "DataGridView6"
                If IsNothing(Me.ComboBox3.SelectedValue) Then Exit Sub
                If IsNothing(Me.ComboBox4.SelectedValue) Then Exit Sub
                ctl.autogeneratecolumns = True
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
                cmb.Name = "PickProc"
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
                cmb2.Name = "PickTimepoint"
                Dim cmb3 As New DataGridViewComboBoxColumn
                cmb3.DataSource = OverClass.TempDataTable("SELECT ProcID, MinsTaken " & _
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
                If IsNothing(Me.ComboBox5.SelectedValue) Then Exit Sub
                SQLCode = "SELECT CohortID, CohortName, NumVols, Generated" & _
                    " FROM Cohort WHERE StudyID=" & Me.ComboBox5.SelectedValue.ToString & _
                    " ORDER BY CohortName ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.Columns("CohortID").visible = False
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
                Dim cmb2 As New DataGridViewImageColumn
                cmb2.HeaderText = "Delete Cohort"
                cmb2.Image = My.Resources.Remove
                cmb2.ImageLayout = DataGridViewImageCellLayout.Zoom
                ctl.columns.add(cmb2)
                cmb2.Name = "DeleteButton"


            Case "DataGridView8"
                If IsNothing(Me.ComboBox7.SelectedValue) Then Exit Sub
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
                If IsNothing(Me.ComboBox10.SelectedValue) Then Exit Sub
                SQLCode = "SELECT VolID, RVLNo, Initials, RoomNo " & _
                    "FROM Volunteer " & _
                    "WHERE CohortID=" & Me.ComboBox10.SelectedValue.ToString & _
                    " ORDER BY Initials ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.AllowUserToAddRows = False
                ctl.columns("VolID").visible = False
                ctl.columns("RVLNo").HeaderText = "RVL Number"
                ctl.columns("RoomNo").HeaderText = "Room Number"
                Dim cmb As New DataGridViewImageColumn
                cmb.HeaderText = "Delete Volunteer"
                cmb.Image = My.Resources.Remove
                cmb.ImageLayout = DataGridViewImageCellLayout.Zoom
                ctl.columns.add(cmb)
                cmb.Name = "DeleteButton"
                

            Case "DataGridView10"
                If IsNothing(Me.ComboBox13.SelectedValue) Then Exit Sub
                SQLCode = "SELECT VolunteerTimepointID, TimepointName, TimepointDateTime, DayNumber " & _
                    "FROM VolunteerTimepoint a INNER JOIN StudyTimepoint b " & _
                    "ON a.StudyTimepointID=b.StudyTimepointID " & _
                    "WHERE VolID=" & Me.ComboBox13.SelectedValue.ToString & _
                    " ORDER BY TimepointDateTime ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.AllowUserToAddRows = False
                ctl.columns("VolunteerTimepointID").visible = False
                ctl.columns("TimepointName").Readonly = True
                ctl.columns("TimepointName").HeaderText = "Timepoint Name"
                ctl.columns("DayNumber").HeaderText = "Day Number"
                ctl.columns("TimepointDateTime").HeaderText = "Date/Time"
                ctl.columns("TimepointDateTime").DefaultCellStyle.Format = "dd-MMM-yyyy HH:mm"

            Case "DataGridView11"
                If IsNothing(Me.ComboBox15.SelectedValue) Then Exit Sub
                SQLCode = "SELECT * FROM Assign WHERE CohortID=" & Me.ComboBox15.SelectedValue.ToString & _
                    " ORDER BY CalcDate ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.AllowUserToAddRows = False
                ctl.columns("VolunteerScheduleID").visible = False
                ctl.columns("StaffID").visible = False
                ctl.columns("CohortID").visible = False
                ctl.columns("Vol").readonly = True
                ctl.columns("Approx").readonly = True
                ctl.columns("ProcName").readonly = True
                ctl.columns("CalcDate").readonly = True
                ctl.columns("CalcDate").HeaderText = "Start"
                ctl.columns("EndFull").HeaderText = "Finish"
                ctl.columns("Approx").HeaderText = "Timepoint"
                ctl.columns("ProcName").HeaderText = "Procedure"
                ctl.columns("Vol").HeaderText = "Volunteer"
                ctl.columns("CalcDate").DefaultCellStyle.Format = "dd-MMM-yyyy HH:mm"
                ctl.columns("EndFull").DefaultCellStyle.Format = "HH:mm"
                Dim cmb As New DataGridViewComboBoxColumn
                cmb.DataSource = OverClass.TempDataTable("SELECT StaffID, FName & ' ' & SName AS Fullname " & _
                                                         "FROM STAFF ORDER BY SName ASC")
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

            Case "DataGridView4"
                If IsNothing(Me.ComboBox16.SelectedValue) Then Exit Sub
                SQLCode = "SELECT * FROM Reassign WHERE CohortID=" & Me.ComboBox16.SelectedValue.ToString & _
                    " ORDER BY CalcDate ASC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.AllowUserToAddRows = False
                ctl.columns("VolunteerScheduleID").visible = False
                ctl.columns("StaffID").visible = False
                ctl.columns("CohortId").visible = False
                ctl.columns("Vol").readonly = True
                ctl.columns("Approx").readonly = True
                ctl.columns("ProcName").readonly = True
                ctl.columns("CalcDate").readonly = True
                ctl.columns("CalcDate").HeaderText = "Start"
                ctl.columns("EndFull").HeaderText = "Finish"
                ctl.columns("Approx").HeaderText = "Timepoint"
                ctl.columns("ProcName").HeaderText = "Procedure"
                ctl.columns("Vol").HeaderText = "Volunteer"
                ctl.columns("CalcDate").DefaultCellStyle.Format = "dd-MMM-yyyy HH:mm"
                ctl.columns("EndFull").DefaultCellStyle.Format = "HH:mm"
                Dim cmb As New DataGridViewComboBoxColumn
                cmb.DataSource = OverClass.TempDataTable("SELECT StaffID, FName & ' ' & SName AS Fullname " & _
                                                         "FROM STAFF ORDER BY SName ASC")
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
                cmb.Name = "Pick"
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
                Dim cmb3 As New DataGridViewComboBoxColumn
                cmb3.DataSource = OverClass.TempDataTable("SELECT ProcID, MinsTaken " & _
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

            Case "DataGridView13"
                SQLCode = "SELECT * FROM ReportArchive ORDER BY ArchiveID DESC, ArchiveDate DESC"
                OverClass.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.columns("ArchiveID").Visible = False
                ctl.columns("ArchivePath").visible = False
                ctl.Columns("ArchiveType").displayindex = 0
                ctl.columns("ArchiveType").HeaderText = "Report Type"
                ctl.columns("ArchiveDate").HeaderText = "Date Ran"
                ctl.columns("ArchiveUser").HeaderText = "User Ran"
                ctl.columns("ArchiveCriteria").HeaderText = "Report Criteria"
                ctl.Columns("ArchiveCriteria").DefaultCellStyle.WrapMode = DataGridViewTriState.True
                ctl.Columns("ArchiveType").DefaultCellStyle.Font = New Font("Arial", 10, FontStyle.Underline)
                ctl.Columns("ArchiveType").DefaultCellStyle.ForeColor = Color.Blue

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

        Select Case e.TabPageIndex

            Case 0
                ctl = Me.DataGridView3
                SQLCode = "SELECT StudyID, StudyCode, Colour FROM Study ORDER BY StudyCode ASC"
                OverClass.CreateDataSet(SQLCode, Bind, ctl)

            Case 1
                StartCombo(Me.ComboBox2)

            Case 2
                StartCombo(Me.ComboBox4)
                StartCombo(Me.ComboBox3)

            Case 3
                StartCombo(Me.ComboBox5)

        End Select


        If Not IsNothing(ctl) Then Call Specifics(ctl)

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

        Select Case e.TabPageIndex

            Case 0
                StartCombo(Me.ComboBox6)
                StartCombo(Me.ComboBox7)

            Case 1
                StartCombo(Me.ComboBox8)

            Case 2
                StartCombo(Me.ComboBox9)
                StartCombo(Me.ComboBox10)

            Case 3
                StartCombo(Me.ComboBox11)
                StartCombo(Me.ComboBox12)
                StartCombo(Me.ComboBox13)

        End Select


        If Not IsNothing(ctl) Then Call Specifics(ctl)

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

        Select Case e.TabPageIndex

            Case 0
                StartCombo(Me.ComboBox14)
                StartCombo(Me.ComboBox15)

            Case 1
                StartCombo(Me.ComboBox1)
                StartCombo(Me.ComboBox16)

            Case 2
                ctl = Me.DataGridView12
                SQLCode = "SELECT StaffProcID, StaffID, ProcID, ProcDateTime " & _
                    "FROM StaffProc " & _
                    "WHERE ProcDateTime > Now() ORDER BY ProcDateTime ASC"
                OverClass.CreateDataSet(SQLCode, Bind, ctl)


        End Select

        If Not IsNothing(ctl) Then Call Specifics(ctl)

    End Sub

    Private Sub DataGridView11_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView11.CellBeginEdit

        LastValue = sender.rows(e.RowIndex).cells("StaffID").value

    End Sub

    Private Sub DataGridView4_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView4.CellBeginEdit

        LastValue = sender.rows(e.RowIndex).cells("StaffID").value

    End Sub

    Private Sub DataGridView12_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView12.CellBeginEdit

        If e.ColumnIndex = sender.Columns("Pick").Index Then
            LastValue = sender.rows(e.RowIndex).cells("StaffID").value
        End If

    End Sub

    Private Sub DataGridView12_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView12.CellEndEdit
        Dim returner As String = vbNullString

        If IsDBNull(sender.rows(e.RowIndex).cells("StaffID").value) Or _
            IsNothing(sender.rows(e.RowIndex).cells("StaffID").value) Or _
            IsDBNull(sender.rows(e.RowIndex).cells("ProcID").value) Or _
            IsNothing(sender.rows(e.RowIndex).cells("ProcID").value) Or _
            IsDBNull(sender.rows(e.RowIndex).cells("CalcDate").value) Or _
            IsNothing(sender.rows(e.RowIndex).cells("CalcDate").value) Then Exit Sub


        Dim Identifier As Long
        If IsDBNull(sender.rows(e.RowIndex).cells("StaffProcID").value) Or _
            IsNothing(sender.rows(e.RowIndex).cells("StaffProcID").value) Then
            Identifier = 0
        Else
            Identifier = sender.rows(e.RowIndex).cells("StaffProcID").value
        End If

        returner = CheckExtraOverlap(sender.rows(e.RowIndex).cells("StaffID").value, Identifier, _
                        sender.rows(e.RowIndex).cells("CalcDate").value, _
                        DateAdd(DateInterval.Minute, sender.rows(e.RowIndex).cells("MinsTaken").FormattedValue, sender.rows(e.RowIndex).cells("CalcDate").value), _
                        sender, e.RowIndex)

        If returner <> vbNullString Then
            If MsgBox("Overlap found - " & vbNewLine & vbNewLine & returner & vbNewLine & vbNewLine & _
                   "Do you want to continue?", MsgBoxStyle.YesNo) = vbNo Then
                sender.rows(e.RowIndex).cells("StaffID").value = LastValue
            End If
        End If

    End Sub

    Private Sub DataGridView4_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellEndEdit

        Dim returner As String = vbNullString

        If IsDBNull(sender.rows(e.RowIndex).cells("StaffID").value) Or _
            IsNothing(sender.rows(e.RowIndex).cells("StaffID").value) Then Exit Sub
        If e.ColumnIndex <> sender.Columns("Pick").Index Then Exit Sub


        returner = CheckVolunteerOverlap(sender.rows(e.RowIndex).cells("StaffID").value, sender.rows(e.RowIndex).cells("VolunteerScheduleID").value, _
                              sender.rows(e.RowIndex).cells("CalcDate").value, sender.rows(e.RowIndex).cells("EndFull").value, sender, True)

        If returner <> vbNullString Then
            If MsgBox("Overlap found - " & vbNewLine & vbNewLine & returner & vbNewLine & vbNewLine & _
                   "Do you want to continue?", MsgBoxStyle.YesNo) = vbNo Then
                sender.rows(e.RowIndex).cells("StaffID").value = LastValue
            End If
        End If

    End Sub

    Private Sub DataGridView11_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView11.CellEndEdit

        Dim returner As String = vbNullString

        If IsDBNull(sender.rows(e.RowIndex).cells("StaffID").value) Or _
            IsNothing(sender.rows(e.RowIndex).cells("StaffID").value) Then Exit Sub
        If e.ColumnIndex <> sender.Columns("Pick").Index Then Exit Sub


        returner = CheckVolunteerOverlap(sender.rows(e.RowIndex).cells("StaffID").value, sender.rows(e.RowIndex).cells("VolunteerScheduleID").value, _
                              sender.rows(e.RowIndex).cells("CalcDate").value, sender.rows(e.RowIndex).cells("EndFull").value, sender)

        If returner <> vbNullString Then
            If MsgBox("Overlap found - " & vbNewLine & vbNewLine & returner & vbNewLine & vbNewLine & _
                   "Do you want to continue?", MsgBoxStyle.YesNo) = vbNo Then
                sender.rows(e.RowIndex).cells("StaffID").value = LastValue
            End If
        End If

    End Sub

    Private Sub DataGridView6_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView6.CellEndEdit

        Dim Returner As String = vbNullString

        If IsDBNull(sender.rows(e.RowIndex).cells("ProcID").value) Or _
            IsNothing(sender.rows(e.RowIndex).cells("ProcID").value) Or _
            IsDBNull(sender.rows(e.RowIndex).cells("MinsPost").value) Or _
            IsNothing(sender.rows(e.RowIndex).cells("MinsPost").value) Or _
            IsDBNull(sender.rows(e.RowIndex).cells("DaysPost").value) Or _
            IsNothing(sender.rows(e.RowIndex).cells("DaysPost").value) Or _
            IsDBNull(sender.rows(e.RowIndex).cells("HoursPost").value) Or _
            IsNothing(sender.rows(e.RowIndex).cells("HoursPost").value) Then Exit Sub


        Returner = ScheduleOverlap(sender, e.RowIndex, sender.rows(e.RowIndex).cells("Dayspost").value, _
                              sender.rows(e.RowIndex).cells("HoursPost").value, sender.rows(e.RowIndex).cells("MinsPost").value, _
                              sender.rows(e.RowIndex).cells("MinsTaken").FormattedValue)

        If Returner <> vbNullString Then MsgBox("Overlap found - " & vbNewLine & vbNewLine & Returner)

    End Sub

    Private Sub DataGridView7_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView7.CellBeginEdit

        If IsDBNull(sender.item("CohortID", e.RowIndex).value) Or IsNothing(sender.item("CohortID", e.RowIndex).value) Then Exit Sub
        If sender.item("Generated", e.RowIndex).value = True And e.ColumnIndex = sender.columns("NumVols").index Then e.Cancel = True

    End Sub

    Private Sub DataGridView7_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView7.CellContentClick

        If IsDBNull(sender.item("CohortID", e.RowIndex).value) Then Exit Sub

        If e.ColumnIndex = sender.columns("DeleteButton").index Then
            Dim row As DataGridViewRow
            row = sender.rows(e.RowIndex)
            sender.rows.remove(row)
        End If


        If e.ColumnIndex = sender.columns("AddVolButton").index Then

            Me.ValidateChildren()
            If OverClass.UnloadData() = True Then Exit Sub

            Dim Response As Integer = MsgBox("Do you want to transfer the volunteer from another study?" _
                                             & vbNewLine & vbNewLine _
                                             & "Timepoints of EXACTLY the same name will be transfered across", MsgBoxStyle.YesNoCancel)
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
                        MsgBox("Initials must be 3 characters long")
                        Continue Do
                    End If
                    If Not Initials Like "[A-Z][A-Z][A-Z]" Then
                        MsgBox("Initials must be 3 text characters such as 'AAA'")
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

                    InsertString = "INSERT INTO Volunteer " & _
                                         "(VolID, RVLNo, Initials, CohortID, RoomNo) " & _
                                     "VALUES (" & VolID & ", " & RVLNo & ", '" & Initials & "', " & CohortID & ", " & RoomNo & ")"


                    cmdInsert = New OleDb.OleDbCommand(InsertString)

                    OverClass.ExecuteSQL(cmdInsert)

                Catch ex As Exception
                    MsgBox(ex.Message)
                    Exit Sub

                End Try



                For Each row In OverClass.TempDataTable("SELECT TimepointName, StudyTimepointID FROM ((StudyTimepoint a " & _
                                                        "INNER JOIN Study b ON a.StudyID=b.StudyID) " & _
                                                        "INNER JOIN Cohort c ON b.StudyID=C.StudyID) " & _
                                                        "WHERE CohortID=" & CohortID & _
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

                        Temp = InputBox("Input " & Initials & "(" & RVLNo & ") " & TimepointName & " Date/Time", _
                                        TimepointName & " Date", "01-Jan-2010 10:00")

                        Try
                            TempDate = CDate(Temp)
                            If Format(TempDate, "HH:mm") = "00:00" Then Throw New System.Exception
                            InsertString = "INSERT INTO VolunteerTimepoint " & _
                                         "(VolID, TimepointDateTime, StudyTimepointID) " & _
                                     "VALUES (" & VolID & ", " & OverClass.SQLDate(TempDate) & ", " & StudyTimepointID & ")"


                            cmdInsert = New OleDb.OleDbCommand(InsertString)

                            OverClass.ExecuteSQL(cmdInsert)


                        Catch ex As Exception
                            MsgBox("Must enter a valid Date/Time to continue")
                            Continue Do

                        End Try

                        'INSERT ALL PROCEDURES
                        For Each row2 In OverClass.TempDataTable("SELECT StudyScheduleID FROM StudySchedule " & _
                                                                "WHERE StudyTimepointID=" & StudyTimepointID).Rows

                            Dim CmdInsert2 As New OleDb.OleDbCommand("INSERT INTO VolunteerSchedule (StudyScheduleID,VolID) " & _
                                                                     "VALUES (" & row2.item(0) & ", " & VolID & ")")

                            OverClass.ExecuteSQL(CmdInsert2)

                        Next


                        Accepted = True
                        Continue For
                    Loop

                Next




                'UPDATE COHORT TO GENERATED
                Dim UpdateString As String
                Dim cmdUpdate As OleDb.OleDbCommand

                UpdateString = "UPDATE Cohort SET Generated=TRUE, NumVols=NumVols+1 " & _
                    "WHERE CohortID=" & CohortID & " AND Generated=TRUE"

                cmdUpdate = New OleDb.OleDbCommand(UpdateString)

                OverClass.ExecuteSQL(cmdUpdate)

                UpdateString = "UPDATE Cohort SET Generated=TRUE, NumVols=1 " & _
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

        If e.ColumnIndex <> sender.columns("DeleteButton").index Then Exit Sub
        If IsDBNull(sender.item("VolID", e.RowIndex).value) Then Exit Sub

        Dim row As DataGridViewRow
        row = sender.rows(e.RowIndex)
        sender.rows.remove(row)

    End Sub

    Private Sub DataGridView6_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView6.CellContentClick

        If e.ColumnIndex <> sender.columns("DeleteButton").index Then Exit Sub
        If IsDBNull(sender.item("StudyScheduleID", e.RowIndex).value) Then Exit Sub

        Dim row As DataGridViewRow
        row = sender.rows(e.RowIndex)
        sender.rows.remove(row)

    End Sub

    Private Sub DataGridView11_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView11.CellContentClick

        If e.ColumnIndex <> sender.columns("DeleteButton").index Then Exit Sub
        If IsDBNull(sender.item("VolunteerScheduleID", e.RowIndex).value) Then Exit Sub

        Dim row As DataGridViewRow
        row = sender.rows(e.RowIndex)
        sender.rows.remove(row)

    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

        If e.ColumnIndex <> sender.columns("DeleteButton").index Then Exit Sub
        If IsDBNull(sender.item("VolunteerScheduleID", e.RowIndex).value) Then Exit Sub

        Dim row As DataGridViewRow
        row = sender.rows(e.RowIndex)
        sender.rows.remove(row)

    End Sub

    Private Sub DataGridView12_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView12.CellContentClick

        If e.ColumnIndex <> sender.columns("DeleteButton").index Then Exit Sub
        If IsDBNull(sender.item("StaffProcID", e.RowIndex).value) Then Exit Sub

        Dim row As DataGridViewRow
        row = sender.rows(e.RowIndex)
        sender.rows.remove(row)

    End Sub

    Private Sub DataGridView13_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView13.CellContentClick

        If e.ColumnIndex = sender.columns("ArchiveType").index Then

            Dim FilePath As String = sender.item("ArchivePath", e.RowIndex).value

            Try
                Process.Start(FilePath)

            Catch ex As Exception
                MsgBox(ex.Message)

            End Try

        End If

    End Sub
End Class

