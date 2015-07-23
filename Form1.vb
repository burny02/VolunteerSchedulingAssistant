Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        Call StartUpCentral()

        Central.LockCheck()

        Central.LoginCheck()

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

        If Central.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        Call ResetDataGrid()

        Select Case e.TabPageIndex

            Case 1
                Me.TabControl2_Selecting(Me.TabControl2, New TabControlCancelEventArgs(TabPage3, 0, False, TabControlAction.Selecting))
            Case 2
                Me.TabControl3_Selecting(Me.TabControl3, New TabControlCancelEventArgs(TabPage5, 0, False, TabControlAction.Selecting))
            Case 3
                Me.TabControl4_Selecting(Me.TabControl4, New TabControlCancelEventArgs(TabPage15, 0, False, TabControlAction.Selecting))

        End Select


        Call Specifics(ctl)

    End Sub

    Private Sub ResetDataGrid()

        Me.DataGridView1.Columns.Clear()
        Me.DataGridView1.DataSource = Nothing
        Me.DataGridView2.Columns.Clear()
        Me.DataGridView2.DataSource = Nothing
        Me.DataGridView3.Columns.Clear()
        Me.DataGridView3.DataSource = Nothing
        Me.DataGridView4.Columns.Clear()
        Me.DataGridView4.DataSource = Nothing
        Me.DataGridView5.Columns.Clear()
        Me.DataGridView5.DataSource = Nothing
        Me.DataGridView6.Columns.Clear()
        Me.DataGridView6.DataSource = Nothing
        Me.DataGridView7.Columns.Clear()
        Me.DataGridView7.DataSource = Nothing
        Me.DataGridView8.Columns.Clear()
        Me.DataGridView8.DataSource = Nothing
        Me.ComboBox7.SelectedText = ""
        Me.ComboBox6.SelectedText = ""
        Me.ComboBox5.SelectedText = ""
        Me.ComboBox4.SelectedText = ""
        Me.ComboBox3.SelectedText = ""
        Me.ComboBox2.SelectedText = ""
        Me.ComboBox1.SelectedText = ""

    End Sub

    Private Sub Specifics(ctl As Object)

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

            Case "DataGridView4"
                SQLCode = "SELECT DayID, DayNumber FROM StudyDay WHERE StudyID=" _
                    & Me.ComboBox1.SelectedValue.ToString & " ORDER BY DayNumber ASC"
                Central.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.Columns(0).visible = False
                ctl.columns(1).headertext = "Day Number"

            Case "DataGridView5"
                SQLCode = "SELECT StudyTimepointID, TimepointName FROM StudyTimepoint WHERE StudyID=" _
                    & Me.ComboBox2.SelectedValue.ToString & " ORDER BY TimepointName ASC"
                Central.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.Columns(0).visible = False
                ctl.columns(1).headertext = "Timepoint Name"

            Case "DataGridView6"
                SQLCode = "SELECT StudyScheduleID, ProcID, HoursPost, MinsPost, Approx, SetTime" & _
                    " FROM StudySchedule WHERE DayID=" & Me.ComboBox3.SelectedValue.ToString & _
                    " ORDER BY HoursPost ASC, MinsPost ASC"
                Central.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.Columns(0).visible = False
                ctl.columns(1).visible = False
                Dim cmb As New DataGridViewComboBoxColumn
                cmb.HeaderText = "Procedure"
                cmb.DataSource = Central.TempDataSet("SELECT ProcID, ProcName" & _
                                                     " FROM ProcTask ORDER BY ProcName ASC").Tables(0)
                cmb.ValueMember = "ProcID"
                cmb.DisplayMember = "ProcName"
                cmb.DataPropertyName = Central.CurrentDataSet.Tables(0).Columns("ProcID").ToString
                ctl.columns.add(cmb)
                cmb.DisplayIndex = 1
                ctl.columns(2).headertext = "Hours"
                ctl.columns(3).headertext = "Minutes"
                ctl.columns(5).headertext = "Set Time"
                Dim cmb2 As New DataGridViewComboBoxColumn
                cmb2.HeaderText = "Timepoint"
                cmb2.DataSource = Central.TempDataSet("SELECT * FROM (SELECT StudyTimepointID & 'A' As ID, TimepointName & ': Approx' As Display " & _
                                                     " FROM StudyTimepoint WHERE StudyID=" & Me.ComboBox4.SelectedValue.ToString & _
                                                     " UNION ALL " & _
                                                     " SELECT StudyTimepointID & 'T' As ID, TimepointName & ': Timed' As Display " & _
                                                     " FROM StudyTimepoint WHERE StudyID=" & Me.ComboBox4.SelectedValue.ToString & _
                                                     " UNION ALL " & _
                                                     " SELECT '0S' AS ID, 'Set Time' AS Display FROM StudyTimepoint" & _
                                                     " GROUP BY '0S', 'Set Time') ORDER BY Display ASC").Tables(0)
                cmb2.ValueMember = "ID"
                cmb2.DisplayMember = "Display"
                cmb2.DataPropertyName = Central.CurrentDataSet.Tables(0).Columns("Approx").ToString
                ctl.columns.add(cmb2)
                ctl.columns("SetTime").DefaultCellStyle.Format = "hh:mm"
                ctl.columns("Approx").visible = False
                cmb2.DisplayIndex = 2

            Case "DataGridView7"
                SQLCode = "SELECT CohortID, CohortName, VolGap, NumVols" & _
                    " FROM Cohort WHERE StudyID=" & Me.ComboBox5.SelectedValue.ToString & _
                    " ORDER BY CohortName ASC"
                Central.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.Columns("CohortID").visible = False
                ctl.columns("NumVols").HeaderText = "Number of volunteers"
                ctl.columns("CohortName").HeaderText = "Cohort Name"
                ctl.columns("VolGap").HeaderText = "Interval (Minutes)"

            Case "DataGridView8"
                SQLCode = "SELECT CohortTimePointID, StudyTimepointID, TimepointDateTime" & _
                    " FROM CohortTimepoint " & _
                    " WHERE CohortID=" & Me.ComboBox7.SelectedValue.ToString & _
                    " ORDER BY TimepointDateTime ASC"
                Central.CreateDataSet(SQLCode, Me.BindingSource1, ctl)
                ctl.Columns("CohortTimePointID").visible = False
                ctl.Columns("StudyTimePointID").visible = False
                ctl.columns("TimepointDateTime").HeaderText = "Date/Time"
                Dim cmb As New DataGridViewComboBoxColumn
                cmb.HeaderText = "Timepoint"
                cmb.DataSource = Central.TempDataSet("SELECT StudyTimepointID, TimepointName " & _
                                                    "FROM StudyTimepoint " & _
                                                    "WHERE StudyID=" & Me.ComboBox6.SelectedValue.ToString).Tables(0)
                cmb.ValueMember = "StudyTimepointID"
                cmb.DisplayMember = "TimepointName"
                cmb.DataPropertyName = Central.CurrentDataSet.Tables(0).Columns("StudyTimepointID").ToString
                ctl.columns.add(cmb)
                cmb.DisplayIndex = 0
        End Select

    End Sub

    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs)
        e.Cancel = False
        Call Central.ErrorHandler(sender, e)
    End Sub

    Private Sub TabControl2_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl2.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If Central.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        Call ResetDataGrid()

        Select Case e.TabPageIndex

            Case 0
                ctl = Me.DataGridView1
                SQLCode = "SELECT ProcID, ProcName, MinsTaken, ProcOrd FROM ProcTask ORDER BY ProcName ASC"
                Central.CreateDataSet(SQLCode, Bind, ctl)

            Case 1
                ctl = Me.DataGridView2
                SQLCode = "SELECT StaffID, FName, SName FROM Staff ORDER BY SName ASC"
                Central.CreateDataSet(SQLCode, Bind, ctl)

        End Select


        Call Specifics(ctl)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call Saver(Me.DataGridView1)
    End Sub

    Private Sub DataGridView2_DataError_1(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView2.DataError
        e.Cancel = False
        Call Central.ErrorHandler(sender, e)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call Saver(Me.DataGridView2)
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Central.UnloadData() = True Then e.Cancel = True
        Call Central.Quitter(True)
    End Sub

    Private Sub DataGridView3_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView3.DataError
        e.Cancel = False
        Call Central.ErrorHandler(sender, e)
    End Sub

    Private Sub TabControl3_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl3.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If Central.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        Call ResetDataGrid()

        Select Case e.TabPageIndex

            Case 0
                ctl = Me.DataGridView3
                SQLCode = "SELECT StudyID, StudyCode, Colour FROM Study ORDER BY StudyCode ASC"
                Central.CreateDataSet(SQLCode, Bind, ctl)

            Case 1
                ctl = Me.DataGridView4
                Me.ComboBox1.DataSource = Central.TempDataSet("SELECT StudyID, " & _
                                                              "StudyCode FROM Study ORDER BY StudyCode ASC").Tables(0)
                Me.ComboBox1.ValueMember = "StudyID"
                Me.ComboBox1.DisplayMember = "StudyCode"
            Case 2
                ctl = Me.DataGridView5
                Me.ComboBox2.DataSource = Central.TempDataSet("SELECT StudyID, " & _
                                                              "StudyCode FROM Study ORDER BY StudyCode ASC").Tables(0)
                Me.ComboBox2.ValueMember = "StudyID"
                Me.ComboBox2.DisplayMember = "StudyCode"
            Case 3
                ctl = Me.DataGridView6
                Me.ComboBox4.DataSource = Central.TempDataSet("SELECT DISTINCT a.StudyID, " & _
                                                              "StudyCode FROM Study a INNER JOIN StudyDay b" & _
                                                              " ON a.StudyID=b.StudyID " & _
                                                              "ORDER BY StudyCode ASC").Tables(0)
                Me.ComboBox4.ValueMember = "StudyID"
                Me.ComboBox4.DisplayMember = "StudyCode"
                Me.ComboBox3.DataSource = Central.TempDataSet("SELECT DayID, " & _
                                                              "DayNumber FROM StudyDay WHERE StudyID=" _
                                                              & Me.ComboBox4.SelectedValue.ToString & _
                                                              " ORDER BY DayNumber ASC").Tables(0)
                Me.ComboBox3.ValueMember = "DayID"
                Me.ComboBox3.DisplayMember = "DayNumber"
            Case 4
                ctl = Me.DataGridView7
                Me.ComboBox5.DataSource = Central.TempDataSet("SELECT StudyID, " & _
                                                              "StudyCode FROM Study ORDER BY StudyCode ASC").Tables(0)
                Me.ComboBox5.ValueMember = "StudyID"
                Me.ComboBox5.DisplayMember = "StudyCode"

        End Select


        Call Specifics(ctl)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call Saver(Me.DataGridView3)
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

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call Saver(Me.DataGridView4)
    End Sub

    Private Sub DataGridView4_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView4.DataError
        e.Cancel = False
        Call Central.ErrorHandler(sender, e)
    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Me.ComboBox1.SelectedValue.ToString <> "System.Data.DataRowView" Then

            If Central.UnloadData() = True Then Exit Sub
            Call ResetDataGrid()
            Call Specifics(Me.DataGridView4)

        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call Saver(Me.DataGridView5)
    End Sub

    Private Sub DataGridView5_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView5.DataError
        e.Cancel = False
        Call Central.ErrorHandler(sender, e)
    End Sub

    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If Me.ComboBox2.SelectedValue.ToString <> "System.Data.DataRowView" Then

            If Central.UnloadData() = True Then Exit Sub
            Call ResetDataGrid()
            Call Specifics(Me.DataGridView5)

        End If
    End Sub

    Private Sub ComboBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox4.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub DataGridView6_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView6.DataError
        e.Cancel = False
        Call Central.ErrorHandler(sender, e)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Call Saver(Me.DataGridView6)
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If Me.ComboBox3.SelectedValue.ToString <> "System.Data.DataRowView" Then

            If Central.UnloadData() = True Then Exit Sub
            Call ResetDataGrid()
            Call Specifics(Me.DataGridView6)

        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        If Me.ComboBox4.SelectedValue.ToString <> "System.Data.DataRowView" Then

            If Central.UnloadData() = True Then Exit Sub
            Me.ComboBox3.DataSource = Central.TempDataSet("SELECT DayID, " & _
                                                              "DayNumber FROM StudyDay WHERE StudyID=" _
                                                              & Me.ComboBox4.SelectedValue.ToString & _
                                                             " ORDER BY DayNumber ASC").Tables(0)
            Call ResetDataGrid()
            Call Specifics(Me.DataGridView6)

        End If
    End Sub

    Private Sub DataGridView6_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView6.CellEnter
        Call Central.SingleClick(sender, e)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Call Saver(Me.DataGridView7)
    End Sub

    Private Sub ComboBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox5.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        If Me.ComboBox5.SelectedValue.ToString <> "System.Data.DataRowView" Then

            If Central.UnloadData() = True Then Exit Sub
            Call ResetDataGrid()
            Call Specifics(Me.DataGridView7)

        End If
    End Sub

    Private Sub DataGridView7_DataError(sender As Object, e As DataGridViewDataErrorEventArgs)
        e.Cancel = False
        Call Central.ErrorHandler(sender, e)
    End Sub

    Private Sub DataGridView8_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView8.DataError
        e.Cancel = False
        Call Central.ErrorHandler(sender, e)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Call Saver(Me.DataGridView8)
    End Sub

    Private Sub ComboBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox6.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub ComboBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox7.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub TabControl4_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl4.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If Central.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        Call ResetDataGrid()

        Select Case e.TabPageIndex

            Case 0
                ctl = Me.DataGridView8
                Me.ComboBox6.DataSource = Central.TempDataSet("SELECT DISTINCT a.StudyID, " & _
                                                              "StudyCode FROM (Study a INNER JOIN StudyTimePoint b " & _
                                                              "ON a.StudyID=b.StudyID) " & _
                                                              "INNER JOIN Cohort c ON a.StudyID=c.StudyID " & _
                                                              "ORDER BY StudyCode ASC").Tables(0)
                Me.ComboBox6.ValueMember = "StudyID"
                Me.ComboBox6.DisplayMember = "StudyCode"
                Me.ComboBox7.DataSource = Central.TempDataSet("SELECT CohortID, " & _
                                                              "CohortName FROM Cohort WHERE StudyID=" _
                                                              & Me.ComboBox6.SelectedValue.ToString & _
                                                              " ORDER BY CohortName ASC").Tables(0)
                Me.ComboBox7.ValueMember = "CohortID"
                Me.ComboBox7.DisplayMember = "CohortName"

        End Select


        Call Specifics(ctl)

    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        If Me.ComboBox7.SelectedValue.ToString <> "System.Data.DataRowView" Then
            If Me.ComboBox7.SelectedValue.ToString = vbNullString Then
                If Central.UnloadData() = True Then Exit Sub
                Call Specifics(Me.DataGridView8)
                MsgBox(1 & Me.ComboBox7.SelectedValue.ToString)
            End If
        End If
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged

        If Me.ComboBox6.SelectedValue.ToString <> "System.Data.DataRowView" Then
            If Me.ComboBox6.SelectedValue.ToString = vbNullString Then
                If Central.UnloadData() = True Then Exit Sub
                Me.ComboBox7.DataSource = Central.TempDataSet("SELECT CohortID, " & _
                                                                            "CohortName FROM Cohort WHERE StudyID=" _
                                                                            & Me.ComboBox6.SelectedValue.ToString & _
                                                                            " ORDER BY CohortName ASC").Tables(0)
                Call Specifics(Me.DataGridView8)
                MsgBox(2 & Me.ComboBox6.SelectedValue.ToString)
            End If
        End If

    End Sub


End Class
