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
                Me.TabControl3_Selecting(Me.TabControl2, New TabControlCancelEventArgs(TabPage5, 0, False, TabControlAction.Selecting))


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
                cmb.DisplayIndex = 2
                ctl.columns(2).headertext = "Hours"
                ctl.columns(3).headertext = "Minutes"
                ctl.columns(5).headertext = "Set Time"

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
                Me.ComboBox4.DataSource = Central.TempDataSet("SELECT StudyID, " & _
                                                              "StudyCode FROM Study ORDER BY StudyCode ASC").Tables(0)
                Me.ComboBox4.ValueMember = "StudyID"
                Me.ComboBox4.DisplayMember = "StudyCode"
                Me.ComboBox3.DataSource = Central.TempDataSet("SELECT DayID, " & _
                                                              "DayNumber FROM StudyDay WHERE StudyID=" _
                                                              & Me.ComboBox4.SelectedValue.ToString & _
                                                              " ORDER BY DayNumber ASC").Tables(0)
                Me.ComboBox3.ValueMember = "DayID"
                Me.ComboBox3.DisplayMember = "DayNumber"
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

End Class
