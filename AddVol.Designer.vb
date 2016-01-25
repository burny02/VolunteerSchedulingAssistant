<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AddVol
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AddVol))
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.SplitContainer12 = New System.Windows.Forms.SplitContainer()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer12, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer12.Panel1.SuspendLayout()
        Me.SplitContainer12.Panel2.SuspendLayout()
        Me.SplitContainer12.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.Gainsboro
        Me.DataGridView1.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.Size = New System.Drawing.Size(494, 379)
        Me.DataGridView1.TabIndex = 0
        '
        'SplitContainer12
        '
        Me.SplitContainer12.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer12.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer12.Name = "SplitContainer12"
        Me.SplitContainer12.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer12.Panel1
        '
        Me.SplitContainer12.Panel1.Controls.Add(Me.Label15)
        Me.SplitContainer12.Panel1.Controls.Add(Me.ComboBox1)
        Me.SplitContainer12.Panel1.Controls.Add(Me.Label16)
        Me.SplitContainer12.Panel1.Controls.Add(Me.ComboBox2)
        '
        'SplitContainer12.Panel2
        '
        Me.SplitContainer12.Panel2.Controls.Add(Me.DataGridView1)
        Me.SplitContainer12.Size = New System.Drawing.Size(494, 414)
        Me.SplitContainer12.SplitterDistance = 31
        Me.SplitContainer12.TabIndex = 13
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(119, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(54, 20)
        Me.Label15.TabIndex = 23
        Me.Label15.Text = "Study:"
        '
        'ComboBox1
        '
        Me.ComboBox1.Dock = System.Windows.Forms.DockStyle.Right
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(173, 0)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ComboBox1.Size = New System.Drawing.Size(161, 21)
        Me.ComboBox1.TabIndex = 22
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(334, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(61, 20)
        Me.Label16.TabIndex = 21
        Me.Label16.Text = "Cohort:"
        '
        'ComboBox2
        '
        Me.ComboBox2.Dock = System.Windows.Forms.DockStyle.Right
        Me.ComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(395, 0)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ComboBox2.Size = New System.Drawing.Size(99, 21)
        Me.ComboBox2.TabIndex = 20
        '
        'AddVol
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(494, 414)
        Me.Controls.Add(Me.SplitContainer12)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "AddVol"
        Me.Text = "Pick Volunteer"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer12.Panel1.ResumeLayout(False)
        Me.SplitContainer12.Panel1.PerformLayout()
        Me.SplitContainer12.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer12, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer12.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents SplitContainer12 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox

    Private Sub AddVol_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'SET UP PICKING FORM
        With Me

            ComboBox1.DataSource = OverClass.TempDataTable("SELECT DISTINCT a.StudyID, " & _
                                                          "StudyCode FROM (Study a INNER JOIN StudyTimePoint b " & _
                                                          "ON a.StudyID=b.StudyID) " & _
                                                          "INNER JOIN Cohort c ON a.StudyID=c.StudyID " & _
                                                          "WHERE Generated=True " & _
                                                          "ORDER BY StudyCode ASC")
            ComboBox1.ValueMember = "StudyID"
            ComboBox1.DisplayMember = "StudyCode"

            If IsNothing(Me.ComboBox1.SelectedValue) Then Exit Sub
            ComboBox2.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                          "CohortName FROM Cohort WHERE StudyID=" _
                                                          & Me.ComboBox1.SelectedValue.ToString & _
                                                          " AND Generated=True " & _
                                                          " ORDER BY CohortName ASC")
            ComboBox2.ValueMember = "CohortID"
            ComboBox2.DisplayMember = "CohortName"


            With .DataGridView1
                .DataSource = OverClass.TempDataTable("SELECT VolID, RVlNo, Initials, RoomNo, a.CohortID " & _
                                                                       "FROM ((Volunteer a INNER JOIN Cohort b ON a.CohortID=b.CohortID) " & _
                                                                        "INNER JOIN Study c ON b.StudyID=c.StudyID) " & _
                                                                        "WHERE a.CohortID=" & Me.ComboBox2.SelectedValue.ToString & _
                                                                    " ORDER BY RVLNo ASC")

                .Columns("VolID").Visible = False
                .Columns("CohortID").Visible = False
                .Columns("RVLNo").HeaderText = "RVL Number"
                .Columns("RoomNo").HeaderText = "Room Number"
                Dim clm As New DataGridViewImageColumn
                clm.HeaderText = "Pick Volunteer"
                clm.Image = My.Resources.Plus
                clm.ImageLayout = DataGridViewImageCellLayout.Zoom
                clm.Name = "AddButton"
                .Columns.Add(clm)


            End With
        End With

    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted

        If sender.SelectedValue.ToString = vbNullString Then Exit Sub

        If IsNothing(Me.ComboBox1.SelectedValue) Then Exit Sub
        ComboBox2.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                      "CohortName FROM Cohort WHERE StudyID=" _
                                                      & Me.ComboBox1.SelectedValue.ToString & _
                                                      " AND Generated=True " & _
                                                      " ORDER BY CohortName ASC")
        ComboBox2.ValueMember = "CohortID"
        ComboBox2.DisplayMember = "CohortName"

        Me.DataGridView1.DataSource = OverClass.TempDataTable("SELECT VolID, RVlNo, Initials, RoomNo, a.CohortID " & _
                                                                       "FROM ((Volunteer a INNER JOIN Cohort b ON a.CohortID=b.CohortID) " & _
                                                                        "INNER JOIN Study c ON b.StudyID=c.StudyID) " & _
                                                                        "WHERE a.CohortID=" & Me.ComboBox2.SelectedValue.ToString & _
                                                                    " ORDER BY RVLNo ASC")

    End Sub

    Private Sub ComboBox2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox2.SelectionChangeCommitted

        If sender.SelectedValue.ToString = vbNullString Then Exit Sub


        Me.DataGridView1.DataSource = OverClass.TempDataTable("SELECT VolID, RVlNo, Initials, RoomNo, a.CohortID " & _
                                                                       "FROM ((Volunteer a INNER JOIN Cohort b ON a.CohortID=b.CohortID) " & _
                                                                        "INNER JOIN Study c ON b.StudyID=c.StudyID) " & _
                                                                        "WHERE a.CohortID=" & Me.ComboBox2.SelectedValue.ToString & _
                                                                    " ORDER BY RVLNo ASC")
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        If IsDBNull(sender.item("VolID", e.RowIndex).value) Then Exit Sub

        If e.ColumnIndex = sender.columns("AddButton").index Then

            Dim VolID As Long = sender.item("VolID", e.RowIndex).value
            Dim Initials, RVLNo As String
            Dim Accepted As Boolean = False
            Dim Temp As String = vbNullString
            Dim TempDate As Date
            Dim NewVolID As Long
            Dim OldCohortID As Long = sender.item("CohortID", e.RowIndex).value
            RVLNo = sender.item("RVLNO", e.RowIndex).value
            Initials = sender.item("Initials", e.RowIndex).value

            Dim InsertString, TimepointName As String
            Dim cmdInsert As OleDb.OleDbCommand

            'GET A NEW VOL ID
            Try
                NewVolID = (OverClass.TempDataTable("SELECT Max(VolID) FROM Volunteer").Rows(0).Item(0)) + 1
            Catch ex As Exception
                NewVolID = 1
            End Try


            'TRY AND INSERT VOLUNTEER (WITH NEW VOLID AND COHORTID)
            Try
                InsertString = "INSERT INTO Volunteer " &
                                         "(VolID, RVLNo, Initials, CohortID, RoomNo) " &
                                     "SELECT " & NewVolID & ", RVLNo, Initials, " & PickCohort & ", RoomNo FROM Volunteer " &
                                     "WHERE VolID=" & VolID


                cmdInsert = New OleDb.OleDbCommand(InsertString)

                OverClass.ExecuteSQL(cmdInsert)

            Catch ex As Exception
                MsgBox(ex.Message)
                Exit Sub

            End Try


            For Each row In OverClass.TempDataTable("SELECT TimepointName, StudyTimepointID FROM ((StudyTimepoint a " & _
                                                        "INNER JOIN Study b ON a.StudyID=b.StudyID) " & _
                                                        "INNER JOIN Cohort c ON b.StudyID=C.StudyID) " & _
                                                        "WHERE CohortID=" & PickCohort & _
                                                        " GROUP BY TimepointName, StudyTimepointID").Rows
                Accepted = False

                Dim StudyTimepointID, OldStudyTimepointID As Long

                Dim dt As DataTable

                TimepointName = row.item("TimepointName")
                StudyTimepointID = row.item("StudyTimepointID")



                'DOES THE TIMEPOINT EXIST ALREADY IN OLD STUDY?
                dt = OverClass.TempDataTable("SELECT StudyTimepointID FROM ((StudyTimepoint a " & _
                                                        "INNER JOIN Study b ON a.StudyID=b.StudyID) " & _
                                                        "INNER JOIN Cohort c ON b.StudyID=C.StudyID) " & _
                                                        "WHERE CohortID=" & OldCohortID & _
                                                        " AND TimepointName='" & TimepointName & "'" & _
                                                        " GROUP BY StudyTimepointID")

                If dt.Rows.Count = 0 Then
                    'IF timepoint name not in old study
                    Do While Accepted = False

                        Temp = InputBox("Input " & Initials & "(" & RVLNo & ") " & TimepointName & " Date/Time", _
                                        TimepointName & " Date", "01-Jan-2010 10:00")

                        Try
                            TempDate = CDate(Temp)
                            If Format(TempDate, "HH:mm") = "00:00" Then Throw New System.Exception

                        Catch ex As Exception
                            MsgBox("Must enter a valid Date/Time to continue")
                            Continue Do

                        End Try
                        Accepted = True
                        Continue Do
                    Loop

                Else
                    'If is same timepoint name in old study
                    OldStudyTimepointID = dt.Rows(0).Item(0)
                    Dim TempTbl As DataTable = OverClass.TempDataTable("SELECT TimepointDateTime FROM VolunteerTimepoint " &
                                                               "WHERE StudyTimepointID=" & OldStudyTimepointID &
                                                               " AND VolID=" & VolID)
                    If TempTbl.Rows.Count = 0 Then
                        Do While Accepted = False

                            Temp = InputBox("Input " & Initials & "(" & RVLNo & ") " & TimepointName & " Date/Time",
                                            TimepointName & " Date", "01-Jan-2010 10:00")

                            Try
                                TempDate = CDate(Temp)
                                If Format(TempDate, "HH:mm") = "00:00" Then Throw New System.Exception

                            Catch ex As Exception
                                MsgBox("Must enter a valid Date/Time to continue")
                                Continue Do

                            End Try
                            Accepted = True
                            Continue Do

                        Loop
                    Else
                        TempDate = TempTbl.Rows(0).Item(0)
                    End If


                End If


                'INSERT INTO VOLUNTEER TIMEPOINT

                InsertString = "INSERT INTO VolunteerTimepoint " & _
                             "(VolID, TimepointDateTime, StudyTimepointID) " & _
                         "VALUES (" & NewVolID & ", " & OverClass.SQLDate(TempDate) & ", " & StudyTimepointID & ")"


                cmdInsert = New OleDb.OleDbCommand(InsertString)

                OverClass.AddToMassSQL(cmdInsert)


                'INSERT ALL PROCEDURES
                For Each row2 In OverClass.TempDataTable("SELECT StudyScheduleID FROM StudySchedule " & _
                                                        "WHERE StudyTimepointID=" & StudyTimepointID).Rows

                    Dim CmdInsert2 As New OleDb.OleDbCommand("INSERT INTO VolunteerSchedule (StudyScheduleID,VolID) " & _
                                                             "VALUES (" & row2.item(0) & ", " & NewVolID & ")")

                    OverClass.AddToMassSQL(CmdInsert2)

                Next

                Accepted = True
                Continue For


            Next


            OverClass.ExecuteMassSQL()


            'UPDATE COHORT TO GENERATED
            Dim UpdateString As String
            Dim cmdUpdate As OleDb.OleDbCommand

            UpdateString = "UPDATE Cohort SET Generated=TRUE, NumVols=NumVols+1 " & _
                    "WHERE CohortID=" & PickCohort & " AND Generated=TRUE"

            cmdUpdate = New OleDb.OleDbCommand(UpdateString)

            OverClass.ExecuteSQL(cmdUpdate)

            UpdateString = "UPDATE Cohort SET Generated=TRUE, NumVols=1 " & _
                "WHERE CohortID=" & PickCohort & " AND Generated=False"

            cmdUpdate = New OleDb.OleDbCommand(UpdateString)

            OverClass.ExecuteSQL(cmdUpdate)



            'RETURN TO MAIN WINDOW
            MsgBox("Volunteer Added")
            Me.Close()

        End If

    End Sub
End Class
