Module SaveModule
    Public Sub Saver(ctl As Object)

        Dim DisplayMessage As Boolean = True

        'Get a generic command list first - Ignore errors (Multi table)
        Dim cb As New OleDb.OleDbCommandBuilder(OverClass.CurrentDataAdapter)

        Try
            OverClass.CurrentDataAdapter.UpdateCommand = cb.GetUpdateCommand()
        Catch
        End Try
        Try
            OverClass.CurrentDataAdapter.InsertCommand = cb.GetInsertCommand()
        Catch
        End Try
        Try
            OverClass.CurrentDataAdapter.DeleteCommand = cb.GetDeleteCommand()
        Catch
        End Try


        'Create and overwrite a custom one if needed (More than 1 table) ...OLEDB Parameters must be added in the order they are used
        Select Case ctl.name


            Case "DataGridView5"

                Dim PKey As Double = Form1.ComboBox2.SelectedValue.ToString

                OverClass.CurrentDataAdapter.InsertCommand = New OleDb.OleDbCommand("INSERT INTO StudyTimepoint " & _
                                                                          "(TimepointName, StudyID) " & _
                                                                          "VALUES (@P1, " & PKey & ")")


                With OverClass.CurrentDataAdapter.InsertCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.VarChar, 255, "TimepointName")
                End With

            Case "DataGridView6"

                If IsDBNull(Form1.TextBox1.Text) Then
                    MsgBox("Default Time Missing")
                    OverClass.CmdList.Clear()
                    Exit Sub
                End If

                If Form1.TextBox1.Text = vbNullString Then
                    MsgBox("Default Time Missing")
                    OverClass.CmdList.Clear()
                    Exit Sub
                End If

                Try
                    OverClass.ExecuteSQL("UPDATE StudyTimepoint " & _
                                         "SET DefaultTime=#" & Form1.TextBox1.Text & "#" & _
                                         " WHERE StudyTimepointID=" & Form1.ComboBox3.SelectedValue.ToString)
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Exit Sub
                End Try


                Dim PKey As Double = Form1.ComboBox3.SelectedValue.ToString
                DisplayMessage = False

                Dim ProcID, DaysPost As Double
                Dim TempDate As Date
                Dim ProcTime As String = vbNullString
                Dim Approx As String = vbNullString
                Dim Combine As String = vbNullString
                Dim PassNumber As Long = 0
                Dim DeleteNumber As Long = 0

                For Each row In OverClass.CurrentDataSet.Tables(0).Rows

                    If row.RowState = DataRowState.Deleted Then

                        DeleteNumber = DeleteNumber + 1

                    End If
                Next

                For Each row In OverClass.CurrentDataSet.Tables(0).Rows

                    Dim rowIndex As Long = OverClass.CurrentDataSet.Tables(0).Rows.IndexOf(row)
                    If row.RowState = DataRowState.Added Then rowIndex = rowIndex - DeleteNumber

                    Dim OrigColour As Color = Color.White
                    Dim OrigAltColour As Color = Color.Gainsboro

                    If row.RowState = DataRowState.Added Or row.RowState = DataRowState.Modified Then

                        Form1.DataGridView6.Rows(rowIndex).DefaultCellStyle.BackColor = Color.Red

                        If IsDBNull(row.item("ProcID")) Then
                            MsgBox("Procedure missing")
                            OverClass.CmdList.Clear()
                            Exit Sub
                        End If

                        If IsDBNull(row.item("DaysPost")) Then
                            MsgBox("Days missing")
                            OverClass.CmdList.Clear()
                            Exit Sub
                        End If

                        If IsDBNull(row.item("Approx")) Then
                            MsgBox("Timepoint missing")
                            OverClass.CmdList.Clear()
                            Exit Sub
                        End If

                        If IsDBNull(row.item("ProcTime")) Then
                            MsgBox("Time missing")
                            OverClass.CmdList.Clear()
                            Exit Sub
                        End If

                        Try
                            TempDate = CDate(row.item("ProcTime"))
                        Catch ex As Exception
                            MsgBox("Incorrect Time")
                            OverClass.CmdList.Clear()
                            Exit Sub
                        End Try

                        Try
                            row.item("ProcTime") = Format(row.item("ProcTime"), "HH:mm")
                        Catch ex As Exception
                            MsgBox("Incorrect Time")
                            OverClass.CmdList.Clear()
                            Exit Sub
                        End Try

                        If rowIndex Mod 2 = 0 Then
                            Form1.DataGridView6.Rows(rowIndex).DefaultCellStyle.BackColor = OrigColour
                        Else
                            Form1.DataGridView6.Rows(rowIndex).DefaultCellStyle.BackColor = OrigAltColour
                        End If

                    End If



                    If row.RowState = DataRowState.Added Then

                        ProcID = CDbl(row.item("ProcID"))
                        DaysPost = CDbl(row.item("DaysPost"))
                        Approx = "'" & row.item("Approx") & "'"
                        ProcTime = "#" & row.item("ProcTime") & "#"


                        PassNumber = PassNumber + 1

                        Dim cmdInsert As OleDb.OleDbCommand = Nothing
                        Dim SchedID As Long = 0

                        Try
                            SchedID = (OverClass.TempDataTable("SELECT Max(StudyScheduleID) FROM StudySchedule").Rows(0).Item(0)) + PassNumber

                        Catch ex As Exception
                            SchedID = 1

                        End Try

                        Try
                            'INSERT TO SCHEDULE TABLE
                            Combine = "INSERT INTO StudySchedule " & _
                                         "(StudyScheduleID, StudyTimepointID, ProcID, DaysPost, Approx, ProcTime) " & _
                                     "VALUES (" & SchedID & ", " & PKey & ", " & ProcID & ", " & DaysPost & ", " & Approx & _
                                     ", " & ProcTime & ")"


                            cmdInsert = New OleDb.OleDbCommand(Combine)

                            OverClass.AddToMassSQL(cmdInsert)

                        Catch ex As Exception
                            MsgBox(ex.Message)
                            Exit Sub

                        End Try


                        'INSERT TO VOL SCHEDULE TABLE
                        Combine = "INSERT INTO VolunteerSchedule " & _
                                     "(StudyScheduleID, VolID) " & _
                                 "SELECT " & SchedID & ", VolID " & _
                            "FROM VolunteerTimepoint WHERE StudyTimepointID=" & PKey


                        cmdInsert = New OleDb.OleDbCommand(Combine)

                        OverClass.AddToMassSQL(cmdInsert)

                    End If

                Next

                Try
                    OverClass.ExecuteMassSQL()
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Exit Sub
                End Try

                For Each row In OverClass.CurrentDataSet.Tables(0).Rows

                    If row.RowState = DataRowState.Added Then row.acceptchanges()

                Next


            Case "DataGridView7"

                Dim PKey As Double = Form1.ComboBox5.SelectedValue.ToString

                OverClass.CurrentDataAdapter.InsertCommand = New OleDb.OleDbCommand("INSERT INTO Cohort " & _
                                                                          "(StudyID, CohortName, NumVols) " & _
                                                                          "VALUES (" & PKey & ", @P1, @P2)")


                With OverClass.CurrentDataAdapter.InsertCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.VarChar, 255, "CohortName")
                    .Add("@P2", OleDb.OleDbType.Integer, 255, "NumVols")
                End With

            Case "DataGridView8"

                Dim PKey As Double = Form1.ComboBox7.SelectedValue.ToString

                OverClass.CurrentDataAdapter.InsertCommand = New OleDb.OleDbCommand("INSERT INTO CohortTimepoint " & _
                                                                          "(CohortID, StudyTimepointID, TimepointDateTime, VolGap) " & _
                                                                          "VALUES (" & PKey & ", @P1, @P2, @P3)")

                With OverClass.CurrentDataAdapter.InsertCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Double, 255, "StudyTimePointID")
                    .Add("@P2", OleDb.OleDbType.DBTimeStamp, 255, "TimepointDateTime")
                    .Add("@P3", OleDb.OleDbType.Double, 255, "VolGap")
                End With

                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE CohortTimepoint " & _
                                                                          "SET TimepointDateTime=@P1, StudyTimepointID=@P2, VolGap=@P3 " & _
                                                                        "WHERE CohortTimepointID=@P4")

                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.DBTimeStamp, 255, "TimepointDateTime")
                    .Add("@P2", OleDb.OleDbType.Double, 255, "StudyTimePointID")
                    .Add("@P3", OleDb.OleDbType.Double, 255, "VolGap")
                    .Add("@P4", OleDb.OleDbType.Double, 255, "CohortTimepointID")

                End With

            Case "DataGridView10"


                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE VolunteerTimepoint " & _
                                                                          "SET TimepointDateTime=@P1, DayNumber=@P2 " & _
                                                                        "WHERE VolunteerTimepointID=@P3")

                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.DBTimeStamp, 255, "TimepointDateTime")
                    .Add("@P2", OleDb.OleDbType.Integer, 255, "DayNumber")
                    .Add("@P3", OleDb.OleDbType.Double, 255, "VolunteerTimepointID")


                End With

            Case "DataGridView11"


                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE VolunteerSchedule " & _
                                                                          "SET StaffID=@P1 " & _
                                                                        "WHERE VolunteerScheduleID=@P2")

                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Double, 255, "StaffID")
                    .Add("@P2", OleDb.OleDbType.Double, 255, "VolunteerScheduleID")


                End With


        End Select


        Call OverClass.SetCommandConnection()
        Call OverClass.UpdateBackend(ctl, DisplayMessage)
        If DisplayMessage = False Then MsgBox("Table Updated")


    End Sub


End Module
