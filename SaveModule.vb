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

                OverClass.CurrentDataAdapter.InsertCommand = New OleDb.OleDbCommand("INSERT INTO StudyTimepoint " &
                                                                          "(TimepointName, StudyID) " &
                                                                          "VALUES (@P1,@P2)")


                With OverClass.CurrentDataAdapter.InsertCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.VarChar, 255, "TimepointName")
                    .Add("@P2", OleDb.OleDbType.Double, 255, "StudyID")
                End With

            Case "DataGridView6"

                OverClass.CurrentDataAdapter.DeleteCommand = New OleDb.OleDbCommand("DELETE FROM StudySchedule " &
                                                                        "WHERE StudyScheduleID=@P5")

                With OverClass.CurrentDataAdapter.DeleteCommand.Parameters

                    .Add("@P5", OleDb.OleDbType.Double, 255, "StudyScheduleID")

                End With

                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE StudySchedule " &
                                                                          "SET ProcTime=@P1, Approx=@P2, DaysPost=@P3, ProcID=@P4 " &
                                                                        "WHERE StudyScheduleID=@P5")

                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.DBTimeStamp, 255, "ProcTime")
                    .Add("@P2", OleDb.OleDbType.VarChar, 255, "Approx")
                    .Add("@P3", OleDb.OleDbType.Double, 255, "DaysPost")
                    .Add("@P4", OleDb.OleDbType.Double, 255, "ProcID")
                    .Add("@P5", OleDb.OleDbType.Double, 255, "StudyScheduleID")

                End With
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
                    OverClass.ExecuteSQL("UPDATE StudyTimepoint " &
                                         "SET DefaultTime=#" & Form1.TextBox1.Text & "#" &
                                         " WHERE StudyTimepointID=" & Form1.FilterCombo21.SelectedValue.ToString)
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Exit Sub
                End Try



                DisplayMessage = False

                Dim PKey, ProcID, DaysPost As Double
                Dim TempDate As Date
                Dim ProcTime As String = vbNullString
                Dim Approx As String = vbNullString
                Dim Combine As String = vbNullString
                Dim PassNumber As Long = 0
                Dim DeleteNumber As Long = 0

                'For Each row In OverClass.CurrentDataSet.Tables(0).Rows

                'If row.RowState = DataRowState.Deleted Then

                'DeleteNumber = DeleteNumber + 1

                '   End If
                'Next

                For Each row In OverClass.CurrentDataSet.Tables(0).Rows

                    'Dim rowIndex As Long = OverClass.CurrentDataSet.Tables(0).Rows.IndexOf(row)
                    'If row.RowState = DataRowState.Added Then rowIndex = rowIndex - DeleteNumber

                    If row.RowState = DataRowState.Added Or row.RowState = DataRowState.Modified Then

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

                    End If



                    If row.RowState = DataRowState.Added Then

                        ProcID = CDbl(row.item("ProcID"))
                        DaysPost = CDbl(row.item("DaysPost"))
                        Approx = "'" & row.item("Approx") & "'"
                        ProcTime = "#" & row.item("ProcTime") & "#"
                        PKey = CDbl(row.item("StudyTimepointID"))


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
                            Combine = "INSERT INTO StudySchedule " &
                                         "(StudyScheduleID, StudyTimepointID, ProcID, DaysPost, Approx, ProcTime) " &
                                     "VALUES (" & SchedID & ", " & PKey & ", " & ProcID & ", " & DaysPost & ", " & Approx &
                                     ", " & ProcTime & ")"


                            cmdInsert = New OleDb.OleDbCommand(Combine)

                            OverClass.AddToMassSQL(cmdInsert)

                        Catch ex As Exception
                            MsgBox(ex.Message)
                            Exit Sub

                        End Try


                        'INSERT TO VOL SCHEDULE TABLE
                        Combine = "INSERT INTO VolunteerSchedule " &
                                     "(StudyScheduleID, VolID) " &
                                 "SELECT " & SchedID & ", VolID " &
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


                OverClass.CurrentDataAdapter.InsertCommand = New OleDb.OleDbCommand("INSERT INTO Cohort " &
                                                                          "(StudyID, CohortName, NumVols) " &
                                                                          "VALUES (@P0, @P1, @P2)")


                With OverClass.CurrentDataAdapter.InsertCommand.Parameters
                    .Add("@P0", OleDb.OleDbType.Double, 255, "StudyID")
                    .Add("@P1", OleDb.OleDbType.VarChar, 255, "CohortName")
                    .Add("@P2", OleDb.OleDbType.Integer, 255, "NumVols")
                End With

            Case "DataGridView8"


                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE VolunteerSchedule " &
                                                                          "SET ProcOffSet=@P1 WHERE VolunteerScheduleID=@P2")

                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Double, 255, "ProcOffSet")
                    .Add("@P2", OleDb.OleDbType.Double, 255, "VolunteerScheduleID")

                End With

            Case "DataGridView9"

                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE Volunteer " &
                                                                          "SET RVLNo=@P1, Initials=@P2, RoomNo=@P3 " &
                                                                        "WHERE VolID=@P4")

                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Double, 255, "RVLNo")
                    .Add("@P2", OleDb.OleDbType.LongVarChar, 255, "Initials")
                    .Add("@P3", OleDb.OleDbType.Double, 255, "RoomNo")
                    .Add("@P4", OleDb.OleDbType.Double, 255, "VolID")

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
