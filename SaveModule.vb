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

                Dim PKey As Double = Form1.ComboBox3.SelectedValue.ToString
                DisplayMessage = False

                Dim ProcID, DaysPost, HoursPost, MinsPost As Object
                Dim Approx As String
                Dim SetTime As Object
                Dim Combine As String = vbNullString

                For Each row In OverClass.CurrentDataSet.Tables(0).Rows

                    If row.RowState = DataRowState.Added Then

                        If IsDBNull(row.item("ProcID")) Then
                            MsgBox("Procedure missing")
                            Exit Sub
                        End If

                        If IsDBNull(row.item("DaysPost")) Then
                            MsgBox("Days missing")
                            Exit Sub
                        End If

                        If IsDBNull(row.item("Approx")) Then
                            MsgBox("Timepoint missing")
                            Exit Sub
                        End If

                        If (IsDBNull(row.item("SetTime")) And row.item("Approx") = "Set Time") _
                            Or (Not IsDBNull(row.item("SetTime")) And row.item("Approx") <> "Set Time") Then
                            MsgBox("Set Time OR Hours/Mins - Not Both")
                            Exit Sub
                        End If

                        If (Not IsDBNull(row.item("SetTime")) And (Not IsDBNull(row.item("MinsPost")) Or Not IsDBNull(row.item("HoursPost")))) _
                            Or (IsDBNull(row.item("SetTime")) And (IsDBNull(row.item("MinsPost")) Or IsDBNull(row.item("HoursPost")))) Then
                            MsgBox("Set Time OR Hours/Mins - Not Both")
                            Exit Sub
                        End If


                        ProcID = CDbl(row.item("ProcID"))
                        DaysPost = CDbl(row.item("DaysPost"))
                        Approx = "'" & row.item("Approx") & "'"

                        If IsDBNull(row.item("HoursPost")) Then
                            HoursPost = "Null"
                        Else
                            HoursPost = CDbl(row.item("HoursPost"))
                        End If

                        If IsDBNull(row.item("MinsPost")) Then
                            MinsPost = "Null"
                        Else
                            MinsPost = CDbl(row.item("MinsPost"))
                        End If

                        If IsDBNull(row.item("SetTime")) Then
                            SetTime = "Null"
                        Else
                            SetTime = OverClass.SQLDate(CDate(row.item("SetTime")))
                        End If

                        Dim cmdInsert As OleDb.OleDbCommand = Nothing
                        Dim SchedID As Long = 0

                        SchedID = (OverClass.TempDataTable("SELECT Max(StudyScheduleID) FROM StudySchedule").Rows(0).Item(0)) + 1


                        Try
                            'INSERT TO SCHEDULE TABLE
                            Combine = "INSERT INTO StudySchedule " & _
                                         "(StudyScheduleID, StudyTimepointID, ProcID, DaysPost, HoursPost, MinsPost, Approx, SetTime) " & _
                                     "VALUES (" & SchedID & ", " & PKey & ", " & ProcID & ", " & DaysPost & ", " & HoursPost & _
                                ", " & MinsPost & ", " & Approx & ", " & SetTime & ")"


                            cmdInsert = New OleDb.OleDbCommand(Combine)

                            OverClass.ExecuteSQL(cmdInsert)

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

                        OverClass.ExecuteSQL(cmdInsert)

                        row.acceptchanges()




                    End If


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

            Case "DataGridView4"


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
