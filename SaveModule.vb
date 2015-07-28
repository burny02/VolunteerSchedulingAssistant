Module SaveModule
    Public Sub Saver(ctl As Object)

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

                OverClass.CurrentDataAdapter.InsertCommand = New OleDb.OleDbCommand("INSERT INTO StudySchedule " & _
                                                                          "(StudyTimepointID, ProcID, DaysPost, HoursPost, MinsPost, Approx, SetTime) " & _
                                                                          "VALUES (" & PKey & ", @P1, @P2, @P3, @P4, @P5, @P6)")


                With OverClass.CurrentDataAdapter.InsertCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Double, 255, "ProcID")
                    .Add("@P2", OleDb.OleDbType.Integer, 255, "DaysPost")
                    .Add("@P3", OleDb.OleDbType.Integer, 255, "HoursPost")
                    .Add("@P4", OleDb.OleDbType.Integer, 255, "MinsPost")
                    .Add("@P5", OleDb.OleDbType.VarChar, 255, "Approx")
                    .Add("@P6", OleDb.OleDbType.DBTimeStamp, 255, "SetTime")
                End With

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
                                                                          "SET TimepointDateTime=@P1 " & _
                                                                        "WHERE VolunteerTimepointID=@P2")

                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.DBTimeStamp, 255, "TimepointDateTime")
                    .Add("@P2", OleDb.OleDbType.Double, 255, "VolunteerTimepointID")


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
        Call OverClass.UpdateBackend(ctl)


    End Sub


End Module
