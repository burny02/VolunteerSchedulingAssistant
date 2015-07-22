Module SaveModule
    Public Sub Saver(ctl As Object)

        'Get a generic command list first - Ignore errors (Multi table)
        Dim cb As New OleDb.OleDbCommandBuilder(Central.CurrentDataAdapter)

        Try
            Central.CurrentDataAdapter.UpdateCommand = cb.GetUpdateCommand()
        Catch
        End Try
        Try
            Central.CurrentDataAdapter.InsertCommand = cb.GetInsertCommand()
        Catch
        End Try
        Try
            Central.CurrentDataAdapter.DeleteCommand = cb.GetDeleteCommand()
        Catch
        End Try


        'Create and overwrite a custom one if needed (More than 1 table) ...OLEDB Parameters must be added in the order they are used
        Select Case ctl.name

            Case "DataGridView4"

                Dim PKey As Double = Form1.ComboBox1.SelectedValue.ToString

                'SET THE Commands, with Parameters (OLDB Parameters must be added in the order they are used in the statement)
                Central.CurrentDataAdapter.InsertCommand = New OleDb.OleDbCommand("INSERT INTO StudyDay (DayNumber, StudyID) " & _
                                                                          "VALUES (@P1, " & PKey & ")")


                'Add parameters with the source columns in the dataset
                With Central.CurrentDataAdapter.InsertCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Double, 255, "DayNumber")
                End With

            Case "DataGridView5"

                Dim PKey As Double = Form1.ComboBox2.SelectedValue.ToString

                Central.CurrentDataAdapter.InsertCommand = New OleDb.OleDbCommand("INSERT INTO StudyTimepoint " & _
                                                                          "(TimepointName, StudyID) " & _
                                                                          "VALUES (@P1, " & PKey & ")")


                With Central.CurrentDataAdapter.InsertCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.VarChar, 255, "TimepointName")
                End With

            Case "DataGridView6"

                Dim PKey As Double = Form1.ComboBox3.SelectedValue.ToString

                Central.CurrentDataAdapter.InsertCommand = New OleDb.OleDbCommand("INSERT INTO StudySchedule " & _
                                                                          "(DayID, ProcID, HoursPost, MinsPost, Approx, SetTime) " & _
                                                                          "VALUES (" & PKey & ", @P1, @P2, @P3, @P4, @P5)")


                With Central.CurrentDataAdapter.InsertCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Double, 255, "ProcID")
                    .Add("@P2", OleDb.OleDbType.Integer, 255, "HoursPost")
                    .Add("@P3", OleDb.OleDbType.Integer, 255, "MinsPost")
                    .Add("@P4", OleDb.OleDbType.VarChar, 255, "Approx")
                    .Add("@P5", OleDb.OleDbType.DBTime, 255, "SetTime")
                End With


        End Select


        Call Central.SetCommandConnection()
        Call Central.UpdateBackend(ctl)


    End Sub


End Module
