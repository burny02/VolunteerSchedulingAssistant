Module ScheduleGenerator

    Public Sub Generator(CohortID As Long)

        'Get and insert cohort Timepoints
        Dim StudyTimepoints As DataTable = OverClass.TempDataTable("SELECT StudyTimepointID, TimepointName " &
        "FROM (Study INNER JOIN StudyTimepoint ON Study.StudyID = StudyTimepoint.StudyID) INNER JOIN Cohort ON Study.StudyID = Cohort.StudyID " &
        "WHERE CohortID=" & CohortID)

        For Each row As DataRow In StudyTimepoints.Rows
            Try
                Dim CohortTimepointDateTime As Date =
                InputBox("Please enter a date/time for cohort timepoint '" & row.Item("TimepointName") & "'",
                row.Item("TimepointName") & " date/time")

                Dim VolGap As Long =
                InputBox("Please enter an interval in minutes between volunteers for timepoint '" & row.Item("TimepointName") & "'",
                row.Item("TimepointName") & " interval")

                OverClass.AddToMassSQL("INSERT INTO CohortTimepoint (CohortID,StudyTimepointID,TimepointDateTime,VolGap) " &
                    "VALUES (" & CohortID & "," & row.Item("StudyTimepointID") & "," & OverClass.SQLDate(CohortTimepointDateTime) & ", " & VolGap & ")")

            Catch ex As Exception
                MsgBox("Incorrect data format entered")
                Exit Sub
            End Try
        Next

        Try
            OverClass.ExecuteMassSQL()
        Catch ex As Exception
            MsgBox("An error occured")
            Exit Sub
        End Try


        Dim DtSchedule As DataTable
        Dim DtCohortTimepoint As DataTable
        Dim DtNumVolunteers As DataTable
        Dim DtVolunteers As DataTable
        Dim i As Long = 0
        Dim j As Long = 0
        Dim NumVols As Long = 0
        Dim NumSchedule As Long = 0

        'Upload into Volunteers Table
        DtNumVolunteers = OverClass.TempDataTable("SELECT NumVols FROM Cohort WHERE CohortID=" & CohortID)

        NumVols = DtNumVolunteers.Rows(0).Item(0)
        DtNumVolunteers = Nothing

        Do While i < NumVols
            OverClass.AddToMassSQL("INSERT INTO Volunteer (CohortID) VALUES (" & CohortID & ")")
            i = i + 1
        Loop
        OverClass.ExecuteMassSQL()
        i = 0


        'Insert volunteers timepoints into Volunteer Timepoints Table
        DtVolunteers = OverClass.TempDataTable("SELECT VolID FROM Volunteer WHERE CohortID=" & CohortID)


        DtCohortTimepoint = OverClass.TempDataTable("SELECT StudyTimepointID, TimepointDateTime, VolGap FROM CohortTimepoint " &
                                              "WHERE CohortID=" & CohortID)

        Do While i < NumVols
            Do While j < DtCohortTimepoint.Rows.Count
                OverClass.AddToMassSQL("INSERT INTO VolunteerTimepoint (StudyTimepointID, VolID, TimepointDateTime) " &
                                     "VALUES (" & DtCohortTimepoint.Rows(j).Item(0) &
                                     ", " & DtVolunteers.Rows(i).Item(0) &
                                     ", '" & DateAdd(DateInterval.Minute, (i * DtCohortTimepoint.Rows(j).Item(2)), CDate(DtCohortTimepoint.Rows(j).Item(1))) & "')")
                j = j + 1
            Loop
            j = 0
            i = i + 1
        Loop
        OverClass.ExecuteMassSQL()
        i = 0
        DtCohortTimepoint = Nothing


        'Insert Schedule into Volunteer Schedule Table

        DtSchedule = OverClass.TempDataTable("SELECT d.StudyScheduleID " &
                                             "FROM ((Cohort a INNER JOIN Study b ON a.StudyID=b.StudyID) " &
                                             "INNER JOIN StudyTimepoint c ON b.StudyID=c.StudyID) " &
                                             "INNER JOIN StudySchedule d ON c.StudyTimepointID=d.StudyTimepointID " &
                                             "WHERE a.CohortID=" & CohortID)
        NumSchedule = DtSchedule.Rows.Count

        Do While j < NumVols
            Do While i < NumSchedule
                OverClass.AddToMassSQL("INSERT INTO VolunteerSchedule (StudyScheduleID, VolID) " &
                                     "VALUES " &
                                     "(" & DtSchedule.Rows(i).Item(0) & ", " & DtVolunteers.Rows(j).Item(0) & ")")
                i = i + 1
            Loop
            i = 0
            j = j + 1
        Loop

        OverClass.ExecuteMassSQL()

        'Update to say Cohort Generated
        OverClass.ExecuteSQL("UPDATE Cohort SET Generated=true WHERE CohortID=" & CohortID)
        Form1.Specifics(Form1.DataGridView7)

        MsgBox("Schedule Generated")


    End Sub


    Public Sub ReSchedule(CohortID As Long, StudyID As Long, CohortName As String, NumVol As Long)

        'Get and insert cohort Timepoints
        Dim StudyTimepoints As DataTable = OverClass.TempDataTable("SELECT StudyTimepointID, TimepointName " &
        "FROM (Study INNER JOIN StudyTimepoint ON Study.StudyID = StudyTimepoint.StudyID) INNER JOIN Cohort ON Study.StudyID = Cohort.StudyID " &
        "WHERE CohortID=" & CohortID)

        Dim VolDetails As DataTable = OverClass.TempDataTable("SELECT * FROM Volunteer WHERE CohortID=" & CohortID)

        OverClass.AddToMassSQL("DELETE * FROM Cohort WHERE CohortID=" & CohortID)
        OverClass.AddToMassSQL("INSERT INTO Cohort (CohortID, StudyID, CohortName, NumVols) " &
        "VALUES(" & CohortID & "," & StudyID & ",'" & CohortName & "', " & NumVol & ")")

        For Each row As DataRow In StudyTimepoints.Rows
            Try
                Dim CohortTimepointDateTime As Date =
                InputBox("Please enter a date/time for cohort timepoint '" & row.Item("TimepointName") & "'",
        row.Item("TimepointName") & " date/time")

                Dim VolGap As Long =
                InputBox("Please enter an interval in minutes between volunteers for timepoint '" & row.Item("TimepointName") & "'",
                row.Item("TimepointName") & " interval")

        OverClass.AddToMassSQL("INSERT INTO CohortTimepoint (CohortID,StudyTimepointID,TimepointDateTime,VolGap) " &
                    "VALUES (" & CohortID & "," & row.Item("StudyTimepointID") & "," & OverClass.SQLDate(CohortTimepointDateTime) & ", " & VolGap & ")")

        Catch ex As Exception
        MsgBox("Incorrect data format entered")
        Exit Sub
        End Try
        Next

        Try
            OverClass.ExecuteMassSQL()
        Catch ex As Exception
            MsgBox("An error occured")
            Exit Sub
        End Try


        Dim DtSchedule As DataTable
        Dim DtCohortTimepoint As DataTable
        Dim DtNumVolunteers As DataTable
        Dim DtVolunteers As DataTable
        Dim i As Long = 0
        Dim j As Long = 0
        Dim NumVols As Long = 0
        Dim NumSchedule As Long = 0

        'Upload into Volunteers Table
        DtNumVolunteers = OverClass.TempDataTable("SELECT NumVols FROM Cohort WHERE CohortID=" & CohortID)

        NumVols = DtNumVolunteers.Rows(0).Item(0)
        DtNumVolunteers = Nothing

        Do While i < NumVols
            Dim VolRVLNo As Long
            Dim VolInitials As String
            Dim VolRoom As Long
            Try
                VolRVLNo = VolDetails.Rows(i).Item("RVLNo")
                VolInitials = VolDetails.Rows(i).Item("Initials")
                VolRoom = VolDetails.Rows(i).Item("RoomNo")
            Catch ex As Exception
                VolRVLNo = 0
                VolInitials = "AAA"
                VolRoom = 0
            End Try
            OverClass.AddToMassSQL("INSERT INTO Volunteer (CohortID, RVLNo, Initials, RoomNo) " &
            "VALUES (" & CohortID & "," & VolRVLNo & ",'" & VolInitials & "'," & VolRoom & ")")
            i = i + 1
        Loop
        OverClass.ExecuteMassSQL()
        i = 0


        'Insert volunteers timepoints into Volunteer Timepoints Table
        DtVolunteers = OverClass.TempDataTable("Select VolID FROM Volunteer WHERE CohortID= " & CohortID)


        DtCohortTimepoint = OverClass.TempDataTable("Select StudyTimepointID, TimepointDateTime, VolGap FROM CohortTimepoint " &
                                              "WHERE CohortID=" & CohortID)

        Do While i < NumVols
            Do While j < DtCohortTimepoint.Rows.Count
                OverClass.AddToMassSQL("INSERT INTO VolunteerTimepoint (StudyTimepointID, VolID, TimepointDateTime) " &
                                     "VALUES (" & DtCohortTimepoint.Rows(j).Item(0) &
                                     ", " & DtVolunteers.Rows(i).Item(0) &
                                     ", '" & DateAdd(DateInterval.Minute, (i * DtCohortTimepoint.Rows(j).Item(2)), CDate(DtCohortTimepoint.Rows(j).Item(1))) & "')")
                j = j + 1
            Loop
            j = 0
            i = i + 1
        Loop
        OverClass.ExecuteMassSQL()
        i = 0
        DtCohortTimepoint = Nothing


        'Insert Schedule into Volunteer Schedule Table

        DtSchedule = OverClass.TempDataTable("SELECT d.StudyScheduleID " &
                                             "FROM ((Cohort a INNER JOIN Study b ON a.StudyID=b.StudyID) " &
                                             "INNER JOIN StudyTimepoint c ON b.StudyID=c.StudyID) " &
                                             "INNER JOIN StudySchedule d ON c.StudyTimepointID=d.StudyTimepointID " &
                                             "WHERE a.CohortID=" & CohortID)
        NumSchedule = DtSchedule.Rows.Count

        Do While j < NumVols
            Do While i < NumSchedule
                OverClass.AddToMassSQL("INSERT INTO VolunteerSchedule (StudyScheduleID, VolID) " &
                                     "VALUES " &
                                     "(" & DtSchedule.Rows(i).Item(0) & ", " & DtVolunteers.Rows(j).Item(0) & ")")
                i = i + 1
            Loop
            i = 0
            j = j + 1
        Loop

        OverClass.ExecuteMassSQL()

        'Update to say Cohort Generated
        OverClass.ExecuteSQL("UPDATE Cohort SET Generated=true WHERE CohortID=" & CohortID)
        Form1.Specifics(Form1.DataGridView7)

        MsgBox("Schedule Generated")


    End Sub

End Module
