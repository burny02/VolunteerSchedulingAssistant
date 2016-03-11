Module ScheduleGenerator
    Public Sub Generator()

        If IsNothing(Form1.FilterCombo13.SelectedValue) Then Exit Sub

        Dim DtSchedule As DataTable
        Dim DtCohortTimepoint As DataTable
        Dim DtNumVolunteers As DataTable
        Dim DtVolunteers As DataTable
        Dim i As Long = 0
        Dim j As Long = 0
        Dim NumVols As Long = 0
        Dim NumSchedule As Long = 0

        'Upload into Volunteers Table
        DtNumVolunteers = OverClass.TempDataTable("SELECT NumVols FROM Cohort WHERE CohortID=" & Form1.FilterCombo13.SelectedValue.ToString)

        NumVols = DtNumVolunteers.Rows(0).Item(0)
        DtNumVolunteers = Nothing

        Do While i < NumVols
            OverClass.AddToMassSQL("INSERT INTO Volunteer (CohortID) VALUES (" & Form1.FilterCombo13.SelectedValue.ToString & ")")
            i = i + 1
        Loop
        OverClass.ExecuteMassSQL()
        i = 0


        'Insert volunteers timepoints into Volunteer Timepoints Table
        DtVolunteers = OverClass.TempDataTable("SELECT VolID FROM Volunteer WHERE CohortID=" & Form1.FilterCombo13.SelectedValue.ToString)


        DtCohortTimepoint = OverClass.TempDataTable("SELECT StudyTimepointID, TimepointDateTime, VolGap FROM CohortTimepoint " &
                                              "WHERE CohortID=" & Form1.FilterCombo13.SelectedValue.ToString)

        Do While i < NumVols
            Do While j < DtCohortTimepoint.Rows.Count
                OverClass.AddToMassSQL("INSERT INTO VolunteerTimepoint (StudyTimepointID, VolID, TimepointDateTime) " & _
                                     "VALUES (" & DtCohortTimepoint.Rows(j).Item(0) & _
                                     ", " & DtVolunteers.Rows(i).Item(0) & _
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
                                             "WHERE a.CohortID=" & Form1.FilterCombo13.SelectedValue.ToString)
        NumSchedule = DtSchedule.Rows.Count

        Do While j < NumVols
            Do While i < NumSchedule
                OverClass.AddToMassSQL("INSERT INTO VolunteerSchedule (StudyScheduleID, VolID) " & _
                                     "VALUES " & _
                                     "(" & DtSchedule.Rows(i).Item(0) & ", " & DtVolunteers.Rows(j).Item(0) & ")")
                i = i + 1
            Loop
            i = 0
            j = j + 1
        Loop

        OverClass.ExecuteMassSQL()

        'Update to say Cohort Generated
        OverClass.ExecuteSQL("UPDATE Cohort SET Generated=true WHERE CohortID=" & Form1.FilterCombo13.SelectedValue.ToString)

        Form1.FilterCombo13.DataSource = OverClass.TempDataTable("SELECT a.CohortID, StudyCode & ' - ' & CohortName AS Display " &
                                                              "FROM (SELECT StudyCode, CohortName, CohortID, " &
                                                              "Count(StudyTimepointID) as NumTimepoint " &
                                                              "FROM (Study a INNER JOIN StudyTimePoint b " &
                                                              "ON a.StudyID=b.StudyID) " &
                                                              "INNER JOIN Cohort c ON a.StudyID=c.StudyID " &
                                                              "GROUP BY StudyCode, CohortName, CohortID) as a " &
                                                              "INNER JOIN " &
                                                              "(SELECT c.CohortID, Count(CohortTimepointID) as NumTimepoint " &
                                                              "FROM CohortTimepoint c INNER JOIN Cohort d " &
                                                              "ON c.CohortID=d.CohortID WHERE Generated=False " &
                                                              "GROUP BY c.CohortID) as b " &
                                                              "ON a.CohortID=b.CohortID AND a.NumTimepoint=b.NumTimepoint")
        Form1.FilterCombo13.ValueMember = "CohortID"
        Form1.FilterCombo13.DisplayMember = "Display"

        MsgBox("Schedule Generated")
        Form1.FilterCombo13.RefreshCombo()

    End Sub
End Module
