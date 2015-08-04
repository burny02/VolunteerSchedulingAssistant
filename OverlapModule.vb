Module OverlapModule
    Public Function CheckVolunteerOverlap(StaffMember As String, ID As Long, StartFull As Date, EndFull As Date, _
                                          Grid As DataGridView, Optional OnlyFrontEnd As Boolean = False) As String

        Dim Returner As String = vbNullString


        Dim CDateStart As String = vbNullString
        Dim CDateEnd As String = vbNullString
        Dim chk As String = vbNullString


        CDateStart = OverClass.SQLDate(StartFull)
        CDateEnd = OverClass.SQLDate(EndFull)

        If OnlyFrontEnd <> True Then
            'Vol Procedures
            Returner = OverClass.CreateCSVString("SELECT ProcType & '(' & format(StartFull,'hh:nn') & '-' & format(EndFull,'hh:nn') & ')' " & _
            "& ' - ' & ProcName as Overlap " & _
            "FROM [CheckStaffOverlap] WHERE ([ID]<>" & ID & " AND [StaffID]=" & StaffMember & " AND [StartFull] > " & CDateStart & " AND [StartFull] < " & CDateEnd & ") " & _
            "OR ([ID]<>" & ID & " AND [StaffID]=" & StaffMember & " AND [EndFull] > " & CDateStart & " AND [EndFull] < " & CDateEnd & ") " & _
            "OR ([ID]<>" & ID & " AND [StaffID]=" & StaffMember & " AND [StartFull]<=" & CDateStart & " AND [EndFull]>=" & CDateEnd & ") " & _
            "OR ([ID]<>" & ID & " AND [StaffID]=" & StaffMember & " AND [StartFull]=" & CDateStart & ") " & _
            "OR ([ID]<>" & ID & " AND [StaffID]=" & StaffMember & " AND [EndFull]=" & CDateEnd & ")")
        End If

        For Each row In Grid.Rows

            If IsDBNull(row.cells("StaffID").value) _
                Or IsNothing(row.cells("StaffID").value) Then Continue For
            If (row.cells("StaffID").value) <> StaffMember Then Continue For
            If (row.cells("VolunteerScheduleID").value) = ID Then Continue For

            Dim RowEndFull, RowStartFull As Date
            RowEndFull = row.cells("EndFull").value
            RowStartFull = row.cells("CalcDate").value

            If (RowStartFull > StartFull And RowStartFull < EndFull) _
                Or (RowEndFull > StartFull And RowEndFull < EndFull) _
                Or (RowStartFull <= StartFull And RowEndFull >= EndFull) _
                Or (RowStartFull = StartFull) _
                Or (RowEndFull = EndFull) Then

                chk = chk & "Vol Procedure (" & Format(RowStartFull, "HH:mm") & "-" & _
                Format(RowEndFull, "HH:mm") & ")" & " - " & row.cells("ProcName").value & ","

            End If

        Next

        If chk <> vbNullString Then chk = Left(chk, Len(chk) - 1)
        If Returner <> vbNullString And chk <> vbNullString Then Returner = Returner & ","

        CheckVolunteerOverlap = Returner & chk

    End Function

    Public Function CheckExtraOverlap(StaffMember As String, ID As Long, StartFull As Date, EndFull As Date, _
                                          Grid As DataGridView, PassedRowIndex As Long) As String

        Dim Returner As String = vbNullString


        Dim CDateStart As String = vbNullString
        Dim CDateEnd As String = vbNullString
        Dim chk As String = vbNullString


        CDateStart = OverClass.SQLDate(StartFull)
        CDateEnd = OverClass.SQLDate(EndFull)

        'Vol Procedures
        Returner = OverClass.CreateCSVString("SELECT ProcType & '(' & format(StartFull,'hh:nn') & '-' & format(EndFull,'hh:nn') & ')' " & _
        "& ' - ' & ProcName as Overlap " & _
        "FROM [CheckStaffOverlap] WHERE ProcType = 'Vol Procedure' AND " & _
        "([ID]<>" & ID & " AND [StaffID]=" & StaffMember & " AND [StartFull] > " & CDateStart & " AND [StartFull] < " & CDateEnd & ") " & _
        "OR ([ID]<>" & ID & " AND [StaffID]=" & StaffMember & " AND [EndFull] > " & CDateStart & " AND [EndFull] < " & CDateEnd & ") " & _
        "OR ([ID]<>" & ID & " AND [StaffID]=" & StaffMember & " AND [StartFull]<=" & CDateStart & " AND [EndFull]>=" & CDateEnd & ") " & _
        "OR ([ID]<>" & ID & " AND [StaffID]=" & StaffMember & " AND [StartFull]=" & CDateStart & ") " & _
        "OR ([ID]<>" & ID & " AND [StaffID]=" & StaffMember & " AND [EndFull]=" & CDateEnd & ")")


        For Each Row In Grid.Rows


            If IsDBNull(Row.Cells("StaffID").Value) _
                Or IsNothing(Row.Cells("StaffID").Value) Then Continue For
            If (Row.Cells("StaffID").Value) <> StaffMember Then Continue For
            If PassedRowIndex = Row.index Then Continue For
            If IsDBNull(Row.Cells("CalcDate").Value) _
                Or IsNothing(Row.Cells("CalcDate").Value) Then Continue For
            If IsDBNull(Row.Cells("MinsTaken").Value) _
                Or IsNothing(Row.Cells("MinsTaken").Value) Then Continue For
            If IsDBNull(Row.Cells("ProcID").Value) _
                Or IsNothing(Row.Cells("ProcID").Value) Then Continue For


            Dim RowEndFull, RowStartFull As Date
            RowEndFull = DateAdd(DateInterval.Minute, Row.Cells("MinsTaken").FormattedValue, Row.Cells("CalcDate").Value)
            RowStartFull = Row.Cells("CalcDate").Value

            If (RowStartFull > StartFull And RowStartFull < EndFull) _
                Or (RowEndFull > StartFull And RowEndFull < EndFull) _
                Or (RowStartFull <= StartFull And RowEndFull >= EndFull) _
                Or (RowStartFull = StartFull) _
                Or (RowEndFull = EndFull) Then

                chk = chk & "Staff Procedure (" & Format(RowStartFull, "HH:mm") & "-" & _
                Format(RowEndFull, "HH:mm") & ")" & " - " & Row.Cells("ProcPick").FormattedValue & ","

            End If

        Next

        If chk <> vbNullString Then chk = Left(chk, Len(chk) - 1)
        If Returner <> vbNullString And chk <> vbNullString Then Returner = Returner & ","

        CheckExtraOverlap = Returner & chk

    End Function

    Public Function ScheduleOverlap(Grid As DataGridView, PassedRowIndex As Long, _
                            NumDays As Long, NumHours As Long, NumMins As Long, NumTaken As Long) As String

        Dim chk As String = vbNullString
        Dim CalculationDate As Date = "#01/01/2000#"
        Dim StartFull, EndFull As Date

        StartFull = DateAdd(DateInterval.Minute, NumMins, _
                                   (DateAdd(DateInterval.Hour, NumHours, _
                                   DateAdd(DateInterval.Day, NumDays, CalculationDate))))
        EndFull = DateAdd(DateInterval.Minute, NumTaken, StartFull)


        For Each Row In Grid.Rows

            If IsDBNull(Row.Cells("ProcID").Value) _
                Or IsNothing(Row.Cells("ProcID").Value) Then Continue For
            If PassedRowIndex = Row.index Then Continue For
            If IsDBNull(Row.Cells("MinsTaken").Value) _
                Or IsNothing(Row.Cells("MinsTaken").Value) Then Continue For
            If IsDBNull(Row.Cells("ProcID").Value) _
                Or IsNothing(Row.Cells("ProcID").Value) Then Continue For
            If IsDBNull(Row.Cells("DaysPost").Value) _
                Or IsNothing(Row.Cells("DaysPost").Value) Then Continue For
            If IsDBNull(Row.Cells("HoursPost").Value) _
                Or IsNothing(Row.Cells("HoursPost").Value) Then Continue For
            If IsDBNull(Row.Cells("MinsPost").Value) _
                Or IsNothing(Row.Cells("MinsPost").Value) Then Continue For
            If Row.Cells("Approx").Value = "Set Time" Then Continue For


            Dim RowEndFull, RowStartFull As Date
            Dim MinsPost, HoursPost, DaysPost, MinsTaken As Long
            MinsPost = Row.Cells("MinsPost").Value
            HoursPost = Row.Cells("HoursPost").Value
            DaysPost = Row.Cells("DaysPost").Value
            MinsTaken = Row.Cells("MinsTaken").FormattedValue

            RowStartFull = DateAdd(DateInterval.Minute, MinsPost, _
                                   (DateAdd(DateInterval.Hour, HoursPost, _
                                   DateAdd(DateInterval.Day, DaysPost, CalculationDate))))
            RowEndFull = DateAdd(DateInterval.Minute, MinsTaken, RowStartFull)


            If (RowStartFull > StartFull And RowStartFull < EndFull) _
                Or (RowEndFull > StartFull And RowEndFull < EndFull) _
                Or (RowStartFull <= StartFull And RowEndFull >= EndFull) _
                Or (RowStartFull = StartFull) _
                Or (RowEndFull = EndFull) Then

                chk = chk & Row.Cells("PickProc").FormattedValue & ","

            End If

        Next

        If chk <> vbNullString Then chk = Left(chk, Len(chk) - 1)

        ScheduleOverlap = chk

    End Function

End Module
