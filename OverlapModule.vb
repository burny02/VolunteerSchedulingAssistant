Module OverlapModule
    Public Function CheckVolunteerOverlap(StaffMember As String, ID As Long, StartFull As Date, EndFull As Date, _
                                          Grid As DataGridView, Optional OnlyFrontEnd As Boolean = False) As String

        Try
            Dim Returner As String = vbNullString


            Dim CDateStart As String = vbNullString
            Dim CDateEnd As String = vbNullString
            Dim chk As String = vbNullString


            CDateStart = OverClass.SQLDate(StartFull)
            CDateEnd = OverClass.SQLDate(EndFull)

            If OnlyFrontEnd <> True Then
                'Vol Procedures

                Dim SqlString As String

                SqlString = "SELECT ProcType & '(' & format(StartFull,'hh:nn') & '-' & format(EndFull,'hh:nn') & ')' " &
            "& ' - ' & ProcName as Overlap " &
            "FROM [CheckStaffOverlap] WHERE " &
            "([ID]<>" & ID & " AND [StaffID]=" & StaffMember & " AND [StartFull]<" & CDateEnd &
            " AND " & CDateStart & "<[EndFull])" &
            " OR ([ID]<>" & ID & " AND [StaffID]=" & StaffMember & " AND [StartFull]=" & CDateStart & " AND [EndFull]=" & CDateEnd & ")"

                Returner = OverClass.CreateCSVString(SqlString)

            End If

            For Each row In Grid.Rows

                If IsDBNull(row.cells("StaffID").value) _
                Or IsNothing(row.cells("StaffID").value) Then Continue For
                If (row.cells("StaffID").value) <> StaffMember Then Continue For
                If (row.cells("VolunteerScheduleID").value) = ID Then Continue For

                Dim RowEndFull, RowStartFull As Date
                RowEndFull = row.cells("EndFull").value
                RowStartFull = row.cells("CalcDate").value

                If ((RowStartFull < EndFull) _
                And (StartFull < RowEndFull)) _
                Or ((RowStartFull = StartFull) _
                And (RowEndFull = EndFull)) Then

                    chk = chk & "Vol Procedure (" & Format(RowStartFull, "HH:mm") & "-" &
                Format(RowEndFull, "HH:mm") & ")" & " - " & row.cells("ProcName").value & ","

                End If

            Next

            If chk <> vbNullString Then chk = Left(chk, Len(chk) - 1)
            If Returner <> vbNullString And chk <> vbNullString Then Returner = Returner & ","

            CheckVolunteerOverlap = Returner & chk
        Catch ex As Exception
            MsgBox(ex.Message)
        Throw
        End Try

    End Function

    Public Function CheckExtraOverlap(StaffMember As String, ID As Long, StartFull As Date, EndFull As Date, _
                                          Grid As DataGridView, PassedRowIndex As Long) As String
        Try
            Dim Returner As String = vbNullString


            Dim CDateStart As String = vbNullString
            Dim CDateEnd As String = vbNullString
            Dim chk As String = vbNullString


            CDateStart = OverClass.SQLDate(StartFull)
            CDateEnd = OverClass.SQLDate(EndFull)

            'Vol Procedures

            Dim SqlString As String

            SqlString = "SELECT ProcType & '(' & format(StartFull,'hh:nn') & '-' & format(EndFull,'hh:nn') & ')' " &
        "& ' - ' & ProcName as Overlap " &
        "FROM [CheckStaffOverlap] WHERE " &
        "([ID]<>" & ID & " AND [StaffID]=" & StaffMember & " AND [StartFull]<" & CDateEnd &
        " AND " & CDateStart & "<[EndFull])" &
        " OR ([ID]<>" & ID & " AND [StaffID]=" & StaffMember & " AND [StartFull]=" & CDateStart & " AND [EndFull]=" & CDateEnd & ")"


            Returner = OverClass.CreateCSVString(SqlString)

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

                If ((RowStartFull < EndFull) _
                And (StartFull < RowEndFull)) _
                Or ((RowStartFull = StartFull) _
                And (RowEndFull = EndFull)) Then

                    chk = chk & "Staff Procedure (" & Format(RowStartFull, "HH:mm") & "-" &
                Format(RowEndFull, "HH:mm") & ")" & " - " & Row.Cells("ProcPick").FormattedValue & ","

                End If

            Next

            If chk <> vbNullString Then chk = Left(chk, Len(chk) - 1)
            If Returner <> vbNullString And chk <> vbNullString Then Returner = Returner & ","

            CheckExtraOverlap = Returner & chk
        Catch ex As Exception
            MsgBox(ex.Message)
            Throw
        End Try

    End Function

    Public Function ScheduleOverlap(Grid As DataGridView, PassedRowIndex As Long, _
                            NumDays As Long, ProcTime As Date, NumTaken As Long) As String

        Try
            Dim chk As String = vbNullString
            Dim DefaultTime As Date = "#12:00#"
            Dim CalculationDate As Date = "#01/01/2000#"
            Dim StartFull, EndFull As Date

            StartFull = DateAdd(DateInterval.Minute,
                    DateDiff(DateInterval.Minute, TimeValue(DefaultTime), TimeValue(ProcTime)),
                    DateAdd(DateInterval.Day, NumDays, CalculationDate))

            EndFull = DateAdd(DateInterval.Minute, NumTaken, StartFull)



            For Each Row In Grid.Rows

                If PassedRowIndex = Row.index Then Continue For
                If IsDBNull(Row.Cells("MinsTaken").Value) Then Continue For
                If IsDBNull(Row.Cells("DaysPost").Value) Then Continue For
                If IsDBNull(Row.Cells("ProcTime").Value) Then Continue For
                If Row.Cells("Approx").Value = "Set Time" Then Continue For
                If Row.Cells("MinsTaken").formattedvalue = vbNullString Then Continue For


                Dim RowEndFull, RowStartFull As Date
                Dim DaysPost, MinsTaken As Long
                Dim RowTime As Date
                RowTime = TimeValue(CDate(Row.Cells("ProcTime").Value))
                DaysPost = Row.Cells("DaysPost").Value
                MinsTaken = Row.Cells("MinsTaken").formattedvalue

                RowStartFull = DateAdd(DateInterval.Minute,
                    DateDiff(DateInterval.Minute, TimeValue(DefaultTime), TimeValue(RowTime)),
                    DateAdd(DateInterval.Day, DaysPost, CalculationDate))


                RowEndFull = DateAdd(DateInterval.Minute, MinsTaken, RowStartFull)

                If ((RowStartFull < EndFull) _
                And (StartFull < RowEndFull)) _
                Or ((RowStartFull = StartFull) _
                And (RowEndFull = EndFull)) Then

                    chk = chk & Row.Cells("PickProc").FormattedValue & ","

                End If

            Next

            If chk <> vbNullString Then chk = Left(chk, Len(chk) - 1)

            ScheduleOverlap = chk

        Catch ex As Exception
            MsgBox(ex.Message)
        Throw
        End Try

    End Function

End Module
