Module OverlapModule
    Public Function CheckVolunteerOverlap(StaffMember As String, ID As Long, StartFull As Date, EndFull As Date,
                                          DT As DataTable) As String

        Try
            Dim Returner As String = vbNullString


            Dim CDateStart As String = vbNullString
            Dim CDateEnd As String = vbNullString
            Dim chk As String = vbNullString


            CDateStart = OverClass.SQLDate(StartFull)
            CDateEnd = OverClass.SQLDate(EndFull)

            'Vol Procedures

            For Each row As DataRow In DT.Rows

                If row.RowState = DataRowState.Deleted Then Continue For
                If IsDBNull(row.Item("StaffID")) _
                Or IsNothing(row.Item("StaffID")) Then Continue For
                If (row.Item("StaffID")) <> StaffMember Then Continue For
                If (row.Item("VolunteerScheduleID")) = ID Then Continue For

                Dim RowEndFull, RowStartFull As Date
                RowEndFull = row.Item("EndFull")
                RowStartFull = row.Item("CalcDate")

                If ((RowStartFull < EndFull) _
                And (StartFull < RowEndFull)) _
                Or ((RowStartFull = StartFull) _
                And (RowEndFull = EndFull)) Then

                    chk = chk & "Vol Procedure (" & Format(RowStartFull, "HH:mm") & "-" &
                Format(RowEndFull, "HH:mm") & ")" & " - " & row.Item("ProcName") & ","

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

    Public Function CheckExtraOverlap(StaffMember As String, ID As Long, StartFull As Date, EndFull As Date,
                                          Grid As DataGridView, PassedRowIndex As Long) As String
        Try
            Dim Returner As String = vbNullString
            Dim chk As String = vbNullString

            'Vol Procedures
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
