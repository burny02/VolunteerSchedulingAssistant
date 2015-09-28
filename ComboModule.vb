Module ComboModule

    Public Sub GenericCombo(sender As Object, e As EventArgs)

        If OverClass.UnloadData() = True Then Exit Sub
        OverClass.ResetCollection()
        Call SubCombo(sender)


    End Sub

    Public Sub SubCombo(sender As ComboBox)

        Select Case sender.Name.ToString

            Case "ComboBox4"
                StartCombo(Form1.ComboBox3)

            Case "ComboBox6"
                StartCombo(Form1.ComboBox7)

            Case "ComboBox9"
                StartCombo(Form1.ComboBox10)

            Case "ComboBox11"
                StartCombo(Form1.ComboBox12)
                StartCombo(Form1.ComboBox13)

            Case "ComboBox12"
                StartCombo(Form1.ComboBox13)

            Case "ComboBox17"
                StartCombo(Form1.ComboBox18)

            Case "ComboBox19", "ComboBox20", "ComboBox14", "ComboBox15", "ComboBox23", "ComboBox24"
                Form1.Specifics(Form1.DataGridView11)
                StartCombo(Form1.ComboBox19)
                StartCombo(Form1.ComboBox20)
                StartCombo(Form1.ComboBox14)
                StartCombo(Form1.ComboBox15)
                StartCombo(Form1.ComboBox23)
                StartCombo(Form1.ComboBox24)


            Case Else
                ComboRefreshData(sender)

        End Select

    End Sub

    Public Sub StartCombo(ctl As ComboBox)

        Select Case ctl.Name.ToString()

            Case "ComboBox17"
                ctl.DataSource = OverClass.TempDataTable("SELECT DISTINCT a.StudyID, " & _
                                                                  "StudyCode FROM (Study a INNER JOIN StudyTimePoint b " & _
                                                                  "ON a.StudyID=b.StudyID) " & _
                                                                  "INNER JOIN Cohort c ON a.StudyID=c.StudyID " & _
                                                                  "WHERE Generated=True " & _
                                                                  "ORDER BY StudyCode ASC")
                ctl.ValueMember = "StudyID"
                ctl.DisplayMember = "StudyCode"

            Case "ComboBox18"
                If IsNothing(Form1.ComboBox17.SelectedValue) Then Exit Sub
                ctl.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                              "CohortName FROM Cohort WHERE StudyID=" _
                                                              & Form1.ComboBox17.SelectedValue.ToString & _
                                                              " AND Generated=True " & _
                                                              " ORDER BY CohortName ASC")
                ctl.ValueMember = "CohortID"
                ctl.DisplayMember = "CohortName"

            Case "ComboBox2"
                ctl.DataSource = OverClass.TempDataTable("SELECT StudyID, " & _
                                                              "StudyCode FROM Study ORDER BY StudyCode ASC")
                ctl.ValueMember = "StudyID"
                ctl.DisplayMember = "StudyCode"

            Case "ComboBox4"
                ctl.DataSource = OverClass.TempDataTable("SELECT DISTINCT a.StudyID, " & _
                             "StudyCode FROM Study a INNER JOIN StudyTimepoint b " & _
                             "ON a.StudyID=b.StudyID " & _
                             "ORDER BY StudyCode ASC")
                ctl.ValueMember = "StudyID"
                ctl.DisplayMember = "StudyCode"

            Case "ComboBox3"
                If IsNothing(Form1.ComboBox4.SelectedValue) Then Exit Sub
                ctl.DataSource = OverClass.TempDataTable("SELECT StudyTimepointID, " & _
                                                              "TimepointName FROM StudyTimepoint WHERE StudyID=" _
                                                              & Form1.ComboBox4.SelectedValue.ToString & _
                                                              " ORDER BY TimepointName ASC")
                ctl.ValueMember = "StudyTimepointID"
                ctl.DisplayMember = "TimepointName"

            Case "ComboBox5"
                ctl.DataSource = OverClass.TempDataTable("SELECT StudyID, " & _
                                                              "StudyCode FROM Study ORDER BY StudyCode ASC")
                ctl.ValueMember = "StudyID"
                ctl.DisplayMember = "StudyCode"

            Case "ComboBox6"
                ctl.DataSource = OverClass.TempDataTable("SELECT DISTINCT a.StudyID, " & _
                                                              "StudyCode FROM (Study a INNER JOIN StudyTimePoint b " & _
                                                              "ON a.StudyID=b.StudyID) " & _
                                                              "INNER JOIN Cohort c ON a.StudyID=c.StudyID " & _
                                                              "WHERE Generated=False " & _
                                                              "ORDER BY StudyCode ASC")
                ctl.ValueMember = "StudyID"
                ctl.DisplayMember = "StudyCode"

            Case "ComboBox7"
                If IsNothing(Form1.ComboBox6.SelectedValue) Then Exit Sub
                ctl.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                              "CohortName FROM Cohort WHERE StudyID=" _
                                                              & Form1.ComboBox6.SelectedValue.ToString & _
                                                              " AND Generated=False" & _
                                                              " ORDER BY CohortName ASC")
                ctl.ValueMember = "CohortID"
                ctl.DisplayMember = "CohortName"

            Case "ComboBox8"
                ctl.DataSource = OverClass.TempDataTable("SELECT a.CohortID, StudyCode & ' - ' & CohortName AS Display " & _
                                                              "FROM (SELECT StudyCode, CohortName, CohortID, " & _
                                                              "Count(StudyTimepointID) as NumTimepoint " & _
                                                              "FROM (Study a INNER JOIN StudyTimePoint b " & _
                                                              "ON a.StudyID=b.StudyID) " & _
                                                              "INNER JOIN Cohort c ON a.StudyID=c.StudyID " & _
                                                              "GROUP BY StudyCode, CohortName, CohortID) as a " & _
                                                              "INNER JOIN " & _
                                                              "(SELECT c.CohortID, Count(CohortTimepointID) as NumTimepoint " & _
                                                              "FROM CohortTimepoint c INNER JOIN Cohort d " & _
                                                              "ON c.CohortID=d.CohortID WHERE Generated=False " & _
                                                              "GROUP BY c.CohortID) as b " & _
                                                              "ON a.CohortID=b.CohortID AND a.NumTimepoint=b.NumTimepoint")
                ctl.ValueMember = "CohortID"
                ctl.DisplayMember = "Display"

            Case "ComboBox9"
                ctl.DataSource = OverClass.TempDataTable("SELECT DISTINCT a.StudyID, " & _
                                                              "StudyCode FROM (Study a INNER JOIN StudyTimePoint b " & _
                                                              "ON a.StudyID=b.StudyID) " & _
                                                              "INNER JOIN Cohort c ON a.StudyID=c.StudyID " & _
                                                              "WHERE Generated=True " & _
                                                              "ORDER BY StudyCode ASC")
                ctl.ValueMember = "StudyID"
                ctl.DisplayMember = "StudyCode"

            Case "ComboBox10"
                If IsNothing(Form1.ComboBox9.SelectedValue) Then Exit Sub
                ctl.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                              "CohortName FROM Cohort WHERE StudyID=" _
                                                              & Form1.ComboBox9.SelectedValue.ToString & _
                                                                " AND Generated=True" & _
                                                                " ORDER BY CohortName ASC")
                ctl.ValueMember = "CohortID"
                ctl.DisplayMember = "CohortName"

            Case "ComboBox11"
                ctl.DataSource = OverClass.TempDataTable("SELECT DISTINCT a.StudyID, " & _
                                                              "StudyCode FROM (Study a INNER JOIN StudyTimePoint b " & _
                                                              "ON a.StudyID=b.StudyID) " & _
                                                              "INNER JOIN Cohort c ON a.StudyID=c.StudyID " & _
                                                              "WHERE Generated=True " & _
                                                              "ORDER BY StudyCode ASC")
                ctl.ValueMember = "StudyID"
                ctl.DisplayMember = "StudyCode"

            Case "ComboBox12"
                If IsNothing(Form1.ComboBox11.SelectedValue) Then Exit Sub
                ctl.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                              "CohortName FROM Cohort WHERE StudyID=" _
                                                              & Form1.ComboBox11.SelectedValue.ToString & _
                                                                " AND Generated=True" & _
                                                                " ORDER BY CohortName ASC")
                ctl.ValueMember = "CohortID"
                ctl.DisplayMember = "CohortName"

            Case "ComboBox13"
                If IsNothing(Form1.ComboBox12.SelectedValue) Then Exit Sub
                ctl.DataSource = OverClass.TempDataTable("SELECT RVLNo & ' - ' & Initials AS Display, VolID " & _
                                                              "FROM Volunteer WHERE CohortID=" _
                                                              & Form1.ComboBox12.SelectedValue.ToString & _
                                                                " ORDER BY Initials ASC")
                ctl.ValueMember = "VolID"
                ctl.DisplayMember = "Display"

            Case "ComboBox14"

                If ctl.SelectedValue <> "" Then Exit Sub
                Dim Fielder As String = "StudyCode"

                Dim dt As DataTable = OverClass.CurrentDataSet.Tables(0)

                Dim QueryResult = (From a In dt.AsEnumerable() _
                                    Select "").Union _
                                    (From a In dt.AsEnumerable()
                                     Order By a.Field(Of String)(Fielder) Ascending
                                    Select a.Field(Of String)(Fielder)).Distinct()

                Dim dt2 As New DataTable
                dt2.Columns.Add(Fielder)
                For Each row In QueryResult
                    dt2.Rows.Add(row)
                Next

                ctl.DataSource = dt2
                ctl.DisplayMember = Fielder
                ctl.ValueMember = Fielder


            Case "ComboBox15"

                If ctl.SelectedValue <> "" Then Exit Sub
                Dim Fielder As String = "CohortName"

                Dim dt As DataTable = OverClass.CurrentDataSet.Tables(0)

                Dim QueryResult = (From a In dt.AsEnumerable() _
                                    Select "").Union _
                                    (From a In dt.AsEnumerable()
                                     Order By a.Field(Of String)(Fielder) Ascending
                                    Select a.Field(Of String)(Fielder)).Distinct()

                Dim dt2 As New DataTable
                dt2.Columns.Add(Fielder)
                For Each row In QueryResult
                    dt2.Rows.Add(row)
                Next

                ctl.DataSource = dt2
                ctl.DisplayMember = Fielder
                ctl.ValueMember = Fielder

            Case "ComboBox19"

                If ctl.SelectedValue <> "" Then Exit Sub
                Dim Fielder As String = "ProcName"

                Dim dt As DataTable = OverClass.CurrentDataSet.Tables(0)

                Dim QueryResult = (From a In dt.AsEnumerable() _
                                    Select "").Union _
                                    (From a In dt.AsEnumerable()
                                     Order By a.Field(Of String)(Fielder) Ascending
                                    Select a.Field(Of String)(Fielder)).Distinct()

                Dim dt2 As New DataTable
                dt2.Columns.Add(Fielder)
                For Each row In QueryResult
                    dt2.Rows.Add(row)
                Next

                ctl.DataSource = dt2
                ctl.DisplayMember = Fielder
                ctl.ValueMember = Fielder

            Case "ComboBox20"

                If ctl.SelectedValue <> "" Then Exit Sub
                Dim Fielder As String = "Vol"

                Dim dt As DataTable = OverClass.CurrentDataSet.Tables(0)

                Dim QueryResult = (From a In dt.AsEnumerable() _
                                    Select "").Union _
                                    (From a In dt.AsEnumerable()
                                     Order By a.Field(Of String)(Fielder) Ascending
                                    Select a.Field(Of String)(Fielder)).Distinct()

                Dim dt2 As New DataTable
                dt2.Columns.Add(Fielder)
                For Each row In QueryResult
                    dt2.Rows.Add(row)
                Next

                ctl.DataSource = dt2
                ctl.DisplayMember = Fielder
                ctl.ValueMember = Fielder

            Case "ComboBox23"

                If ctl.SelectedValue <> "" Then Exit Sub
                Dim Fielder As String = "CalcDate"

                Dim dt As DataTable = OverClass.CurrentDataSet.Tables(0)

                Dim QueryResult = (From a In dt.AsEnumerable() _
                                    Select "").Union _
                                    (From a In dt.AsEnumerable()
                                     Order By a.Field(Of DateTime)(Fielder) Ascending
                                    Select Format(a.Field(Of DateTime)(Fielder), "dd-MMM-yyyy")).Distinct()

                Dim dt2 As New DataTable
                dt2.Columns.Add(Fielder)
                For Each row In QueryResult
                    dt2.Rows.Add(row)
                Next

                ctl.DataSource = dt2
                ctl.DisplayMember = Fielder
                ctl.ValueMember = Fielder

            Case "ComboBox24"

                If ctl.SelectedValue <> "" Then Exit Sub
                Dim Fielder As String = "FullName"

                Dim dt As DataTable = OverClass.CurrentDataSet.Tables(0)

                Dim QueryResult = ((From a In dt.AsEnumerable() _
                                    Select "").Union _
                                    (From a In dt.AsEnumerable() _
                                    Select "UnAssigned").Union _
                                    (From a In dt.AsEnumerable()
                                     Where a.Field(Of String)(Fielder) <> ""
                                     Where a.Field(Of String)(Fielder) <> Nothing
                                        Where a.Field(Of String)(Fielder) <> " "
                                     Order By a.Field(Of String)(Fielder) Ascending
                                    Select a.Field(Of String)(Fielder))).Distinct()

                Dim dt2 As New DataTable
                dt2.Columns.Add(Fielder)
                For Each row In QueryResult
                    dt2.Rows.Add(row)
                Next

                ctl.DataSource = dt2
                ctl.DisplayMember = Fielder
                ctl.ValueMember = Fielder


        End Select

        ComboRefreshData(ctl)

    End Sub

    Public Sub ComboRefreshData(sender As ComboBox)

        Dim Grid As DataGridView = Nothing

        Select Case sender.Name.ToString()

            Case "ComboBox13"
                Grid = Form1.DataGridView10

            Case "ComboBox10"
                Grid = Form1.DataGridView9

            Case "ComboBox7"
                Grid = Form1.DataGridView8

            Case "ComboBox5"
                Grid = Form1.DataGridView7

            Case "ComboBox3"
                Grid = Form1.DataGridView6

            Case "ComboBox2"
                Grid = Form1.DataGridView5


        End Select


        If Not IsNothing(Grid) Then Call Form1.Specifics(Grid)

    End Sub

End Module
