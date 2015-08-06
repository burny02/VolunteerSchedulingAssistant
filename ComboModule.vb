﻿Module ComboModule

    Public Sub GenericCombo(sender As Object, e As EventArgs)

        If sender.SelectedValue.ToString = vbNullString Then Exit Sub

        If OverClass.UnloadData() = True Then Exit Sub
        OverClass.ResetCollection()
        Call SubCombo(sender)


    End Sub

    Private Sub SubCombo(sender As ComboBox)

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

            Case "ComboBox14"
                StartCombo(Form1.ComboBox15)

            Case "ComboBox1"
                StartCombo(Form1.ComboBox16)

            Case "ComboBox17"
                StartCombo(Form1.ComboBox18)

            Case "ComboBox19"
                StartCombo(Form1.ComboBox20)
                StartCombo(Form1.ComboBox21)
                StartCombo(Form1.ComboBox22)

            Case "ComboBox20"
                StartCombo(Form1.ComboBox21)
                StartCombo(Form1.ComboBox22)

            Case "ComboBox21"
                StartCombo(Form1.ComboBox22)

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
                ctl.DataSource = OverClass.TempDataTable("SELECT DISTINCT a.StudyID, " & _
                                                              "StudyCode FROM (Study a INNER JOIN StudyTimePoint b " & _
                                                              "ON a.StudyID=b.StudyID) " & _
                                                              "INNER JOIN Cohort c ON a.StudyID=c.StudyID " & _
                                                              "WHERE Generated=True " & _
                                                              "ORDER BY StudyCode ASC")
                ctl.ValueMember = "StudyID"
                ctl.DisplayMember = "StudyCode"

            Case "ComboBox15"
                If IsNothing(Form1.ComboBox14.SelectedValue) Then Exit Sub
                ctl.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                              "CohortName FROM Cohort WHERE StudyID=" _
                                                              & Form1.ComboBox14.SelectedValue.ToString & _
                                                              " AND Generated=True " & _
                                                              " ORDER BY CohortName ASC")
                ctl.ValueMember = "CohortID"
                ctl.DisplayMember = "CohortName"

            Case "ComboBox1"
                ctl.DataSource = OverClass.TempDataTable("SELECT DISTINCT a.StudyID, " & _
                                                              "StudyCode FROM (Study a INNER JOIN StudyTimePoint b " & _
                                                              "ON a.StudyID=b.StudyID) " & _
                                                              "INNER JOIN Cohort c ON a.StudyID=c.StudyID " & _
                                                              "WHERE Generated=True " & _
                                                              "ORDER BY StudyCode ASC")
                ctl.ValueMember = "StudyID"
                ctl.DisplayMember = "StudyCode"

            Case "ComboBox16"
                If IsNothing(Form1.ComboBox1.SelectedValue) Then Exit Sub
                ctl.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                              "CohortName FROM Cohort WHERE StudyID=" _
                                                              & Form1.ComboBox1.SelectedValue.ToString & _
                                                              " AND Generated=True " & _
                                                              " ORDER BY CohortName ASC")
                ctl.ValueMember = "CohortID"
                ctl.DisplayMember = "CohortName"

            Case "ComboBox19"
                ctl.DataSource = OverClass.TempDataTable("SELECT DISTINCT a.StudyID, " & _
                                                              "StudyCode FROM (Study a INNER JOIN StudyTimePoint b " & _
                                                              "ON a.StudyID=b.StudyID) " & _
                                                              "INNER JOIN Cohort c ON a.StudyID=c.StudyID " & _
                                                              "WHERE Generated=True " & _
                                                              "ORDER BY StudyCode ASC")
                ctl.ValueMember = "StudyID"
                ctl.DisplayMember = "StudyCode"

            Case "ComboBox20"
                If IsNothing(Form1.ComboBox19.SelectedValue) Then Exit Sub
                ctl.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                              "CohortName FROM Cohort WHERE StudyID=" _
                                                              & Form1.ComboBox19.SelectedValue.ToString & _
                                                                " AND Generated=True" & _
                                                                " ORDER BY CohortName ASC")
                ctl.ValueMember = "CohortID"
                ctl.DisplayMember = "CohortName"

            Case "ComboBox21"
                If IsNothing(Form1.ComboBox20.SelectedValue) Then Exit Sub
                ctl.DataSource = OverClass.TempDataTable("SELECT RVLNo & ' - ' & Initials AS Display, VolID " & _
                                                              "FROM Volunteer WHERE CohortID=" _
                                                              & Form1.ComboBox20.SelectedValue.ToString & _
                                                                " ORDER BY Initials ASC")
                ctl.ValueMember = "VolID"
                ctl.DisplayMember = "Display"

            Case "ComboBox22"
                If IsNothing(Form1.ComboBox21.SelectedValue) Then Exit Sub
                ctl.DataSource = OverClass.TempDataTable("SELECT TimepointName, a.StudyTimepointID " & _
                                                              "FROM StudyTimepoint a INNER JOIN VolunteerTimepoint b " & _
                                                              "ON a.StudyTimepointID=b.StudyTimepointID " & _
                                                              "WHERE VolID=" & Form1.ComboBox21.SelectedValue.ToString & _
                                                                " ORDER BY TimepointName ASC")
                ctl.ValueMember = "StudyTimepointID"
                ctl.DisplayMember = "TimepointName"

        End Select

        ComboRefreshData(ctl)

    End Sub

    Public Sub ComboRefreshData(sender As ComboBox)

        Dim Grid As DataGridView = Nothing

        Select Case sender.Name.ToString()

            Case "ComboBox22"
                Grid = Form1.DataGridView13

            Case "ComboBox16"
                Grid = Form1.DataGridView4

            Case "ComboBox15"
                Grid = Form1.DataGridView11

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