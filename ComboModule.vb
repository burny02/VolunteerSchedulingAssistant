Module ComboModule

    Public Sub GenericCombo(sender As Object, e As EventArgs)

        If sender.SelectedValue.ToString <> "System.Data.DataRowView" Then
            If sender.SelectedValue.ToString <> vbNullString Then

                If OverClass.UnloadData() = True Then Exit Sub
                OverClass.ResetCollection()
                Call ComboSpecifics(sender)

            End If
        End If

    End Sub

    Private Sub ComboSpecifics(sender As Object)

        Dim ctl As Object = Nothing

        Select Case sender.name.ToString

            Case "ComboBox2"
                ctl = Form1.DataGridView5
            Case "ComboBox3"
                ctl = Form1.DataGridView6
            Case "ComboBox4"
                ctl = Form1.DataGridView6
                Form1.ComboBox3.DataSource = OverClass.TempDataTable("SELECT StudyTimepointID, " & _
                                                              "TimepointName FROM StudyTimepoint WHERE StudyID=" _
                                                              & Form1.ComboBox4.SelectedValue.ToString & _
                                                              " ORDER BY TimepointName ASC")
            Case "ComboBox5"
                ctl = Form1.DataGridView7
            Case "ComboBox6"
                ctl = Form1.DataGridView8
                Form1.ComboBox7.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                                            "CohortName FROM Cohort WHERE StudyID=" _
                                                                            & Form1.ComboBox6.SelectedValue.ToString & _
                                                                            " ORDER BY CohortName ASC")
            Case "ComboBox7"
                ctl = Form1.DataGridView8

            Case "ComboBox9"
                ctl = Form1.DataGridView9
                Form1.ComboBox10.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                                            "CohortName FROM Cohort WHERE StudyID=" _
                                                                            & Form1.ComboBox9.SelectedValue.ToString & _
                                                                            " ORDER BY CohortName ASC")
            Case "ComboBox10"
                ctl = Form1.DataGridView9

            Case "ComboBox11"
                ctl = Form1.DataGridView10
                Form1.ComboBox12.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                              "CohortName FROM Cohort WHERE StudyID=" _
                                                              & Form1.ComboBox11.SelectedValue.ToString & _
                                                                " AND Generated=True" & _
                                                                " ORDER BY CohortName ASC")

                Form1.ComboBox13.DataSource = OverClass.TempDataTable("SELECT RVLNO & ' - ' & Initials as Display, VolID " & _
                                                              "FROM Volunteer WHERE CohortID=" _
                                                              & Form1.ComboBox12.SelectedValue.ToString & _
                                                                " ORDER BY Display ASC")
            Case "ComboBox12"
                ctl = Form1.DataGridView10
                Form1.ComboBox13.DataSource = OverClass.TempDataTable("SELECT RVLNO & ' - ' & Initials as Display, VolID " & _
                                                              "FROM Volunteer WHERE CohortID=" _
                                                              & Form1.ComboBox12.SelectedValue.ToString & _
                                                                " ORDER BY Display ASC")

            Case "ComboBox13"
                ctl = Form1.DataGridView10

            Case "ComboBox14"
                ctl = Form1.DataGridView11
                Form1.ComboBox15.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                              "CohortName FROM Cohort WHERE StudyID=" _
                                                              & Form1.ComboBox14.SelectedValue.ToString & _
                                                              " AND Generated=True " & _
                                                              " ORDER BY CohortName ASC")

            Case "ComboBox15"
                ctl = Form1.DataGridView11


            Case "ComboBox1"
                ctl = Form1.DataGridView4
                Form1.ComboBox15.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
                                                              "CohortName FROM Cohort WHERE StudyID=" _
                                                              & Form1.ComboBox1.SelectedValue.ToString & _
                                                              " AND Generated=True " & _
                                                              " ORDER BY CohortName ASC")

            Case "ComboBox16"
                ctl = Form1.DataGridView4

        End Select

        If Not IsNothing(ctl) Then Call Form1.Specifics(ctl)

    End Sub


End Module
