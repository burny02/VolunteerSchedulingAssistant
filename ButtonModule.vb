Module ButtonModule

    Public Sub ButtonSpecifics(sender As Object, e As EventArgs)

        Dim ctl As Object = Nothing

        Select Case sender.name.ToString

            Case "Button1"
                Call Saver(Form1.DataGridView1)
            Case "Button3"
                Call Saver(Form1.DataGridView2)
            Case "Button4"
                Call Saver(Form1.DataGridView3)
            Case "Button6"
                Call Saver(Form1.DataGridView5)
            Case "Button7"
                Call Saver(Form1.DataGridView6)
            Case "Button8"
                Call Saver(Form1.DataGridView7)
            Case "Button9"
                Call Saver(Form1.DataGridView8)
            Case "Button10"
                Call Generator()
            Case "Button11"
                Call Saver(Form1.DataGridView9)
            Case "Button12"
                Call Saver(Form1.DataGridView10)
            Case "Button13"
                Call Saver(Form1.DataGridView11)
            Case "Button5"
                Call Saver(Form1.DataGridView4)
            Case "Button14"
                Call Saver(Form1.DataGridView4)

        End Select



    End Sub


End Module
