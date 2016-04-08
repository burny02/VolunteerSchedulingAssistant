Public Class RotaForm
    Private Sub RotaForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Label1.Text = Format(DateTime.Today, "dddd dd-MMM-yyyy")
        Label2.Text = Format(DateAdd(DateInterval.Day, 1, DateTime.Today), "dddd dd-MMM-yyyy")
        Label3.Text = Format(DateAdd(DateInterval.Day, 2, DateTime.Today), "dddd dd-MMM-yyyy")
        Label4.Text = Format(DateAdd(DateInterval.Day, 3, DateTime.Today), "dddd dd-MMM-yyyy")
        Label5.Text = Format(DateAdd(DateInterval.Day, 4, DateTime.Today), "dddd dd-MMM-yyyy")

        RefreshData()

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        Label1.Text = Format(DateAdd(DateInterval.Day, -1, CDate(Label1.Text)), "dddd dd-MMM-yyyy")
        Label2.Text = Format(DateAdd(DateInterval.Day, -1, CDate(Label2.Text)), "dddd dd-MMM-yyyy")
        Label3.Text = Format(DateAdd(DateInterval.Day, -1, CDate(Label3.Text)), "dddd dd-MMM-yyyy")
        Label4.Text = Format(DateAdd(DateInterval.Day, -1, CDate(Label4.Text)), "dddd dd-MMM-yyyy")
        Label5.Text = Format(DateAdd(DateInterval.Day, -1, CDate(Label5.Text)), "dddd dd-MMM-yyyy")

        RefreshData()

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click

        Label1.Text = Format(DateAdd(DateInterval.Day, 1, CDate(Label1.Text)), "dddd dd-MMM-yyyy")
        Label2.Text = Format(DateAdd(DateInterval.Day, 1, CDate(Label2.Text)), "dddd dd-MMM-yyyy")
        Label3.Text = Format(DateAdd(DateInterval.Day, 1, CDate(Label3.Text)), "dddd dd-MMM-yyyy")
        Label4.Text = Format(DateAdd(DateInterval.Day, 1, CDate(Label4.Text)), "dddd dd-MMM-yyyy")
        Label5.Text = Format(DateAdd(DateInterval.Day, 1, CDate(Label5.Text)), "dddd dd-MMM-yyyy")

        RefreshData()

    End Sub

    Private Sub RefreshData()

        Dim SQLTemplate As String = "SELECT RotaID, FName & ' ' & SName AS FullName, " &
        "format(RotaStart,'HH:mm') & ' > ' & format(RotaEnd,'HH:mm') AS Timeframe FROM Rota a INNER JOIN Staff b on a.StaffID=b.StaffID "

        Dim SQLOrder As String = " ORDER BY RotaStart ASC, RotaEnd ASC"

        Dim SQL1 As String = SQLTemplate & "WHERE format(RotaStart, 'dddd dd-MMM-yyyy')='" & Label1.Text & "' OR format(RotaEnd, 'dddd dd-MMM-yyyy')='" & Label1.Text & "'" & SQLOrder
        Dim SQL2 As String = SQLTemplate & "WHERE format(RotaStart, 'dddd dd-MMM-yyyy')='" & Label2.Text & "' OR format(RotaEnd, 'dddd dd-MMM-yyyy')='" & Label2.Text & "'" & SQLOrder
        Dim SQL3 As String = SQLTemplate & "WHERE format(RotaStart, 'dddd dd-MMM-yyyy')='" & Label3.Text & "' OR format(RotaEnd, 'dddd dd-MMM-yyyy')='" & Label3.Text & "'" & SQLOrder
        Dim SQL4 As String = SQLTemplate & "WHERE format(RotaStart, 'dddd dd-MMM-yyyy')='" & Label4.Text & "' OR format(RotaEnd, 'dddd dd-MMM-yyyy')='" & Label4.Text & "'" & SQLOrder
        Dim SQL5 As String = SQLTemplate & "WHERE format(RotaStart, 'dddd dd-MMM-yyyy')='" & Label5.Text & "' OR format(RotaEnd, 'dddd dd-MMM-yyyy')='" & Label5.Text & "'" & SQLOrder

        Dim Dt() As DataTable = OverClass.MultiTempDataTable(SQL1, SQL2, SQL3, SQL4, SQL5)

        DataGridView1.DataSource = Dt(0)
        DataGridView2.DataSource = Dt(1)
        DataGridView3.DataSource = Dt(2)
        DataGridView4.DataSource = Dt(3)
        DataGridView5.DataSource = Dt(4)

        DataGridView1.Columns("RotaID").Visible = False
        DataGridView2.Columns("RotaID").Visible = False
        DataGridView3.Columns("RotaID").Visible = False
        DataGridView4.Columns("RotaID").Visible = False
        DataGridView5.Columns("RotaID").Visible = False


        On Error Resume Next
        RemoveHandler DataGridView1.CellContentDoubleClick, AddressOf CancelRota
        RemoveHandler DataGridView2.CellContentDoubleClick, AddressOf CancelRota
        RemoveHandler DataGridView3.CellContentDoubleClick, AddressOf CancelRota
        RemoveHandler DataGridView4.CellContentDoubleClick, AddressOf CancelRota
        RemoveHandler DataGridView5.CellContentDoubleClick, AddressOf CancelRota
        On Error GoTo 0
        AddHandler DataGridView1.CellContentDoubleClick, AddressOf CancelRota
        AddHandler DataGridView2.CellContentDoubleClick, AddressOf CancelRota
        AddHandler DataGridView3.CellContentDoubleClick, AddressOf CancelRota
        AddHandler DataGridView4.CellContentDoubleClick, AddressOf CancelRota
        AddHandler DataGridView5.CellContentDoubleClick, AddressOf CancelRota

    End Sub

    Private Sub CancelRota(sender As Object, e As EventArgs)

        Dim Cell As DataGridViewCell = sender.CurrentCell
        Dim RotaID As Long = Cell.DataGridView("RotaID", Cell.RowIndex).Value

        If MsgBox("Are you sure you want to remove " & Cell.DataGridView("FullName", Cell.RowIndex).Value & " " &
                  Cell.DataGridView("TimeFrame", Cell.RowIndex).Value, MsgBoxStyle.YesNo) = vbYes Then

            OverClass.ExecuteSQL("DELETE * FROM Rota WHERE RotaID=" & RotaID)
            RefreshData()

        End If

    End Sub

End Class