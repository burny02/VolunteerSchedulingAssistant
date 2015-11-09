Imports System.Drawing.Printing
Imports iTextSharp.text.pdf
Imports iTextSharp.text
Imports System.IO

Public Class SOAForm

    Dim YRange As Long = 15
    Dim counter As Long = 0

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Dim PrintDocument1 As New PrintDocument()

        Dim pdiag As New PrintDialog

        AddHandler PrintDocument1.PrintPage, AddressOf PrintDocument1_PrintPage

        PrintDocument1.DefaultPageSettings.Landscape = True


        pdiag.AllowSelection = True
        pdiag.Document = PrintDocument1

        If pdiag.ShowDialog() = DialogResult.OK Then PrintDocument1.Print()

        ExportPDF()


    End Sub

    Protected Sub PrintDocument1_PrintPage(ByVal sender As Object, ByVal e As PrintPageEventArgs)

        DataGridView1.CurrentCell = Nothing
        DataGridView1.ScrollBars = ScrollBars.None
        e.HasMorePages = False
        Dim RowsLeft As Long = 0
        Dim RowHeight As Long = DataGridView1.Rows(1).Height + 3



        DataGridView1.ColumnHeadersVisible = True
        Dim bmHead As New Bitmap(DataGridView1.Width, DataGridView1.Rows(1).Height + 3)
        DataGridView1.DrawToBitmap(bmHead, New System.Drawing.Rectangle(0, 0, Me.DataGridView1.Width, RowHeight))
        e.Graphics.DrawImage(bmHead, 0, YRange, e.PageBounds.Width - 15, RowHeight)
        DataGridView1.ColumnHeadersVisible = False
        YRange += RowHeight



        For Each row As DataGridViewRow In DataGridView1.Rows

            If row.Visible = True Then


                Dim bm As New Bitmap(DataGridView1.Width, DataGridView1.Rows(1).Height + 3)

                DataGridView1.DrawToBitmap(bm, New System.Drawing.Rectangle(0, 0, Me.DataGridView1.Width, RowHeight))
                e.Graphics.DrawImage(bm, 0, YRange, e.PageBounds.Width - 15, RowHeight)
                YRange += RowHeight

                counter += 1
                row.Visible = False
                If (counter Mod 16) = 0 Then


                    For Each row2 As DataGridViewRow In DataGridView1.Rows
                        If row2.Visible = True Then RowsLeft += 1
                    Next

                    If RowsLeft <> 0 Then

                        e.HasMorePages = True
                        YRange = 15
                        Exit Sub

                    End If

                End If

            End If

        Next

        If e.HasMorePages = False Then
            DataGridView1.ColumnHeadersVisible = True
            DataGridView1.ScrollBars = ScrollBars.Both
            For Each row As DataGridViewRow In DataGridView1.Rows
                row.Visible = True
            Next
            YRange = 15
            counter = 0
        End If

    End Sub

    Private Sub ExportPDF()

        'Creating iTextSharp Table from the DataTable data
        Dim ColCount As Long = 0

        For Each column As DataGridViewColumn In DataGridView1.Columns

            If column.Visible = True Then

                ColCount = ColCount + 1

            End If

        Next

        Dim pdfTable As New PdfPTable(ColCount)

        pdfTable.DefaultCell.Padding = 3

        pdfTable.WidthPercentage = 100

        pdfTable.HorizontalAlignment = Element.ALIGN_CENTER

        pdfTable.DefaultCell.BorderWidth = 1



        'Adding Header row

        For Each column As DataGridViewColumn In DataGridView1.Columns

            If column.Visible = True Then

                Dim cell As New PdfPCell(New Phrase(column.HeaderText))


                pdfTable.AddCell(cell)

            End If

        Next



        'Adding DataRow

        For Each row As DataGridViewRow In DataGridView1.Rows

            For Each cell As DataGridViewCell In row.Cells

                If cell.Visible = True Then pdfTable.AddCell(cell.Value.ToString())

            Next

        Next



        'Exporting to PDF

        Dim NextID As Long
        Try
            NextID = OverClass.TempDataTable("SELECT max(ArchiveID) FROM Reportarchive").Rows(0).Item(0) + 1
        Catch ex As Exception
            NextID = 1
        End Try

        Dim ArchiveType As String = "SchedulePreview"
        Dim Criteria As String = "Dates: " & Format(Form1.DateTimePicker1.Value, "dd-MMM-yyyy HH:mm") _
                                             & " -> " & Format(Form1.DateTimePicker2.Value, "dd-MMM-yyyy HH:mm") _
                                             & vbNewLine & "Study: " & Form1.ComboBox17.Text

        Dim pdfPath As String = ReportPath & NextID & ".pdf"

        OverClass.ExecuteSQL("INSERT INTO ReportArchive (ArchiveID, ArchivePath, ArchiveUser, ArchiveType, ArchiveCriteria) " & _
        "VALUES (" & NextID & ", '" & pdfPath & "', '" & OverClass.GetUserName & "', '" & ArchiveType & "', '" _
        & Criteria & "')")

        Using stream As New FileStream(pdfPath, FileMode.Create)

            Dim pdfDoc As New Document(PageSize.A4.Rotate, 10.0F, 10.0F, 10.0F, 0.0F)


            PdfWriter.GetInstance(pdfDoc, stream)

            pdfDoc.Open()

            pdfDoc.Add(pdfTable)

            pdfDoc.Close()

            stream.Close()

        End Using

    End Sub

End Class