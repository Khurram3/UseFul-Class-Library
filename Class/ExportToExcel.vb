Imports System.Data
Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel

Public Class ExportToExcel
    Public Shared Sub Export(ByVal dataTable As DataTable)
        Try
            Dim excelApp As New Excel.Application()
            Dim workbook As Excel.Workbook = excelApp.Workbooks.Add()
            Dim worksheet As Excel.Worksheet = CType(workbook.Sheets(1), Excel.Worksheet)

            ' Column Headers
            For i As Integer = 0 To dataTable.Columns.Count - 1
                worksheet.Cells(1, i + 1) = dataTable.Columns(i).ColumnName
            Next

            ' Data Rows
            For i As Integer = 0 To dataTable.Rows.Count - 1
                For j As Integer = 0 To dataTable.Columns.Count - 1
                    worksheet.Cells(i + 2, j + 1) = dataTable.Rows(i)(j).ToString()
                Next
            Next

            ' Save the file
            Dim saveFileDialog As New SaveFileDialog With {
                .Filter = "Excel Files|*.xlsx",
                .Title = "Save Excel File"
            }

            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                workbook.SaveAs(saveFileDialog.FileName)
                MessageBox.Show("Export Successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            workbook.Close()
            excelApp.Quit()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
