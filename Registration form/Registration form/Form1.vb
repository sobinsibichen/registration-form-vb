
Imports System.IO
Imports System.Windows

Public Class Form1
    Dim table As New DataTable("table")
    Dim index As Integer
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        table.Columns.Add("ID", Type.GetType("System.Int32"))

        table.Columns.Add("First Name", Type.GetType("System.String"))

        table.Columns.Add("Last Name", Type.GetType("System.String"))

        table.Columns.Add("Email", Type.GetType("System.String"))

        table.Columns.Add("Gender", Type.GetType("System.String"))

        DataGridView1.DataSource = table
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        table.Rows.Add(TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text)
        DataGridView1.DataSource = table
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        index = e.RowIndex
        Dim selectedrow As DataGridViewRow

        selectedrow = DataGridView1.Rows(index)

        TextBox1.Text = selectedrow.Cells(0).Value.ToString

        TextBox2.Text = selectedrow.Cells(1).Value.ToString

        TextBox3.Text = selectedrow.Cells(2).Value.ToString

        TextBox4.Text = selectedrow.Cells(3).Value.ToString

        TextBox5.Text = selectedrow.Cells(4).Value.ToString
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim newdata As DataGridViewRow
        newdata = DataGridView1.Rows(index)

        newdata.Cells(0).Value = TextBox1.Text


        newdata.Cells(1).Value = TextBox2.Text


        newdata.Cells(2).Value = TextBox3.Text


        newdata.Cells(3).Value = TextBox4.Text

        newdata.Cells(4).Value = TextBox5.Text

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        DataGridView1.Rows.RemoveAt(index)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim writer As New StreamWriter("D:\sample projects\Registration form\Exported data.txt")

        For i1 As Integer = 0 To DataGridView1.Rows.Count - 2 Step +1

            For j1 As Integer = 0 To DataGridView1.Columns.Count - 1 Step +1

                ' if last column
                If j1 = DataGridView1.Columns.Count - 1 Then
                    writer.Write(vbTab & DataGridView1.Rows(i1).Cells(j1).Value.ToString())
                Else
                    writer.Write(vbTab & DataGridView1.Rows(i1).Cells(j1).Value.ToString() & vbTab & "|")
                End If


            Next j1

            writer.WriteLine("")

        Next i1

        writer.Close()



        Dim xlApp As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer
        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        For i = 0 To DataGridView1.RowCount - 2
            For j = 0 To DataGridView1.ColumnCount - 1
                For k As Integer = 1 To DataGridView1.Columns.Count
                    xlWorkSheet.Cells(1, k) = DataGridView1.Columns(k - 1).HeaderText
                    xlWorkSheet.Cells(i + 2, j + 1) = DataGridView1(j, i).Value.ToString()
                Next
            Next
        Next
        xlWorkSheet.SaveAs("D:\sample projects\Registration form\exported.xlsx")
        xlWorkBook.Close()
        xlApp.Quit()
        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        MessageBox.Show("Data Exported")


    End Sub


    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Class
