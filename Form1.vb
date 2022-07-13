Imports System.IO
Imports System.Data.OleDb

Public Class Form1

    Private Sub CSVFilePathTxtBx_TextChanged(sender As Object, e As EventArgs) Handles CSVFilePathTxtBx.TextChanged

    End Sub

    Private Sub CSVFilePathTxtBx_DragDrop(sender As Object, e As DragEventArgs) Handles CSVFilePathTxtBx.DragDrop
        Dim CSVfilePath As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim Path As String = ""


        For Each filepath As String In CSVfilePath
            Path = filepath
        Next

        If Path <> "" Then

            Dim Extension As String = Path.Substring(Path.LastIndexOf(".") + 1).ToLower()
            Dim firstRow As String()
            Dim headerTable As New DataTable()
            Dim table As New DataTable()
            Dim columnsData As New DataTable()
            Dim lines() As String
            Dim vals() As String


            If Extension = "csv" Then
                'Have no problem
                CSVFilePathTxtBx.Text = Path

                'Read CSV
                'Using MyReader As New FileIO.TextFieldParser(Path)
                '    MyReader.TextFieldType = FileIO.FieldType.Delimited
                '    MyReader.SetDelimiters(",")
                '
                '    firstRow = MyReader.ReadFields()
                '    For Each currentField In firstRow
                '        'Console.WriteLine(CurrentField)
                '        headerTable.Columns.Add(currentField, Type.GetType("System.String"))
                '    Next
                '    DataGridView1.DataSource = headerTable
                '
                '    'lines = File.ReadAllLines(Path)
                '    lines = MyReader.ReadFields()
                '    For i As Integer = 0 To lines.Length - 1 Step +1
                '        vals = lines(i).ToString().Split(",")
                '        Dim row(vals.Length - 1) As String
                '        For j As Integer = 0 To vals.Length - 1 Step +1
                '            row(j) = vals(j).Trim()
                '        Next j
                '        headerTable.Rows.Add(row)
                '    Next i
                '    DataGridView1.DataSource = headerTable
                '    'While Not MyReader.EndOfData
                '    '    Try
                '    '        Dim currentRow As String() = MyReader.ReadFields()
                '    '        For Each currentField In currentRow
                '    '            headerTable.Columns.Add(currentField)
                '    '        Next
                '    '        DataGridView1.DataSource = headerTable
                '    '    Catch ex As Exception
                '    '        MsgBox("Line " & ex.Message & " is not valid")
                '    '    End Try
                '    'End While
                '
                'End Using

                MessageBox.Show("入力しました。")
            Else
                CSVFilePathTxtBx.Text = ""
                MessageBox.Show(".csvファイルのみを入力してください。")
            End If
        End If

    End Sub

    Private Sub CSVFilePathTxtBx_DragEnter(sender As Object, e As DragEventArgs) Handles CSVFilePathTxtBx.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Function GetFileName(ByVal path As String) As String
        Dim filename As String = System.IO.Path.GetFileName(path)
        Return filename
    End Function

End Class
