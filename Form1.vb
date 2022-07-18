Imports System.IO
Imports System.Data.OleDb
Imports Microsoft.VisualBasic
'Imports Microsoft.Office.Interop.Excel
'Imports System.Data.DataTable
'Imports Microsoft.Office.Interop

Public Class Form1

    Dim fullPath As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AssignValidation(NenDoTxtBx, ValidationType.Only_Numbers)

        NenDoTxtBx.MaxLength = 4
    End Sub

    Private Function GetFileName(ByVal path As String) As String
        Dim filename As String = System.IO.Path.GetFileName(path)
        Return filename
    End Function
    Public Sub TextBoxValidation()
        Dim flag As Boolean = True

        For Each cntrl As Control In Panel1.Controls
            If TypeOf cntrl Is TextBox Then

                If CType(cntrl, TextBox).Text.Equals(String.Empty) Or (CType(cntrl, TextBox).Text = "") Then
                    If flag Then
                        'MessageBox.Show(cntrl.Name.ToString() & "は空です。確認してください。")
                        MessageBox.Show("空のボックスにすべて記入してください。")
                        flag = False
                    End If

                    cntrl.BackColor = Color.OrangeRed
                End If
            End If
        Next
    End Sub

    Public Sub FilledTxtBox(ByVal cntrlTxtxBx As Control)
        If cntrlTxtxBx.Text <> "" Then
            cntrlTxtxBx.BackColor = Color.LightGreen
        Else
            cntrlTxtxBx.BackColor = Color.White
        End If
    End Sub

    Private Sub clear1Btn_Click(sender As Object, e As EventArgs) Handles clear1Btn.Click
        Dim dialog As DialogResult
        dialog = MsgBox("全て削除しますか。", MsgBoxStyle.YesNo)

        If dialog = DialogResult.Yes Then
            TempoMeiTxtBx.Text = ""
            NenDoTxtBx.Text = ""
            IryouhinTxtBx.Text = ""
            JukyoyokaTxtBx.Text = ""
            ShokuhinTxtBx.Text = ""
            HibuppanTxtBx.Text = ""
            ResultShareTxtBx.Text = ""

            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
        End If
    End Sub

    Private Sub clear2Btn_Click(sender As Object, e As EventArgs) Handles clear2Btn.Click
        Dim dialog As DialogResult
        dialog = MsgBox("全て削除しますか。", MsgBoxStyle.YesNo)

        If dialog = DialogResult.Yes Then
            KurumaJikanTxtBx.Text = ""
            DoushinenTxtBx.Text = ""
            JisseiMapTxtBx.Text = ""
            Chirashi1TxtBx.Text = ""
            Chirashi2TxtBx.Text = ""
            Chirashi3TxtBx.Text = ""
            SaveFolderTxtBx.Text = ""
        End If
    End Sub

    Public Function getFilePath(ByVal filePathName As String(), ByVal fileType As String) As (Flag As Boolean, Path As String)
        Dim Extension As String = ""
        Dim Path As String = ""
        Dim Flag As Boolean = False

        For Each filepath As String In filePathName
            Path = filepath
        Next

        If Path <> "" Then
            Extension = Path.Substring(Path.LastIndexOf(".") + 1).ToLower()
            If Extension = fileType Then
                Flag = True
            Else
                Path = ""
            End If
        End If

        Return (Flag, Path)
    End Function
    Public Sub csvReader(ByRef Path As String)
        Dim TxtNewLine As String
        Dim IsFlagFound As Boolean = True
        Dim NewColName As String
        Dim SplitLine() As String

        Using reader As New StreamReader(Path)

            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()

            Do Until reader.EndOfStream
                TxtNewLine = Trim(reader.ReadLine())
                SplitLine = Split(TxtNewLine, ";")
                If IsFlagFound Then
                    For i = 0 To SplitLine.Length - 1
                        NewColName = Trim(SplitLine(i))
                        NewColName = NewColName.Replace(vbTab, Nothing)
                        DataGridView1.Columns.Add(NewColName, NewColName)
                    Next
                    IsFlagFound = False
                Else
                    DataGridView1.Rows.Add(SplitLine)
                End If
            Loop

        End Using
    End Sub

    Private Sub TempoMeiTxtBx_TextChanged(sender As Object, e As EventArgs) Handles TempoMeiTxtBx.TextChanged
        FilledTxtBox(TempoMeiTxtBx)
    End Sub

    Private Sub NenDoTxtBx_TextChanged(sender As Object, e As EventArgs) Handles NenDoTxtBx.TextChanged
        FilledTxtBox(NenDoTxtBx)
    End Sub

    Private Sub IryouhinTxtBx_TextChanged(sender As Object, e As EventArgs) Handles IryouhinTxtBx.TextChanged
        FilledTxtBox(IryouhinTxtBx)
    End Sub

    Private Sub IryouhinTxtBx_DragEnter(sender As Object, e As DragEventArgs) Handles IryouhinTxtBx.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub IryouhinTxtBx_DragDrop(sender As Object, e As DragEventArgs) Handles IryouhinTxtBx.DragDrop
        Dim filePath As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim fileType As String = "csv"

        Dim filePathData = getFilePath(filePath, fileType)

        If String.IsNullOrEmpty(TempoMeiTxtBx.Text) = False Then
            If filePathData.Flag Then
                '衣料品入力
                csvReader(filePathData.Path)
                MessageBox.Show("CSVファイル入力しました。")
            Else
                MessageBox.Show(".CSVファイルのみを入力してください。")
            End If

            fullPath = filePathData.Path
        Else
            fullPath = ""
            MessageBox.Show("店舗名を入力して下さい。")
        End If

        IryouhinTxtBx.Text = fullPath

    End Sub

    Private Sub JukyoyokaTxtBx_TextChanged(sender As Object, e As EventArgs) Handles JukyoyokaTxtBx.TextChanged
        FilledTxtBox(JukyoyokaTxtBx)
    End Sub

    Private Sub JukyoyokaTxtBx_DragEnter(sender As Object, e As DragEventArgs) Handles JukyoyokaTxtBx.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub JukyoyokaTxtBx_DragDrop(sender As Object, e As DragEventArgs) Handles JukyoyokaTxtBx.DragDrop
        Dim filePath As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim fileType As String = "csv"

        Dim filePathData = getFilePath(filePath, fileType)

        If String.IsNullOrEmpty(TempoMeiTxtBx.Text) = False Then
            If filePathData.Flag Then
                'CSV入力入力
                csvReader(filePathData.Path)
                MessageBox.Show("CSVファイル入力しました。")
            Else
                MessageBox.Show(".csvファイルのみを入力してください。")
            End If

            fullPath = filePathData.Path
        Else
            fullPath = ""
            MessageBox.Show("店舗名を入力して下さい。")
        End If

        JukyoyokaTxtBx.Text = fullPath
    End Sub

    Private Sub ShokuhinTxtBx_TextChanged(sender As Object, e As EventArgs) Handles ShokuhinTxtBx.TextChanged
        FilledTxtBox(ShokuhinTxtBx)
    End Sub

    Private Sub ShokuhinTxtBx_DragEnter(sender As Object, e As DragEventArgs) Handles ShokuhinTxtBx.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub ShokuhinTxtBx_DragDrop(sender As Object, e As DragEventArgs) Handles ShokuhinTxtBx.DragDrop
        Dim filePath As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim fileType As String = "csv"

        Dim filePathData = getFilePath(filePath, fileType)

        If String.IsNullOrEmpty(TempoMeiTxtBx.Text) = False Then
            If filePathData.Flag Then
                'CSV入力
                csvReader(filePathData.Path)
                MessageBox.Show("CSVファイル入力しました。")
            Else
                MessageBox.Show(".csvファイルのみを入力してください。")
            End If

            fullPath = filePathData.Path
        Else
            fullPath = ""
            MessageBox.Show("店舗名を入力して下さい。")
        End If

        ShokuhinTxtBx.Text = fullPath
    End Sub

    Private Sub HibuppanTxtBx_TextChanged(sender As Object, e As EventArgs) Handles HibuppanTxtBx.TextChanged
        FilledTxtBox(HibuppanTxtBx)
    End Sub

    Private Sub HibuppanTxtBx_DragEnter(sender As Object, e As DragEventArgs) Handles HibuppanTxtBx.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub HibuppanTxtBx_DragDrop(sender As Object, e As DragEventArgs) Handles HibuppanTxtBx.DragDrop
        Dim filePath As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim fileType As String = "csv"

        Dim filePathData = getFilePath(filePath, fileType)

        If String.IsNullOrEmpty(TempoMeiTxtBx.Text) = False Then
            If filePathData.Flag Then
                'CSV入力
                csvReader(filePathData.Path)
                MessageBox.Show("CSVファイル入力しました。")
            Else
                MessageBox.Show(".csvファイルのみを入力してください。")
            End If

            fullPath = filePathData.Path
        Else
            fullPath = ""
            MessageBox.Show("店舗名を入力して下さい。")
        End If

        HibuppanTxtBx.Text = fullPath
    End Sub

    Private Sub ResultShareTxtBx_TextChanged(sender As Object, e As EventArgs) Handles ResultShareTxtBx.TextChanged
        FilledTxtBox(ResultShareTxtBx)
    End Sub

    Private Sub ResultShareTxtBx_DragEnter(sender As Object, e As DragEventArgs) Handles ResultShareTxtBx.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub ResultShareTxtBx_DragDrop(sender As Object, e As DragEventArgs) Handles ResultShareTxtBx.DragDrop
        Dim filePath As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim fileType As String = "mdb"

        Dim filePathData = getFilePath(filePath, fileType)

        If String.IsNullOrEmpty(TempoMeiTxtBx.Text) = False Then
            If filePathData.Flag Then
                '.mdbファイル入力
                MessageBox.Show(".mdbファイル入力しました。")
            Else
                MessageBox.Show(".mdbファイルのみを入力してください。")
            End If

            fullPath = filePathData.Path
        Else
            fullPath = ""
            MessageBox.Show("店舗名を入力して下さい。")
        End If

        ResultShareTxtBx.Text = fullPath
    End Sub

    Private Sub KurumaJikanTxtBx_TextChanged(sender As Object, e As EventArgs) Handles KurumaJikanTxtBx.TextChanged
        FilledTxtBox(KurumaJikanTxtBx)
    End Sub

    Private Sub KurumaJikanTxtBx_DragEnter(sender As Object, e As DragEventArgs) Handles KurumaJikanTxtBx.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub KurumaJikanTxtBx_DragDrop(sender As Object, e As DragEventArgs) Handles KurumaJikanTxtBx.DragDrop
        Dim filePath As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim fileType As String = "tab"

        Dim filePathData = getFilePath(filePath, fileType)

        If String.IsNullOrEmpty(TempoMeiTxtBx.Text) = False Then
            If filePathData.Flag Then
                '車時間入力
                MessageBox.Show("車時間入力しました。")
            Else
                MessageBox.Show(".TABファイルのみを入力してください。")
            End If

            fullPath = filePathData.Path
        Else
            fullPath = ""
            MessageBox.Show("店舗名を入力して下さい。")
        End If

        KurumaJikanTxtBx.Text = fullPath
    End Sub

    Private Sub DoushinenTxtBx_TextChanged(sender As Object, e As EventArgs) Handles DoushinenTxtBx.TextChanged
        FilledTxtBox(DoushinenTxtBx)
    End Sub

    Private Sub DoushinenTxtBx_DragEnter(sender As Object, e As DragEventArgs) Handles DoushinenTxtBx.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub DoushinenTxtBx_DragDrop(sender As Object, e As DragEventArgs) Handles DoushinenTxtBx.DragDrop
        Dim filePath As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim fileType As String = "tab"

        Dim filePathData = getFilePath(filePath, fileType)

        If String.IsNullOrEmpty(TempoMeiTxtBx.Text) = False Then
            If filePathData.Flag Then
                '同心円入力
                MessageBox.Show("同心円入力しました。")
            Else
                MessageBox.Show(".TABファイルのみを入力してください。")
            End If

            fullPath = filePathData.Path
        Else
            fullPath = ""
            MessageBox.Show("店舗名を入力して下さい。")
        End If

        DoushinenTxtBx.Text = fullPath

    End Sub

    Private Sub JisseiMapTxtBx_TextChanged(sender As Object, e As EventArgs) Handles JisseiMapTxtBx.TextChanged
        FilledTxtBox(JisseiMapTxtBx)
    End Sub

    Private Sub JisseiMapTxtBx_DragEnter(sender As Object, e As DragEventArgs) Handles JisseiMapTxtBx.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub JisseiMapTxtBx_DragDrop(sender As Object, e As DragEventArgs) Handles JisseiMapTxtBx.DragDrop
        Dim filePath As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim fileType As String = "tab"

        Dim filePathData = getFilePath(filePath, fileType)

        If String.IsNullOrEmpty(TempoMeiTxtBx.Text) = False Then
            If filePathData.Flag Then
                '実勢商圏入力
                MessageBox.Show("実勢商圏入力しました。")
            Else
                MessageBox.Show(".TABファイルのみを入力してください。")
            End If

            fullPath = filePathData.Path
        Else
            fullPath = ""
            MessageBox.Show("店舗名を入力して下さい。")
        End If

        JisseiMapTxtBx.Text = fullPath
    End Sub

    Private Sub Chirashi1TxtBx_TextChanged(sender As Object, e As EventArgs) Handles Chirashi1TxtBx.TextChanged
        FilledTxtBox(Chirashi1TxtBx)
    End Sub

    Private Sub Chirashi1TxtBx_DragEnter(sender As Object, e As DragEventArgs) Handles Chirashi1TxtBx.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Chirashi1TxtBx_DragDrop(sender As Object, e As DragEventArgs) Handles Chirashi1TxtBx.DragDrop
        Dim filePath As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim fileType As String = "tab"


        Dim filePathData = getFilePath(filePath, fileType)

        If String.IsNullOrEmpty(TempoMeiTxtBx.Text) = False Then
            If filePathData.Flag Then
                'チラシ入力
                MessageBox.Show("チラシ1入力しました。")
            Else
                MessageBox.Show(".TABファイルのみを入力してください。")
            End If

            fullPath = filePathData.Path
        Else
            fullPath = ""
            MessageBox.Show("店舗名を入力して下さい。")
        End If

        Chirashi1TxtBx.Text = fullPath
    End Sub

    Private Sub Chirashi2TxtBx_TextChanged(sender As Object, e As EventArgs) Handles Chirashi2TxtBx.TextChanged
        FilledTxtBox(Chirashi2TxtBx)
    End Sub

    Private Sub Chirashi2TxtBx_DragEnter(sender As Object, e As DragEventArgs) Handles Chirashi2TxtBx.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Chirashi2TxtBx_DragDrop(sender As Object, e As DragEventArgs) Handles Chirashi2TxtBx.DragDrop
        Dim filePath As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim fileType As String = "tab"

        Dim filePathData = getFilePath(filePath, fileType)

        If String.IsNullOrEmpty(TempoMeiTxtBx.Text) = False Then
            If filePathData.Flag Then
                'チラシ入力
                MessageBox.Show("チラシ2入力しました。")
            Else
                MessageBox.Show(".TABファイルのみを入力してください。")
            End If

            fullPath = filePathData.Path
        Else
            fullPath = ""
            MessageBox.Show("店舗名を入力して下さい。")
        End If

        Chirashi2TxtBx.Text = fullPath
    End Sub

    Private Sub Chirashi3TxtBx_TextChanged(sender As Object, e As EventArgs) Handles Chirashi3TxtBx.TextChanged
        FilledTxtBox(Chirashi3TxtBx)
    End Sub

    Private Sub Chirashi3TxtBx_DragEnter(sender As Object, e As DragEventArgs) Handles Chirashi3TxtBx.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Chirashi3TxtBx_DragDrop(sender As Object, e As DragEventArgs) Handles Chirashi3TxtBx.DragDrop
        Dim filePath As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim fileType As String = "tab"

        Dim filePathData = getFilePath(filePath, fileType)

        If String.IsNullOrEmpty(TempoMeiTxtBx.Text) = False Then
            If filePathData.Flag Then
                'チラシ入力
                MessageBox.Show("チラシ3入力しました。")
            Else
                MessageBox.Show(".TABファイルのみを入力してください。")
            End If

            fullPath = filePathData.Path
        Else
            fullPath = ""
            MessageBox.Show("店舗名を入力して下さい。")
        End If

        Chirashi3TxtBx.Text = fullPath
    End Sub

    'Public Shared Sub ConvertAnWorksheetToCsv(ByVal Path As String)
    '    ''Create an instance of Workbook class
    '    'Dim workbook As Workbook = New Workbook()
    '    ''Load an Excel file
    '    'workbook.LoadFromFile(Path)
    '    '
    '    ''Get the first worksheet
    '    'Dim sheet As Worksheet = workbook.Worksheets(0)
    '    '
    '    ''Save the worksheet as CSV
    '    'sheet.SaveToFile("ExcelToCSV.csv", ",", Encoding.UTF8)
    '
    '    Dim fullPath As String
    '    Dim fileResult As String
    '    Dim numRow As Integer
    '
    '    fileResult = SaveFolderTxtBx.Text
    '
    '    Dim obook As Excel.Workbook
    '    Dim oapp As Excel.Application
    '    oapp = New Excel.Application
    '    obook = oapp.Workbooks.Open(fileResult)
    '    numRow = 3
    '
    '    While (obook.ActiveSheet.Cells(numRow, 1).Value IsNot Nothing)
    '        numRow = numRow + 1
    '    End While
    '
    '    MsgBox(numRow)
    'End Sub

    Private Sub SaveFolderTxtBx_TextChanged(sender As Object, e As EventArgs) Handles SaveFolderTxtBx.TextChanged
        FilledTxtBox(SaveFolderTxtBx)
    End Sub

    Private Sub SaveFolderTxtBx_DragEnter(sender As Object, e As DragEventArgs) Handles SaveFolderTxtBx.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub SaveFolderTxtBx_DragDrop(sender As Object, e As DragEventArgs) Handles SaveFolderTxtBx.DragDrop
        Dim filePathName As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim Path As String = ""

        For Each filepath As String In filePathName
            Path = filepath
        Next

        If String.IsNullOrEmpty(TempoMeiTxtBx.Text) = False Then
            MessageBox.Show("保存パス入力しました。")
        Else
            Path = ""
            MessageBox.Show("店舗名を入力して下さい。")
        End If

        SaveFolderTxtBx.Text = Path
    End Sub

    Private Sub ShareCalcBtn_Click(sender As Object, e As EventArgs) Handles ShareCalcBtn.Click
        TextBoxValidation()
    End Sub

    Private Sub ShareWPCBtn_Click(sender As Object, e As EventArgs) Handles ShareWPCBtn.Click
        TextBoxValidation()
    End Sub

    Private Sub CreatePolygonBtn_Click(sender As Object, e As EventArgs) Handles CreatePolygonBtn.Click
        TextBoxValidation()
    End Sub

    Private Sub CreateJisseiBtn_Click(sender As Object, e As EventArgs) Handles CreateJisseiBtn.Click
        TextBoxValidation()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim text As String = TempoMeiTxtBx.Text
        Dim hankaku, zenkaku As String

        hankaku = Strings.StrConv(text, 4)
        Console.WriteLine(hankaku) '123ｱｲｳあいう

        zenkaku = Strings.StrConv(text, VbStrConv.Wide)
        Console.WriteLine(zenkaku) '１２３アイウあいう
    End Sub
End Class
