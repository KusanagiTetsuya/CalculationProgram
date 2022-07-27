Imports System.IO
Imports System.Data.OleDb
Imports Microsoft.VisualBasic
Imports System.Data.SqlClient

Public Class Form1
    Dim Conn As SqlConnection
    Dim MyDB As String
    Dim Ds As New DataSet
    Dim Da As SqlDataAdapter
    Dim flag As Boolean = False

    Sub ConnectionDB()
        MyDB = "Data Source = 192.168.1.3; initial catalog=WPC;User ID=sa;Password=Msmskmykmsny7741;Integrated Security=False;Trusted_Connection=False;"
        Conn = New SqlConnection(MyDB)

        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
            'MsgBox("DB Connected")
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call ConnectionDB()

        Dim cmd As SqlCommand = New SqlCommand("SELECT * FROM [dbo].[T_AMTourokuFile]", Conn)
        Da = New SqlDataAdapter(cmd)

        Dim autoCmpltStr As New AutoCompleteStringCollection

        Ds.Clear()
        Da.Fill(Ds)

        'オートコンプリート文字列
        For i As Integer = 0 To Ds.Tables(0).Rows.Count - 1
            autoCmpltStr.Add(Ds.Tables(0).Rows(i)("店舗名称").ToString)
        Next
        TempoMeiTxtBx.AutoCompleteSource = AutoCompleteSource.CustomSource
        TempoMeiTxtBx.AutoCompleteCustomSource = autoCmpltStr
        TempoMeiTxtBx.AutoCompleteMode = AutoCompleteMode.Suggest

        '番号のみのテキストボックスの設定
        AssignValidation(NenDoTxtBx, ValidationType.Only_Numbers)
        AssignValidation(IryouUriTxtBx, ValidationType.Only_Numbers)
        AssignValidation(JukyoUriTxtBx, ValidationType.Only_Numbers)
        AssignValidation(ShokuUriTxtBx, ValidationType.Only_Numbers)
        AssignValidation(HibupUriTxtBx, ValidationType.Only_Numbers)

        '最大桁数の設定
        NenDoTxtBx.MaxLength = 4
        IryouUriTxtBx.MaxLength = 14
        JukyoUriTxtBx.MaxLength = 14
        ShokuUriTxtBx.MaxLength = 14
        HibupUriTxtBx.MaxLength = 14

    End Sub

    Private Function GetFileName(ByVal path As String) As String
        Dim filename As String = System.IO.Path.GetFileName(path)
        Return filename
    End Function

    Public Sub TextBoxEmptyValidation()
        Dim flag As Boolean = True

        For Each cntrl As Control In Panel1.Controls
            If TypeOf cntrl Is TextBox Then
                If CType(cntrl, TextBox).Text.Equals(String.Empty) Or (CType(cntrl, TextBox).Text = "") Then
                    If flag Then
                        MessageBox.Show("空のボックスにすべて記入してください。")
                        flag = False
                    End If
                    cntrl.BackColor = Color.OrangeRed
                End If
            End If
        Next
    End Sub

    Public Sub TextBoxColorValidation()
        For Each cntrl1 As Control In Panel1.Controls
            If TypeOf cntrl1 Is TextBox Then
                If CType(cntrl1, TextBox).Enabled.Equals(True) Then
                    If CType(cntrl1, TextBox).Text.Equals(String.Empty) Or (CType(cntrl1, TextBox).Text = "") Then
                        CType(cntrl1, TextBox).BackColor = Color.White
                    Else
                        CType(cntrl1, TextBox).BackColor = Color.LightGreen
                    End If
                Else
                    CType(cntrl1, TextBox).BackColor = Color.White
                End If
            End If
        Next

        For Each cntrl2 As Control In Panel2.Controls
            If TypeOf cntrl2 Is TextBox Then
                If CType(cntrl2, TextBox).Enabled.Equals(True) Then
                    If CType(cntrl2, TextBox).Text.Equals(String.Empty) Or (CType(cntrl2, TextBox).Text = "") Then
                        CType(cntrl2, TextBox).BackColor = Color.White
                    Else
                        CType(cntrl2, TextBox).BackColor = Color.LightGreen
                    End If
                Else
                    CType(cntrl2, TextBox).BackColor = Color.White
                End If
            End If
        Next
    End Sub

    Public Sub ControlStatusChange(ByRef CntrlStat As Boolean)
        For Each cntrl1 As Control In Panel1.Controls
            If cntrl1.Name <> "TempoMeiTxtBx" Then
                If CntrlStat Then
                    cntrl1.Enabled = True
                Else
                    cntrl1.Enabled = False
                End If
            End If
        Next

        For Each cntrl2 As Control In Panel2.Controls
            If CntrlStat Then
                cntrl2.Enabled = True
            Else
                cntrl2.Enabled = False
            End If
        Next
    End Sub

    'テキストボックスの状態に応じて色を変える機能
    Public Sub FilledTxtBox(ByVal cntrlTxtxBx As Control)
        If cntrlTxtxBx.Text <> "" Then
            cntrlTxtxBx.BackColor = Color.LightGreen
        Else
            cntrlTxtxBx.BackColor = Color.White
        End If
    End Sub

    Private Sub clear1Btn_Click(sender As Object, e As EventArgs) Handles clear1Btn.Click
        Dim dialog As DialogResult

        dialog = MsgBox("全て削除しますか。", MsgBoxStyle.YesNo, "注意")
        If dialog = DialogResult.Yes Then
            'テキストボックス空に設定
            TempoMeiTxtBx.Text = ""

            '色設定デフォルト（白色）
            TempoMeiTxtBx.BackColor = Color.White
            NenDoTxtBx.BackColor = Color.White
            IryouhinTxtBx.BackColor = Color.White
            JukyoyokaTxtBx.BackColor = Color.White
            ShokuhinTxtBx.BackColor = Color.White
            HibuppanTxtBx.BackColor = Color.White
            ResultShareTxtBx.BackColor = Color.White
            KurumaJikanTxtBx.BackColor = Color.White
            DoushinenTxtBx.BackColor = Color.White
            JisseiMapTxtBx.BackColor = Color.White
            Chirashi1TxtBx.BackColor = Color.White
            Chirashi2TxtBx.BackColor = Color.White
            Chirashi3TxtBx.BackColor = Color.White
            SaveFolderTxtBx.BackColor = Color.White

            'DataGridView空に設定
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
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

    Sub uploadToDB(ByRef dataType As String, ByVal pathName As String)
        'Conn.Close()
        Dim cmd As SqlCommand = New SqlCommand("SELECT * FROM [dbo].[T_AMTourokuFile] WHERE 
            [店舗名称] = @TempoMei AND 
            [種別] = '" & dataType & "' AND 
            [FilePath] = @pathFile", Conn)
        cmd.Parameters.AddWithValue("@TempoMei", TempoMeiTxtBx.Text)
        cmd.Parameters.AddWithValue("@pathFile", pathName)

        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If

        Dim reader As SqlDataReader = cmd.ExecuteReader()
        If reader.HasRows Then

            '[店舗名称]と[FilePath]はDBに合わせる場合
            MsgBox("Already exist!")
        Else
            Conn.Close()
            cmd = New SqlCommand("SELECT * FROM [dbo].[T_AMTourokuFile] WHERE 
            [店舗名称] = @TempoMei AND 
            [種別] = '" & dataType & "'", Conn)

            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            cmd.Parameters.AddWithValue("@TempoMei", TempoMeiTxtBx.Text)
            reader = cmd.ExecuteReader()
            If reader.HasRows Then
                Conn.Close()

                '[店舗名称]しかDBに合わさない場合
                cmd = New SqlCommand("UPDATE [dbo].[T_AMTourokuFile]
                SET [FilePath] = '" & pathName & "'
                WHERE [店舗名称] = '" & TempoMeiTxtBx.Text & "' AND [種別] = '" & dataType & "'", Conn)

                If Conn.State = ConnectionState.Closed Then
                    Conn.Open()
                End If

                cmd.ExecuteNonQuery()
            Else
                Conn.Close()

                '[店舗名称]と[FilePath]はDBに合わさない場合
                cmd = New SqlCommand("INSERT INTO [dbo].[T_AMTourokuFile]
                ([店舗名称],[種別],[FilePath])
                VALUES
                ('" & TempoMeiTxtBx.Text & "','" & dataType & "','" & pathName & "')", Conn)

                If Conn.State = ConnectionState.Closed Then
                    Conn.Open()
                End If

                cmd.ExecuteNonQuery()
            End If

        End If

        Conn.Close()
    End Sub
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
        Dim storeNmDatabase As String
        Dim typeNmDatabase As String

        'テキストボックス空に設定
        NenDoTxtBx.Text = ""
        IryouhinTxtBx.Text = ""
        IryouUriTxtBx.Text = ""
        JukyoyokaTxtBx.Text = ""
        JukyoUriTxtBx.Text = ""
        ShokuhinTxtBx.Text = ""
        ShokuUriTxtBx.Text = ""
        HibuppanTxtBx.Text = ""
        HibupUriTxtBx.Text = ""
        ResultShareTxtBx.Text = ""
        KurumaJikanTxtBx.Text = ""
        DoushinenTxtBx.Text = ""
        JisseiMapTxtBx.Text = ""
        Chirashi1TxtBx.Text = ""
        Chirashi2TxtBx.Text = ""
        Chirashi3TxtBx.Text = ""
        SaveFolderTxtBx.Text = ""

        For i As Integer = 0 To Ds.Tables(0).Rows.Count - 1
            storeNmDatabase = Ds.Tables(0).Rows(i)(0).ToString()
            If TempoMeiTxtBx.Text = storeNmDatabase Then

                typeNmDatabase = Ds.Tables(0).Rows(i)(1).ToString()

                Select Case typeNmDatabase
                    Case "Iryouhin"
                        IryouhinTxtBx.Text = Ds.Tables(0).Rows(i)(2).ToString()
                    Case "Jukyoyoka"
                        JukyoyokaTxtBx.Text = Ds.Tables(0).Rows(i)(2).ToString()
                    Case "Shokuhin"
                        ShokuhinTxtBx.Text = Ds.Tables(0).Rows(i)(2).ToString()
                    Case "Hibuppan"
                        HibuppanTxtBx.Text = Ds.Tables(0).Rows(i)(2).ToString()
                    Case "Kuruma"
                        KurumaJikanTxtBx.Text = Ds.Tables(0).Rows(i)(2).ToString()
                    Case "Dousin"
                        DoushinenTxtBx.Text = Ds.Tables(0).Rows(i)(2).ToString()
                    Case "Region1"
                        Chirashi1TxtBx.Text = Ds.Tables(0).Rows(i)(2).ToString()
                    Case "Region2"
                        Chirashi2TxtBx.Text = Ds.Tables(0).Rows(i)(2).ToString()
                    Case "Region3"
                        Chirashi3TxtBx.Text = Ds.Tables(0).Rows(i)(2).ToString()
                    Case "Jissei"
                        JisseiMapTxtBx.Text = Ds.Tables(0).Rows(i)(2).ToString()
                    Case "ResultShare"
                        ResultShareTxtBx.Text = Ds.Tables(0).Rows(i)(2).ToString()
                    Case Else
                        'Nothing
                End Select
            End If
        Next


        If String.IsNullOrEmpty(TempoMeiTxtBx.Text) = False Then
            'テキストボックス・ボタンの有効化ステータス（TRUE）
            ControlStatusChange(True)
        Else
            'テキストボックス・ボタンの有効化ステータス（FALSE）
            ControlStatusChange(False)
        End If

        FilledTxtBox(TempoMeiTxtBx)
        TextBoxColorValidation()
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

        If filePathData.Flag Then
            csvReader(filePathData.Path)
            MessageBox.Show("CSVファイル入力しました。")
        Else
            MessageBox.Show(".CSVファイルのみを入力してください。")
        End If

        IryouhinTxtBx.Text = filePathData.Path
        uploadToDB("Iryouhin", filePathData.Path)
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

        If filePathData.Flag Then
            'CSV入力入力
            csvReader(filePathData.Path)
            MessageBox.Show("CSVファイル入力しました。")
        Else
            MessageBox.Show(".csvファイルのみを入力してください。")
        End If

        JukyoyokaTxtBx.Text = filePathData.Path
        uploadToDB("Jukyoyoka", filePathData.Path)
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

        If filePathData.Flag Then
            'CSV入力
            csvReader(filePathData.Path)
            MessageBox.Show("CSVファイル入力しました。")
        Else
            MessageBox.Show(".csvファイルのみを入力してください。")
        End If

        ShokuhinTxtBx.Text = filePathData.Path
        uploadToDB("Shokuhin", filePathData.Path)
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

        If filePathData.Flag Then
            'CSV入力
            csvReader(filePathData.Path)
            MessageBox.Show("CSVファイル入力しました。")
        Else
            MessageBox.Show(".csvファイルのみを入力してください。")
        End If

        HibuppanTxtBx.Text = filePathData.Path
        uploadToDB("Hibuppan", filePathData.Path)
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

        If filePathData.Flag Then
            '.mdbファイル入力
            MessageBox.Show(".mdbファイル入力しました。")
        Else
            MessageBox.Show(".mdbファイルのみを入力してください。")
        End If

        ResultShareTxtBx.Text = filePathData.Path
        uploadToDB("ResultShare", filePathData.Path)
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

        If filePathData.Flag Then
            '車時間入力
            MessageBox.Show("車時間入力しました。")
        Else
            MessageBox.Show(".TABファイルのみを入力してください。")
        End If

        KurumaJikanTxtBx.Text = filePathData.Path
        uploadToDB("Kuruma", filePathData.Path)
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

        If filePathData.Flag Then
            '同心円入力
            MessageBox.Show("同心円入力しました。")
        Else
            MessageBox.Show(".TABファイルのみを入力してください。")
        End If

        DoushinenTxtBx.Text = filePathData.Path
        uploadToDB("Dousin", filePathData.Path)
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

        If filePathData.Flag Then
            '実勢商圏入力
            MessageBox.Show("実勢商圏入力しました。")
        Else
            MessageBox.Show(".TABファイルのみを入力してください。")
        End If

        JisseiMapTxtBx.Text = filePathData.Path
        uploadToDB("Jissei", filePathData.Path)
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

        If filePathData.Flag Then
            'チラシ入力
            MessageBox.Show("チラシ1入力しました。")
        Else
            MessageBox.Show(".TABファイルのみを入力してください。")
        End If

        Chirashi1TxtBx.Text = filePathData.Path
        uploadToDB("Region1", filePathData.Path)
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

        If filePathData.Flag Then
            'チラシ入力
            MessageBox.Show("チラシ2入力しました。")
        Else
            MessageBox.Show(".TABファイルのみを入力してください。")
        End If

        Chirashi2TxtBx.Text = filePathData.Path
        uploadToDB("Region2", filePathData.Path)
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

        If filePathData.Flag Then
            'チラシ入力
            MessageBox.Show("チラシ3入力しました。")
        Else
            MessageBox.Show(".TABファイルのみを入力してください。")
        End If

        Chirashi3TxtBx.Text = filePathData.Path
        uploadToDB("Region3", filePathData.Path)
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
        TextBoxEmptyValidation()
    End Sub

    Private Sub ShareWPCBtn_Click(sender As Object, e As EventArgs) Handles ShareWPCBtn.Click
        TextBoxEmptyValidation()
    End Sub

    Private Sub CreatePolygonBtn_Click(sender As Object, e As EventArgs) Handles CreatePolygonBtn.Click
        TextBoxEmptyValidation()
    End Sub

    Private Sub CreateJisseiBtn_Click(sender As Object, e As EventArgs) Handles CreateJisseiBtn.Click
        TextBoxEmptyValidation()
    End Sub

    Private Sub IryouUriTxtBx_TextChanged(sender As Object, e As EventArgs) Handles IryouUriTxtBx.TextChanged
        IryouUriTxtBx.TextAlign = HorizontalAlignment.Center
        FilledTxtBox(IryouUriTxtBx)
    End Sub

    Private Sub IryouUriTxtBx_Leave(sender As Object, e As EventArgs) Handles IryouUriTxtBx.Leave
        Dim blockCode As String = "1_衣料品"
        Dim parameterCode As String = "L"
        If flag Then
            UriageDB(blockCode, parameterCode, IryouUriTxtBx)
            flag = False
        Else
            flag = True
        End If

    End Sub

    Sub UriageDB(ByVal blockCode As String, ByVal parameterCode As String, ByVal getControl As Control)
        Conn.Close()

        Dim getNendo As Integer = Val(NenDoTxtBx.Text) - 1
        Dim getUri As Single = Val(getControl.Text)
        Dim cmd As SqlCommand = New SqlCommand("SELECT * FROM [dbo].[T_AM年間売上] WHERE 
            [店舗名称] = @TempoMei AND 
            [ブロックコード] = '" & blockCode & "' AND 
            [年度] = @NenDo AND
            [売上] = @Uriage ", Conn)
        cmd.Parameters.AddWithValue("@TempoMei", TempoMeiTxtBx.Text)
        cmd.Parameters.AddWithValue("@NenDo", getNendo)
        cmd.Parameters.AddWithValue("@Uriage", getControl.Text)
        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If

        Dim reader As SqlDataReader = cmd.ExecuteReader()
        If reader.HasRows Then

            '[店舗名称]と[FilePath]はDBに合わせる場合
            MsgBox("Already exist!")
        Else
            Conn.Close()
            cmd = New SqlCommand("SELECT * FROM [dbo].[T_AM年間売上] WHERE 
            [店舗名称] = @TempoMei AND 
            [ブロックコード] = '" & blockCode & "'", Conn)

            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            cmd.Parameters.AddWithValue("@TempoMei", TempoMeiTxtBx.Text)
            reader = cmd.ExecuteReader()
            If reader.HasRows Then
                Conn.Close()

                '[店舗名称]しかDBに合わさない場合
                cmd = New SqlCommand("UPDATE [dbo].[T_AM年間売上]
                SET [年度] = '" & getNendo & "' AND
                [売上] = '" & getControl.Text & "'
                WHERE [店舗名称] = '" & TempoMeiTxtBx.Text & "' AND [ブロックコード] = '" & blockCode & "'", Conn)

                If Conn.State = ConnectionState.Closed Then
                    Conn.Open()
                End If

                cmd.ExecuteNonQuery()
            Else
                Conn.Close()

                '[店舗名称]と[FilePath]はDBに合わさない場合
                cmd = New SqlCommand("INSERT INTO [dbo].[T_AM年間売上]
                      ([店舗名称]
                      ,[パターンコード]
                      ,[ブロックコード]
                      ,[年度]
                      ,[売上])
                       VALUES('" & TempoMeiTxtBx.Text & "','" & parameterCode & "','" & blockCode & "'," & getNendo & "," & getControl.Text & ")", Conn)
                If Conn.State = ConnectionState.Closed Then
                    Conn.Open()
                End If

                cmd.ExecuteNonQuery()
            End If
        End If
    End Sub

    Private Sub JukyoUriTxtBx_TextChanged(sender As Object, e As EventArgs) Handles JukyoUriTxtBx.TextChanged
        JukyoUriTxtBx.TextAlign = HorizontalAlignment.Center
        FilledTxtBox(JukyoUriTxtBx)
    End Sub

    Private Sub ShokuUriTxtBx_TextChanged(sender As Object, e As EventArgs) Handles ShokuUriTxtBx.TextChanged
        ShokuUriTxtBx.TextAlign = HorizontalAlignment.Center
        FilledTxtBox(ShokuUriTxtBx)
    End Sub

    Private Sub HibupUriTxtBx_TextChanged(sender As Object, e As EventArgs) Handles HibupUriTxtBx.TextChanged
        HibupUriTxtBx.TextAlign = HorizontalAlignment.Center
        FilledTxtBox(HibupUriTxtBx)
    End Sub

    'Validation for 全角・半角 ( StrConv() function can only used in Japan )
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Dim cmd As SqlCommand = New SqlCommand("SELECT * FROM [dbo].[T_AMTourokuFile] WHERE [店舗名称] = @TempoMei AND [種別] = 'Iryouhin'", Conn)
        'cmd.Parameters.AddWithValue("@TempoMei", TempoMeiTxtBx.Text)
        '
        'Dim reader As SqlDataReader = cmd.ExecuteReader()
        'If reader.HasRows Then
        '    MsgBox("Already exist!")
        'Else
        '    MsgBox("Not exist!")
        'End If
        ResultShare("Aeon", "D:\DBS\PT. ARI\CSAR\CSAR New Jobs\test files\Save")
    End Sub


End Class
