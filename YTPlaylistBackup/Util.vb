Imports System.IO
Imports System.IO.Compression
Imports Microsoft.VisualBasic.FileIO

Module Util

#Region "Variables"
    Public Property ClientID As String = Configuration.ConfigurationManager.AppSettings("ClientID")
    Public Property ClientSecret As String = Configuration.ConfigurationManager.AppSettings("ClientSecret")
    Public Property SqlConn As String = Configuration.ConfigurationManager.AppSettings("SqlConnection")
    Public Property BackupPath As String = Configuration.ConfigurationManager.AppSettings("BackupPath")
    Public Property TempPath As String = Configuration.ConfigurationManager.AppSettings("TempPath")
    Public Property PlaylistFileName As String = "Playlist.csv"
    Public Property PlaylistItemsFileName As String = "PlaylistItems.csv"
    Public Property PlaylistItemsRecoveredFileName As String = "PlaylistItemsRecovered.csv"
    Public Property PlaylistItemsRemovedFileName As String = "PlaylistItemsRemoved.csv"
    Public Property PlaylistItemsLostFileName As String = "PlaylistItemsLost.csv"
    Public Property SyncHistoryFileName As String = "SyncHistory.csv"
#End Region

#Region "Sub/Func"
    Public Sub SetForm(ByRef Form As Form, Optional Resizable As Boolean = False)
        Form.AutoScroll = True
        Form.AutoScrollMargin = New System.Drawing.Size(0, 7)
        If Not Resizable Then
            Form.AutoSize = True
            Form.FormBorderStyle = FormBorderStyle.FixedDialog
            Form.MaximizeBox = False
        End If
    End Sub

    Public Sub InitDGV(ByRef DGV As DataGridView)
        DGV.ReadOnly = True
        DGV.AllowUserToAddRows = False
        DGV.AllowUserToDeleteRows = False
        DGV.AllowUserToOrderColumns = False
        DGV.AllowUserToResizeRows = False
        DGV.RowHeadersWidth = 35
        DGV.DataSource = Nothing
        DGV.Rows.Clear()
        DGV.Columns.Clear()
    End Sub

    Public Sub SetFilterDataGridViewData(ByVal _Data As DataTable, ByRef _DataGridView As DataGridView, Optional _CustomFilter As String = "", Optional _Sort As String = "")
        Dim DataFiltered = _Data.DefaultView

        DataFiltered.RowFilter = If(_CustomFilter <> "", _CustomFilter, "")

        If _Sort <> "" Then
            DataFiltered.Sort = _Sort
        End If

        _DataGridView.DataSource = DataFiltered
    End Sub

    Public Function CheckDGVCellValue(ByRef CellValue As Object) As Boolean
        If CellValue IsNot Nothing AndAlso CellValue IsNot DBNull.Value AndAlso Not String.IsNullOrEmpty(CellValue.ToString) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub InitCombobox(ByRef cmbBox As ComboBox)
        cmbBox.DropDownStyle = ComboBoxStyle.DropDownList
        cmbBox.FlatStyle = FlatStyle.Popup
        cmbBox.DisplayMember = "Value"
        cmbBox.ValueMember = "Key"
    End Sub

    Public Sub OpenLink(link As String)
        Dim process As New Process
        process.StartInfo.UseShellExecute = True
        process.StartInfo.FileName = link
        process.Start()
    End Sub

    Public Sub ResetFirstDisplayedRowIndex(dataGridView As DataGridView)
        If dataGridView.Rows.Count > 0 Then
            dataGridView.FirstDisplayedScrollingRowIndex = 0
        End If
    End Sub

    Public Sub ExportDataTableToCSV(dt As DataTable, filePath As String)
        Using writer As New StreamWriter(filePath)
            For i As Integer = 0 To dt.Columns.Count - 1
                writer.Write(dt.Columns(i).ColumnName)
                If i < dt.Columns.Count - 1 Then
                    writer.Write(",")
                End If
            Next
            writer.WriteLine()

            For Each row As DataRow In dt.Rows
                For i As Integer = 0 To dt.Columns.Count - 1
                    If row(i).ToString.Contains(","c) Then
                        writer.Write("""" & row(i).ToString().Replace("""", """""") & """")
                    ElseIf row(i).ToString.Contains(Environment.NewLine) OrElse
                           row(i).ToString.Contains(vbCr) OrElse
                           row(i).ToString.Contains(vbLf) OrElse
                           row(i).ToString.Contains(vbCrLf) Then
                        writer.Write("""" & row(i).ToString().Replace("""", """""") & """")
                    ElseIf row(i).ToString.Contains(""""c) Then
                        writer.Write("""" & row(i).ToString().Replace("""", """""") & """")
                    Else
                        writer.Write(row(i).ToString())
                    End If
                    If i < dt.Columns.Count - 1 Then
                        writer.Write(",")
                    End If
                Next
                writer.WriteLine()
            Next
        End Using
    End Sub

    Public Sub CompressFilesToZip(sourceFilePaths As List(Of String), zipFilePath As String)
        Directory.CreateDirectory(Path.GetDirectoryName(zipFilePath))

        Using archive As ZipArchive = ZipFile.Open(zipFilePath, ZipArchiveMode.Create)
            For Each filePath In sourceFilePaths
                archive.CreateEntryFromFile(filePath, Path.GetFileName(filePath), CompressionLevel.SmallestSize)
            Next
        End Using
    End Sub

    Public Sub DeleteCSVFiles(sourceFilePaths As List(Of String))
        For Each filePath In sourceFilePaths
            Try
                If File.Exists(filePath) Then
                    File.Delete(filePath)
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Next
    End Sub

    Public Sub MoveCSVZip(sourceFilePath As String, destinationFolderPath As String)
        If File.Exists(sourceFilePath) Then
            Dim fileName As String = Path.GetFileName(sourceFilePath)
            Dim destinationFilePath As String = Path.Combine(destinationFolderPath, fileName)
            File.Move(sourceFilePath, destinationFilePath)
        End If
    End Sub

    Public Function ExtractCsvFromZip(ByVal zipFilePath As String, ByVal csvFileName As String) As String
        Dim csvData As String = ""

        Using archive As ZipArchive = ZipFile.OpenRead(zipFilePath)
            Dim entry As ZipArchiveEntry = archive.GetEntry(csvFileName)
            If entry IsNot Nothing Then
                Using reader As StreamReader = New StreamReader(entry.Open())
                    csvData = reader.ReadToEnd()
                End Using
            Else
                Console.WriteLine("CSV file not found in the ZIP archive.")
            End If
        End Using

        Return csvData
    End Function

    Public Function CsvToDataTable(ByVal csvData As String) As DataTable
        Dim dataTable As New DataTable()

        Using parser As New TextFieldParser(New StringReader(csvData))
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters(",")

            ' Read column names
            Dim columnNames As String() = parser.ReadFields()
            For Each columnName In columnNames
                dataTable.Columns.Add(columnName)
            Next

            ' Read data rows
            While Not parser.EndOfData
                Dim fields As String() = parser.ReadFields()
                dataTable.Rows.Add(fields)
            End While
        End Using

        Return dataTable
    End Function

    Public Function FindMissingRows(sourceTable As DataTable, targetTable As DataTable, primaryKeyColumns As Integer()) As DataTable
        Dim missingRows = From sourceRow In sourceTable.AsEnumerable()
                          Where Not targetTable.AsEnumerable().Any(Function(targetRow) primaryKeyColumns.All(Function(columnName) targetRow(columnName).ToString.Equals(sourceRow(columnName).ToString)))
                          Select sourceRow

        Dim missingRowsTable As New DataTable

        If missingRows.Any Then
            missingRowsTable = missingRows.CopyToDataTable()
        End If

        Return missingRowsTable
    End Function

    Public Sub SetAutoSizeModeForColumns(dgv As DataGridView)
        For Each column As DataGridViewColumn In dgv.Columns
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        Next
    End Sub

#End Region

End Module
