Imports System.IO
Imports System.Threading
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Services
Imports Google.Apis.Util.Store
Imports Google.Apis.YouTube.v3
Imports Google.Apis.YouTube.v3.Data

Public Class Form1
    'reference https://stackoverflow.com/questions/65357223/get-youtube-channel-data-using-google-youtube-data-api-in-vb-net

#Region "Variables"
    Private credential As UserCredential
    Private ytService As YouTubeService
    Private SearchedDatatable As DataTables = -1
#End Region

#Region "Form Events"
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetForm(Me, True)
        Timer1.Interval = 300
        CenterToScreen()

        Database.GetSqlAll()
        SetDGVData()

        SetComboboxValues(-1)
    End Sub

    Private Async Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        If MessageBox.Show("Are you sure?", "Expensive API Call", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            ToolStripButton3.Enabled = False

            ToolStripProgressBar1.Visible = True
            ToolStripProgressBar1.Value = 0
            ToolStripProgressBar1.Maximum = Database.playlistListData.Rows.Count + 2

            Await OAuth()
            GetPlaylistList()
            GetPlaylistItemLists()

            Database.GetSqlAll()
            SetDGVData()
            SetComboboxValues(-1)
            ToolStripProgressBar1.Value += 1
            ToolStripProgressBar1.Visible = False
            MessageBox.Show("Sync Completed", "Complete")
            ToolStripButton3.Enabled = True
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub ImportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportToolStripMenuItem.Click
        If MessageBox.Show("Are you sure?", "Importing Data from CSV", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            ImportToolStripMenuItem.Enabled = False
            ImportFromCSV()
            ImportToolStripMenuItem.Enabled = True
        End If
    End Sub

    Private Sub ExportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportToolStripMenuItem.Click
        If MessageBox.Show("Are you sure?" & Environment.NewLine & "Exporting file to " & BackupPath, "Export to CSV", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            ExportToolStripMenuItem.Enabled = False
            ExportToCSV()
            ExportToolStripMenuItem.Enabled = True
        End If
    End Sub

    Private Sub StatsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StatsToolStripMenuItem.Click
        Stats.ShowDialog()
    End Sub
#End Region

#Region "YT API"
    Private Async Function OAuth() As Task
        Dim scopes As IList(Of String) = New List(Of String) From {
            YouTubeService.Scope.Youtube
        }
        Dim clientsercret As New ClientSecrets With {
            .ClientId = ClientID,
            .ClientSecret = ClientSecret
        }
        credential = Await GoogleWebAuthorizationBroker.AuthorizeAsync(
           clientsercret,
           scopes,
           "user",
           CancellationToken.None,
           New FileDataStore([GetType].ToString)
        )

        ytService = New YouTubeService(New BaseClientService.Initializer() With {
            .HttpClientInitializer = credential,
            .ApplicationName = [GetType].ToString
        })
    End Function

    Private Sub GetPlaylistList()
        'Process Regular Playlists
        ProcessPlaylists(ytService, "", True, "")
        'Process Liked Playlist
        ProcessPlaylists(ytService, "", False, "LL")

        Database.GetSqlPlaylistList()
        ToolStripProgressBar1.Value += 1
    End Sub

    Private Sub ProcessPlaylists(ytService As YouTubeService, ByVal pageToken As String, ByVal isMine As Boolean, ByVal playlistId As String)
        Dim playlistCache As New Dictionary(Of String, String)

        Do
            Dim playlistListReq = ytService.Playlists.List("contentDetails,id,snippet")
            playlistListReq.MaxResults = 50
            playlistListReq.PageToken = pageToken

            If isMine Then
                playlistListReq.Mine = True
            ElseIf Not String.IsNullOrEmpty(playlistId) Then
                playlistListReq.Id = playlistId
            End If

            Dim playlistListResp = playlistListReq.Execute

            For Each playlist In playlistListResp.Items
                Dim id = playlist.Id
                Dim title = playlist.Snippet.Title
                Dim desc = playlist.Snippet.Description
                Dim itemCount = playlist.ContentDetails.ItemCount

                Dim filterRow = Database.playlistListData.AsEnumerable _
                .Where(Function(dr) dr(0).ToString = id).FirstOrDefault

                If filterRow Is Nothing Then
                    Database.InsertSqlPlaylistList(id, title, desc, itemCount)
                    ToolStripProgressBar1.Maximum += 1
                ElseIf filterRow(1) <> title OrElse filterRow(2) <> desc OrElse filterRow(3) <> itemCount Then
                    Database.UpdateSqlPlaylist(id, title, desc, itemCount)
                End If

                If Not playlistCache.ContainsKey(id) Then
                    playlistCache.Add(id, title)
                End If
            Next

            pageToken = playlistListResp.NextPageToken
        Loop While Not String.IsNullOrEmpty(pageToken)

        If isMine Then
            'removed
            For Each rows In Database.playlistListData.Rows
                If Not playlistCache.ContainsKey(rows(0)) AndAlso Not rows(0).ToString = "LL" Then
                    Database.DeleteSqlPlaylist(rows(0))
                End If
            Next
        End If
    End Sub

    Private Sub GetPlaylistItemLists()
        Dim playlistItemsCache As New Dictionary(Of String, String)

        Dim AddedCount = 0
        Dim RemovedCount = 0
        Dim RecoveredCount = 0
        Dim LostCount = 0

        For Each playlist In Database.playlistListData.Rows
            Dim playlistId = playlist(0)
            Dim nextPageToken As String = ""
            playlistItemsCache.Clear()

            While nextPageToken IsNot Nothing
                Dim playlistItemListReq = ytService.PlaylistItems.List("contentDetails,snippet,status")
                playlistItemListReq.PlaylistId = playlistId
                playlistItemListReq.MaxResults = 50
                playlistItemListReq.PageToken = nextPageToken

                Dim playlistItemListResp = playlistItemListReq.Execute
                For Each playlistItem In playlistItemListResp.Items
                    Dim videoId = playlistItem.ContentDetails.VideoId
                    Dim title = playlistItem.Snippet.Title
                    Dim desc = playlistItem.Snippet.Description
                    Dim videoOwnerChannelId = playlistItem.Snippet.VideoOwnerChannelId
                    Dim videoOwnerChannelTitle = playlistItem.Snippet.VideoOwnerChannelTitle
                    Dim status = playlistItem.Status.PrivacyStatus

                    Dim filterRow = Database.playlistItemListData.AsEnumerable.Where(Function(dr) dr(0).ToString = playlistId AndAlso dr(1).ToString = videoId).FirstOrDefault

                    If videoId <> "" AndAlso videoOwnerChannelId Is Nothing AndAlso
                        videoOwnerChannelTitle Is Nothing AndAlso
                        (status = "privacyStatusUnspecified" OrElse status = "private" OrElse
                        title = "Deleted video" OrElse title = "Private video" OrElse
                        desc = "This video is unavailable." OrElse desc = "This video is private.") Then
                        If filterRow IsNot Nothing Then
                            'recovered
                            Database.InsertSqlPlaylistItemRecoveredList(playlistId, filterRow(1), filterRow(2), filterRow(3), filterRow(4), filterRow(5))
                            Database.DeleteSqlPlaylistItem(filterRow(1), playlistId)
                            RecoveredCount += 1
                        Else
                            'check if not in recovered and not in lost
                            Dim filterRowRecover = Database.playlistItemListRecoveredData.AsEnumerable.Where(Function(dr) dr(0).ToString = playlistId AndAlso dr(1).ToString = videoId).FirstOrDefault
                            Dim filterRowLost = Database.playlistItemListLostData.AsEnumerable.Where(Function(dr) dr(0).ToString = playlistId AndAlso dr(1).ToString = videoId).FirstOrDefault
                            If filterRowRecover Is Nothing AndAlso filterRowLost Is Nothing Then
                                'no backup
                                Database.InsertSqlPlaylistItemLostList(playlistId, videoId, title, desc, videoOwnerChannelId, videoOwnerChannelTitle)
                                LostCount += 1
                            End If
                        End If
                    Else
                        If filterRow Is Nothing Then
                            'new added
                            Database.InsertSqlPlaylistItemList(playlistId, videoId, title, desc, videoOwnerChannelId, videoOwnerChannelTitle)
                            AddedCount += 1
                        End If
                    End If

                    If Not playlistItemsCache.ContainsKey(videoId) Then
                        playlistItemsCache.Add(videoId, title)
                    End If
                Next

                nextPageToken = playlistItemListResp.NextPageToken
            End While

            'removed
            Dim filteredTable = Database.playlistItemListData.AsEnumerable.Where(Function(dr) dr(0).ToString = playlistId).ToList
            If filteredTable IsNot Nothing Then
                For Each rows In filteredTable
                    If Not playlistItemsCache.ContainsKey(rows(1)) Then
                        Database.InsertSqlPlaylistItemRemovedList(playlistId, rows(1), rows(2), rows(3), rows(4), rows(5))
                        Database.DeleteSqlPlaylistItem(rows(1), playlistId)
                        RemovedCount += 1
                    End If
                Next
            End If
            ToolStripProgressBar1.Value += 1
        Next

        Database.InsertSqlSyncHistory(AddedCount, RemovedCount, RecoveredCount, LostCount, "Synced")
    End Sub
#End Region

#Region "Sub/Func"
    Private Sub SetDGVData()
        InitDGV(DataGridView1)
        InitDGV(DataGridView2)
        InitDGV(DataGridView3)
        InitDGV(DataGridView4)
        InitDGV(DataGridView5)
        InitDGV(DataGridView6)

        SetFilterDataGridViewData(Database.playlistListData, DataGridView1, _Sort:="syncDate DESC")
        SetFilterDataGridViewData(Database.playlistItemListData, DataGridView2, _Sort:="syncDate DESC")
        SetFilterDataGridViewData(Database.playlistItemListRecoveredData, DataGridView3, _Sort:="syncDate DESC")
        SetFilterDataGridViewData(Database.playlistItemListRemovedData, DataGridView4, _Sort:="syncDate DESC")
        SetFilterDataGridViewData(Database.playlistItemListLostData, DataGridView5, _Sort:="syncDate DESC")
        SetFilterDataGridViewData(Database.syncHistoryData, DataGridView6, _Sort:="syncDate DESC")

        Dim linkPlaylist As New DataGridViewLinkColumn With {
            .HeaderText = "YT Link",
            .TrackVisitedState = False
        }
        DataGridView1.Columns.Insert(0, linkPlaylist)

        Dim linkPlaylistList As New DataGridViewLinkColumn With {
            .HeaderText = "YT Link",
            .TrackVisitedState = False
        }
        DataGridView2.Columns.Insert(0, linkPlaylistList)

        Dim linkPlaylistListRecovered As New DataGridViewLinkColumn With {
            .HeaderText = "YT Link",
            .TrackVisitedState = False
        }
        DataGridView3.Columns.Insert(0, linkPlaylistListRecovered)

        Dim linkPlaylistListRemoved As New DataGridViewLinkColumn With {
            .HeaderText = "YT Link",
            .TrackVisitedState = False
        }
        DataGridView4.Columns.Insert(0, linkPlaylistListRemoved)

        Dim linkPlaylistListLost As New DataGridViewLinkColumn With {
            .HeaderText = "YT Link",
            .TrackVisitedState = False
        }
        DataGridView5.Columns.Insert(0, linkPlaylistListLost)

        SetRowValue(-1)

        DataGridView1.Columns(1).Visible = False
        DataGridView2.Columns(2).Visible = False
        DataGridView3.Columns(2).Visible = False
        DataGridView4.Columns(2).Visible = False
        DataGridView5.Columns(2).Visible = False

        DataGridView2.Columns(1).HeaderText = "Playlist"
        DataGridView3.Columns(1).HeaderText = "Playlist"
        DataGridView4.Columns(1).HeaderText = "Playlist"
        DataGridView5.Columns(1).HeaderText = "Playlist"

        ResetFirstDisplayedRowIndex(DataGridView1)
        ResetFirstDisplayedRowIndex(DataGridView2)
        ResetFirstDisplayedRowIndex(DataGridView3)
        ResetFirstDisplayedRowIndex(DataGridView4)
        ResetFirstDisplayedRowIndex(DataGridView5)
        ResetFirstDisplayedRowIndex(DataGridView6)

        SetAutoSizeModeForColumns(DataGridView1)
        SetAutoSizeModeForColumns(DataGridView2)
        SetAutoSizeModeForColumns(DataGridView3)
        SetAutoSizeModeForColumns(DataGridView4)
        SetAutoSizeModeForColumns(DataGridView5)
        SetAutoSizeModeForColumns(DataGridView6)

        DataGridView1.Refresh()
        DataGridView2.Refresh()
        DataGridView3.Refresh()
        DataGridView4.Refresh()
        DataGridView5.Refresh()
        DataGridView6.Refresh()
    End Sub

    Private Sub SetRowValue(data As Integer)
        '-1 = all else use enum DataTables

        If data = -1 OrElse data = DataTables.playlistList Then
            For Each rows As DataGridViewRow In DataGridView1.Rows
                If CheckDGVCellValue(rows.Cells(1).Value) Then
                    rows.Cells(0).Value = "https://www.youtube.com/playlist?list=" & rows.Cells(1).Value
                End If
            Next
        End If

        If data = -1 OrElse data = DataTables.playlistItemList Then
            For Each rows As DataGridViewRow In DataGridView2.Rows
                If CheckDGVCellValue(rows.Cells(2).Value) Then
                    rows.Cells(0).Value = "https://www.youtube.com/watch?v=" & rows.Cells(2).Value
                End If
            Next
        End If

        If data = -1 OrElse data = DataTables.playlistItemListRecovered Then
            For Each rows As DataGridViewRow In DataGridView3.Rows
                If CheckDGVCellValue(rows.Cells(2).Value) Then
                    rows.Cells(0).Value = "https://www.youtube.com/watch?v=" & rows.Cells(2).Value
                End If
            Next
        End If

        If data = -1 OrElse data = DataTables.playlistItemListRemoved Then
            For Each rows As DataGridViewRow In DataGridView4.Rows
                If CheckDGVCellValue(rows.Cells(2).Value) Then
                    rows.Cells(0).Value = "https://www.youtube.com/watch?v=" & rows.Cells(2).Value
                End If
            Next
        End If

        If data = -1 OrElse data = DataTables.playlistItemListLost Then
            For Each rows As DataGridViewRow In DataGridView5.Rows
                If CheckDGVCellValue(rows.Cells(2).Value) Then
                    rows.Cells(0).Value = "https://www.youtube.com/watch?v=" & rows.Cells(2).Value
                End If
            Next
        End If
    End Sub

    Private Sub SetComboboxValues(data As Integer)
        '-1 = all else use enum DataTables

        Dim playlistList As New Dictionary(Of String, String) From {
            {"All Playlist", "All"}
        }

        For Each rows In Database.playlistListData.Rows
            If Not playlistList.ContainsKey(rows(0)) Then
                playlistList.Add(rows(0), rows(1))
            End If
        Next

        If data = -1 OrElse data = DataTables.playlistItemList Then
            InitCombobox(ComboBox1)
            ComboBox1.DataSource = playlistList.ToList
        End If

        If data = -1 OrElse data = DataTables.playlistItemListRecovered Then
            InitCombobox(ComboBox2)
            ComboBox2.DataSource = playlistList.ToList
        End If

        If data = -1 OrElse data = DataTables.playlistItemListRemoved Then
            InitCombobox(ComboBox3)
            ComboBox3.DataSource = playlistList.ToList
        End If

        If data = -1 OrElse data = DataTables.playlistItemListLost Then
            InitCombobox(ComboBox4)
            ComboBox4.DataSource = playlistList.ToList
        End If
    End Sub

    Private Sub ExportToCSV()
        Dim zipFileName = "ytPlaylistGalleryBak " & Now.ToString("yyyy-MM-dd_hh-mm-ss") & ".zip"
        Dim zipFilePath = BackupPath & zipFileName
        Dim playlistFilePath = BackupPath & PlaylistFileName
        Dim playlistItemsFilePath = BackupPath & PlaylistItemsFileName
        Dim playlistItemsRecoveredFilePath = BackupPath & PlaylistItemsRecoveredFileName
        Dim playlistItemsRemovedFilePath = BackupPath & PlaylistItemsRemovedFileName
        Dim playlistItemsLostFilePath = BackupPath & PlaylistItemsLostFileName
        Dim syncHistoryFilePath = BackupPath & SyncHistoryFileName
        Dim csvFiles As New List(Of String) From {playlistFilePath, playlistItemsFilePath, playlistItemsRecoveredFilePath, playlistItemsRemovedFilePath, playlistItemsLostFilePath, syncHistoryFilePath}

        Try
            If Not File.Exists(BackupPath & zipFileName) Then
                ExportDataTableToCSV(Database.playlistListData, playlistFilePath)
                ExportDataTableToCSV(Database.playlistItemListData, playlistItemsFilePath)
                ExportDataTableToCSV(Database.playlistItemListRecoveredData, playlistItemsRecoveredFilePath)
                ExportDataTableToCSV(Database.playlistItemListRemovedData, playlistItemsRemovedFilePath)
                ExportDataTableToCSV(Database.playlistItemListLostData, playlistItemsLostFilePath)
                ExportDataTableToCSV(Database.syncHistoryData, syncHistoryFilePath)

                CompressFilesToZip(csvFiles, zipFilePath)
                DeleteCSVFiles(csvFiles)
                MoveCSVZip(zipFilePath, BackupPath)
                MessageBox.Show("Saved: " & zipFileName & " to " & BackupPath, "Done Backup")
            Else
                MessageBox.Show("Backup already exists")
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub ImportFromCSV()
        If Directory.Exists(BackupPath) Then
            OpenFileDialog1.InitialDirectory = BackupPath
            OpenFileDialog1.FileName = ""
        End If
        OpenFileDialog1.Filter = "ZIP files (*.zip)|*.zip"
        If OpenFileDialog1.ShowDialog <> DialogResult.Cancel Then
            ToolStripProgressBar1.Value = 0
            ToolStripProgressBar1.Maximum = 25
            ToolStripProgressBar1.Visible = True

            Database.GetSqlAll()
            ToolStripProgressBar1.Value += 1
            SetDGVData()
            ToolStripProgressBar1.Value += 1

            Dim playlistCSVData As String = ExtractCsvFromZip(OpenFileDialog1.FileName, PlaylistFileName)
            ToolStripProgressBar1.Value += 1
            Dim playlistItemsCSVData As String = ExtractCsvFromZip(OpenFileDialog1.FileName, PlaylistItemsFileName)
            ToolStripProgressBar1.Value += 1
            Dim playlistItemsRecoveredCSVData As String = ExtractCsvFromZip(OpenFileDialog1.FileName, PlaylistItemsRecoveredFileName)
            ToolStripProgressBar1.Value += 1
            Dim playlistItemsRemovedCSVData As String = ExtractCsvFromZip(OpenFileDialog1.FileName, PlaylistItemsRemovedFileName)
            ToolStripProgressBar1.Value += 1
            Dim playlistItemsLostCSVData As String = ExtractCsvFromZip(OpenFileDialog1.FileName, PlaylistItemsLostFileName)
            ToolStripProgressBar1.Value += 1

            Dim playlistCSVDataTable As DataTable = CsvToDataTable(playlistCSVData)
            ToolStripProgressBar1.Value += 1
            Dim playlistItemsCSVDataTable As DataTable = CsvToDataTable(playlistItemsCSVData)
            ToolStripProgressBar1.Value += 1
            Dim playlistItemsRecoveredCSVDataTable As DataTable = CsvToDataTable(playlistItemsRecoveredCSVData)
            ToolStripProgressBar1.Value += 1
            Dim playlistItemsRemovedCSVDataTable As DataTable = CsvToDataTable(playlistItemsRemovedCSVData)
            ToolStripProgressBar1.Value += 1
            Dim playlistItemsLostCSVDataTable As DataTable = CsvToDataTable(playlistItemsLostCSVData)
            ToolStripProgressBar1.Value += 1

            Dim missingPlaylistData As DataTable = FindMissingRows(playlistCSVDataTable, Database.playlistListData, {0})
            ToolStripProgressBar1.Value += 1
            Dim missingPlaylistItemsData As DataTable = FindMissingRows(playlistItemsCSVDataTable, Database.playlistItemListData, {0, 1})
            ToolStripProgressBar1.Value += 1
            Dim missingPlaylistItemsRecoveredData As DataTable = FindMissingRows(playlistItemsRecoveredCSVDataTable, Database.playlistItemListRecoveredData, {0, 1})
            ToolStripProgressBar1.Value += 1
            Dim missingPlaylistItemsRemovedData As DataTable = FindMissingRows(playlistItemsRemovedCSVDataTable, Database.playlistItemListRemovedData, {0, 1})
            ToolStripProgressBar1.Value += 1
            Dim missingPlaylistItemsLostData As DataTable = FindMissingRows(playlistItemsLostCSVDataTable, Database.playlistItemListLostData, {0, 1})
            ToolStripProgressBar1.Value += 1

            If missingPlaylistData.Rows.Count > 0 Then
                For Each rows In missingPlaylistData.Rows
                    Database.InsertSqlPlaylistList(rows(0), rows(1), rows(2), rows(3))
                Next
            End If
            ToolStripProgressBar1.Value += 1

            If missingPlaylistItemsData.Rows.Count > 0 Then
                For Each rows In missingPlaylistItemsData.Rows
                    Database.InsertSqlPlaylistItemList(rows(0), rows(1), rows(2), rows(3), rows(4), rows(5))
                Next
            End If
            ToolStripProgressBar1.Value += 1

            If missingPlaylistItemsRecoveredData.Rows.Count > 0 Then
                For Each rows In missingPlaylistItemsRecoveredData.Rows
                    Database.InsertSqlPlaylistItemRecoveredList(rows(0), rows(1), rows(2), rows(3), rows(4), rows(5))
                Next
            End If
            ToolStripProgressBar1.Value += 1

            If missingPlaylistItemsRemovedData.Rows.Count > 0 Then
                For Each rows In missingPlaylistItemsRemovedData.Rows
                    Database.InsertSqlPlaylistItemRemovedList(rows(0), rows(1), rows(2), rows(3), rows(4), rows(5))
                Next
            End If
            ToolStripProgressBar1.Value += 1

            If missingPlaylistItemsLostData.Rows.Count > 0 Then
                For Each rows In missingPlaylistItemsLostData.Rows
                    Database.InsertSqlPlaylistItemLostList(rows(0), rows(1), rows(2), rows(3), rows(4), rows(5))
                Next
            End If
            ToolStripProgressBar1.Value += 1

            Dim anyDataImported As Boolean = {missingPlaylistData, missingPlaylistItemsData, missingPlaylistItemsRecoveredData, missingPlaylistItemsRemovedData, missingPlaylistItemsLostData}.Any(Function(dt) dt.Rows.Count > 0)
            If anyDataImported Then
                Database.InsertSqlSyncHistory(missingPlaylistItemsData.Rows.Count, missingPlaylistItemsRemovedData.Rows.Count, missingPlaylistItemsRecoveredData.Rows.Count, missingPlaylistItemsLostData.Rows.Count, "Imported")
            End If
            ToolStripProgressBar1.Value += 1
            Database.GetSqlAll()
            ToolStripProgressBar1.Value += 1
            SetDGVData()
            ToolStripProgressBar1.Value += 1
            MessageBox.Show(If(anyDataImported, "Done Import", "No Data Imported"))
        End If
        ToolStripProgressBar1.Visible = False
    End Sub

    Private Sub SetDGVFilter(Datatables As DataTables)
        Dim filter As String = ""
        Dim searchText As String = ""

        If Datatables = DataTables.playlistList Then
            searchText = TextBox1.Text.Trim
        ElseIf Datatables = DataTables.playlistItemList Then
            searchText = TextBox2.Text.Trim
        ElseIf Datatables = DataTables.playlistItemListRecovered Then
            searchText = TextBox3.Text.Trim
        ElseIf Datatables = DataTables.playlistItemListRemoved Then
            searchText = TextBox4.Text.Trim
        ElseIf Datatables = DataTables.playlistItemListLost Then
            searchText = TextBox5.Text.Trim
        ElseIf Datatables = DataTables.syncHistory Then
            searchText = TextBox6.Text.Trim
        End If

        If String.IsNullOrEmpty(searchText) Then
            If Datatables = DataTables.playlistList Then
                SetFilterDataGridViewData(Database.playlistListData, DataGridView1, filter)
                SetRowValue(DataTables.playlistList)
            ElseIf Datatables = DataTables.playlistItemList Then
                If Not ComboBox1.SelectedItem.Value = "All" Then
                    filter = String.Format("playlistID = '{0}'", ComboBox1.SelectedItem.Key)
                End If
                SetFilterDataGridViewData(Database.playlistItemListData, DataGridView2, filter)
                SetRowValue(DataTables.playlistItemList)
            ElseIf Datatables = DataTables.playlistItemListRecovered Then
                If Not ComboBox2.SelectedItem.Value = "All" Then
                    filter = String.Format("playlistID = '{0}'", ComboBox2.SelectedItem.Key)
                End If
                SetFilterDataGridViewData(Database.playlistItemListRecoveredData, DataGridView3, filter)
                SetRowValue(DataTables.playlistItemListRecovered)
            ElseIf Datatables = DataTables.playlistItemListRemoved Then
                If Not ComboBox3.SelectedItem.Value = "All" Then
                    filter = String.Format("playlistID = '{0}'", ComboBox3.SelectedItem.Key)
                End If
                SetFilterDataGridViewData(Database.playlistItemListRemovedData, DataGridView4, filter)
                SetRowValue(DataTables.playlistItemListRemoved)
            ElseIf Datatables = DataTables.playlistItemListLost Then
                If Not ComboBox4.SelectedItem.Value = "All" Then
                    filter = String.Format("playlistID = '{0}'", ComboBox4.SelectedItem.Key)
                End If
                SetFilterDataGridViewData(Database.playlistItemListLostData, DataGridView5, filter)
                SetRowValue(DataTables.playlistItemListLost)
            ElseIf Datatables = DataTables.syncHistory Then
                SetFilterDataGridViewData(Database.syncHistoryData, DataGridView6, filter)
                SetRowValue(DataTables.syncHistory)
            End If
        Else
            If Datatables = DataTables.playlistList Then
                filter = $"playlistID LIKE '%{searchText}%' 
                            OR playlistID LIKE '%{searchText.Replace("https://www.youtube.com/playlist?list=", "")}%' 
                            OR title LIKE '%{searchText}%' 
                            OR description LIKE '%{searchText}%'"

                Dim tempNumber As Integer
                If Integer.TryParse(searchText, tempNumber) Then
                    filter &= $" OR itemCount = {tempNumber}"
                End If

                Dim tempDate As Date
                If Date.TryParse(searchText, tempDate) Then
                    Dim startDate As Date = tempDate.Date
                    Dim endDate As Date = startDate.AddDays(1)
                    filter &= $" OR (syncDate >= #{startDate}# AND syncDate < #{endDate}#)"
                End If

                SetFilterDataGridViewData(Database.playlistListData, DataGridView1, filter)
                SetRowValue(DataTables.playlistList)
            ElseIf {DataTables.playlistItemList, DataTables.playlistItemListRecovered, DataTables.playlistItemListRemoved, DataTables.playlistItemListLost}.Contains(Datatables) Then
                filter = $"(title LIKE '%{searchText}%' 
                            OR description LIKE '%{searchText}%' 
                            OR videoID LIKE '%{searchText}%' 
                            OR videoID LIKE '%{searchText.Replace("https://www.youtube.com/watch?v=", "")}%' 
                            OR videoOwnerChannelId LIKE '%{searchText}%' 
                            OR videoOwnerChannelTitle LIKE '%{searchText}%'"

                Dim tempDate As Date
                If Date.TryParse(searchText, tempDate) Then
                    Dim startDate As Date = tempDate.Date
                    Dim endDate As Date = startDate.AddDays(1)
                    filter &= $" OR (syncDate >= #{startDate}# AND syncDate < #{endDate}#)"
                End If

                filter &= ")"

                If Datatables = DataTables.playlistItemList Then
                    If Not ComboBox1.SelectedItem.Value = "All" Then
                        filter &= $" AND playlistID = '{ComboBox1.SelectedItem.Key}'"
                    End If
                    SetFilterDataGridViewData(Database.playlistItemListData, DataGridView2, filter)
                    SetRowValue(DataTables.playlistItemList)
                ElseIf Datatables = DataTables.playlistItemListRecovered Then
                    If Not ComboBox2.SelectedItem.Value = "All" Then
                        filter &= $" AND playlistID = '{ComboBox2.SelectedItem.Key}'"
                    End If
                    SetFilterDataGridViewData(Database.playlistItemListRecoveredData, DataGridView3, filter)
                    SetRowValue(DataTables.playlistItemListRecovered)
                ElseIf Datatables = DataTables.playlistItemListRemoved Then
                    If Not ComboBox3.SelectedItem.Value = "All" Then
                        filter &= $" AND playlistID = '{ComboBox3.SelectedItem.Key}'"
                    End If
                    SetFilterDataGridViewData(Database.playlistItemListRemovedData, DataGridView4, filter)
                    SetRowValue(DataTables.playlistItemListRemoved)
                ElseIf Datatables = DataTables.playlistItemListLost Then
                    If Not ComboBox4.SelectedItem.Value = "All" Then
                        filter &= $" AND playlistID = '{ComboBox4.SelectedItem.Key}'"
                    End If
                    SetFilterDataGridViewData(Database.playlistItemListLostData, DataGridView5, filter)
                    SetRowValue(DataTables.playlistItemListLost)
                End If
            ElseIf Datatables = DataTables.syncHistory Then
                filter = $"Notes LIKE '%{searchText}%'"

                Dim tempNumber1 As Integer, tempNumber2 As Integer, tempNumber3 As Integer, tempNumber4 As Integer
                If Integer.TryParse(searchText, tempNumber1) Then
                    filter &= $" OR AddedCount = {tempNumber1}"
                End If
                If Integer.TryParse(searchText, tempNumber2) Then
                    filter &= $" OR RemovedCount = {tempNumber2}"
                End If
                If Integer.TryParse(searchText, tempNumber3) Then
                    filter &= $" OR RecoveredCount = {tempNumber3}"
                End If
                If Integer.TryParse(searchText, tempNumber4) Then
                    filter &= $" OR LostCount = {tempNumber4}"
                End If

                Dim tempDate As Date
                If Date.TryParse(searchText, tempDate) Then
                    Dim startDate As Date = tempDate.Date
                    Dim endDate As Date = startDate.AddDays(1)
                    filter &= $" OR (syncDate >= #{startDate}# AND syncDate < #{endDate}#)"
                End If

                SetFilterDataGridViewData(Database.syncHistoryData, DataGridView6, filter)
                SetRowValue(DataTables.syncHistory)
            End If
        End If
    End Sub
#End Region

#Region "Control Handlers"
    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""

        If TabControl1.SelectedIndex = 0 Then
            SetRowValue(DataTables.playlistList)
            ResetFirstDisplayedRowIndex(DataGridView1)
            DataGridView1.Refresh()
        ElseIf TabControl1.SelectedIndex = 1 Then
            SetRowValue(DataTables.playlistItemList)
            SetComboboxValues(DataTables.playlistItemList)
            ResetFirstDisplayedRowIndex(DataGridView2)
            DataGridView2.Refresh()
        ElseIf TabControl1.SelectedIndex = 2 Then
            SetRowValue(DataTables.playlistItemListRecovered)
            SetComboboxValues(DataTables.playlistItemListRecovered)
            ResetFirstDisplayedRowIndex(DataGridView3)
            DataGridView3.Refresh()
        ElseIf TabControl1.SelectedIndex = 3 Then
            SetRowValue(DataTables.playlistItemListRemoved)
            SetComboboxValues(DataTables.playlistItemListRemoved)
            ResetFirstDisplayedRowIndex(DataGridView4)
            DataGridView4.Refresh()
        ElseIf TabControl1.SelectedIndex = 4 Then
            SetRowValue(DataTables.playlistItemListLost)
            SetComboboxValues(DataTables.playlistItemListLost)
            ResetFirstDisplayedRowIndex(DataGridView5)
            DataGridView5.Refresh()
        ElseIf TabControl1.SelectedIndex = 5 Then
            ResetFirstDisplayedRowIndex(DataGridView6)
            DataGridView6.Refresh()
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim row = DataGridView1.Rows(e.RowIndex)
            If row.Cells(0).Value Is Nothing Then Return
            Dim url = row.Cells(0).Value.ToString()
            OpenLink(url)
        End If
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim row = DataGridView2.Rows(e.RowIndex)
            If row.Cells(0).Value Is Nothing Then Return
            Dim url = row.Cells(0).Value.ToString()
            OpenLink(url)
        End If
    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim row = DataGridView3.Rows(e.RowIndex)
            If row.Cells(0).Value Is Nothing Then Return
            Dim url = row.Cells(0).Value.ToString()
            OpenLink(url)
        End If
    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim row = DataGridView4.Rows(e.RowIndex)
            If row.Cells(0).Value Is Nothing Then Return
            Dim url = row.Cells(0).Value.ToString()
            OpenLink(url)
        End If
    End Sub

    Private Sub DataGridView5_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView5.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim row = DataGridView5.Rows(e.RowIndex)
            If row.Cells(0).Value Is Nothing Then Return
            Dim url = row.Cells(0).Value.ToString()
            OpenLink(url)
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        SetDGVFilter(DataTables.playlistItemList)
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        SetDGVFilter(DataTables.playlistItemListRecovered)
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        SetDGVFilter(DataTables.playlistItemListRemoved)
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        SetDGVFilter(DataTables.playlistItemListLost)
    End Sub

    Private Sub DataGridView1_Sorted(sender As Object, e As EventArgs) Handles DataGridView1.Sorted
        SetRowValue(DataTables.playlistList)
    End Sub

    Private Sub DataGridView2_Sorted(sender As Object, e As EventArgs) Handles DataGridView2.Sorted
        SetRowValue(DataTables.playlistItemList)
    End Sub

    Private Sub DataGridView3_Sorted(sender As Object, e As EventArgs) Handles DataGridView3.Sorted
        SetRowValue(DataTables.playlistItemListRecovered)
    End Sub

    Private Sub DataGridView4_Sorted(sender As Object, e As EventArgs) Handles DataGridView4.Sorted
        SetRowValue(DataTables.playlistItemListRemoved)
    End Sub

    Private Sub DataGridView5_Sorted(sender As Object, e As EventArgs) Handles DataGridView5.Sorted
        SetRowValue(DataTables.playlistItemListLost)
    End Sub

    Private Sub DataGridView2_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting
        If e.ColumnIndex = 1 AndAlso e.RowIndex >= 0 Then
            Dim playlistId As String = DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            Dim filterRow = Database.playlistListData.AsEnumerable.Where(Function(row) row(0).ToString = playlistId).FirstOrDefault
            If filterRow IsNot Nothing Then
                playlistId = filterRow(1)
            End If
            e.Value = playlistId 'change displayed value
            e.FormattingApplied = True
        End If
    End Sub

    Private Sub DataGridView3_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView3.CellFormatting
        If e.ColumnIndex = 1 AndAlso e.RowIndex >= 0 Then
            Dim playlistId As String = DataGridView3.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            Dim filterRow = Database.playlistListData.AsEnumerable.Where(Function(row) row(0).ToString = playlistId).FirstOrDefault
            If filterRow IsNot Nothing Then
                playlistId = filterRow(1)
            End If
            e.Value = playlistId 'change displayed value
            e.FormattingApplied = True
        End If
    End Sub

    Private Sub DataGridView4_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView4.CellFormatting
        If e.ColumnIndex = 1 AndAlso e.RowIndex >= 0 Then
            Dim playlistId As String = DataGridView4.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            Dim filterRow = Database.playlistListData.AsEnumerable.Where(Function(row) row(0).ToString = playlistId).FirstOrDefault
            If filterRow IsNot Nothing Then
                playlistId = filterRow(1)
            End If
            e.Value = playlistId 'change displayed value
            e.FormattingApplied = True
        End If
    End Sub

    Private Sub DataGridView5_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView5.CellFormatting
        If e.ColumnIndex = 1 AndAlso e.RowIndex >= 0 Then
            Dim playlistId As String = DataGridView5.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            Dim filterRow = Database.playlistListData.AsEnumerable.Where(Function(row) row(0).ToString = playlistId).FirstOrDefault
            If filterRow IsNot Nothing Then
                playlistId = filterRow(1)
            End If
            e.Value = playlistId 'change displayed value
            e.FormattingApplied = True
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        SearchedDatatable = DataTables.playlistList
        Timer1.Stop()
        Timer1.Start()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        SearchedDatatable = DataTables.playlistItemList
        Timer1.Stop()
        Timer1.Start()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        SearchedDatatable = DataTables.playlistItemListRecovered
        Timer1.Stop()
        Timer1.Start()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        SearchedDatatable = DataTables.playlistItemListRemoved
        Timer1.Stop()
        Timer1.Start()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        SearchedDatatable = DataTables.playlistItemListLost
        Timer1.Stop()
        Timer1.Start()
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        SearchedDatatable = DataTables.syncHistory
        Timer1.Stop()
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Stop()

        If SearchedDatatable = -1 Then
            Return
        End If

        SetDGVFilter(SearchedDatatable)

        SearchedDatatable = -1
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            ' Suppress the ding sound
            e.Handled = True
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            ' Suppress the ding sound
            e.Handled = True
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            ' Suppress the ding sound
            e.Handled = True
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            ' Suppress the ding sound
            e.Handled = True
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            ' Suppress the ding sound
            e.Handled = True
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            ' Suppress the ding sound
            e.Handled = True
            e.SuppressKeyPress = True
        End If
    End Sub
#End Region
End Class
