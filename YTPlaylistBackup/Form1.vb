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
#End Region

#Region "Form Events"
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Util.SetForm(Me)
        CenterToScreen()

        Database.InitDatatables()
        Database.GetSqlPlaylistList()
        Database.GetSqlPlaylistItemsList()
        Database.GetSqlPlaylistItemsRecoveredList()
        Database.GetSqlPlaylistItemsRemovedList()
        Database.GetSqlPlaylistItemsLostList()
        Database.GetSqlSyncHistory()
        SetDGVData()

        SetComboboxValues(-1)
    End Sub

    Private Async Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        If MessageBox.Show("Are you sure?", "Expensive API Call", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            ToolStripButton3.Enabled = False

            Await OAuth()
            GetPlaylistList()
            GetPlaylistItemLists()

            Database.InitDatatables()
            Database.GetSqlPlaylistList()
            Database.GetSqlPlaylistItemsList()
            Database.GetSqlPlaylistItemsRecoveredList()
            Database.GetSqlPlaylistItemsRemovedList()
            Database.GetSqlPlaylistItemsLostList()
            Database.GetSqlSyncHistory()
            SetDGVData()

            SetComboboxValues(-1)

            MessageBox.Show("Sync Completed", "Complete")
            ToolStripButton3.Enabled = True
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
#End Region

#Region "YT API"
    Private Async Function OAuth() As Task
        Dim scopes As IList(Of String) = New List(Of String) From {
            YouTubeService.Scope.Youtube
        }
        Dim clientsercret As New ClientSecrets With {
            .ClientId = Util.ClientID,
            .ClientSecret = Util.ClientSecret
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
        Dim nextPageToken As String = ""
        Dim playlistCache As New Dictionary(Of String, String)

        While nextPageToken IsNot Nothing
            Dim playlistListReq = ytService.Playlists.List("contentDetails,id,snippet")
            playlistListReq.Mine = True
            playlistListReq.MaxResults = 50
            playlistListReq.PageToken = nextPageToken

            Dim playlistListResp = playlistListReq.Execute
            For Each playlists In playlistListResp.Items
                Dim id = playlists.Id
                Dim title = playlists.Snippet.Title
                Dim desc = playlists.Snippet.Description
                Dim itemCount = playlists.ContentDetails.ItemCount

                Dim filterRow = Database.playlistListData.AsEnumerable.Where(Function(dr) dr(0).ToString = id).FirstOrDefault
                If filterRow Is Nothing Then
                    Database.InsertSqlPlaylistList(id, title, desc, itemCount)
                Else
                    If filterRow(1) <> title OrElse filterRow(2) <> desc OrElse filterRow(3) <> itemCount Then
                        Database.UpdateSqlPlaylist(id, title, desc, itemCount)
                    End If
                End If

                If Not playlistCache.ContainsKey(id) Then
                    playlistCache.Add(id, title)
                End If
            Next

            nextPageToken = playlistListResp.NextPageToken
        End While

        'removed
        For Each rows In Database.playlistListData.Rows
            If Not playlistCache.ContainsKey(rows(0)) Then
                Database.DeleteSqlPlaylist(rows(0))
            End If
        Next

        Database.GetSqlPlaylistList()
    End Sub

    Private Sub GetPlaylistItemLists()
        Dim playlistItemsCache As New Dictionary(Of String, String)

        Dim AddedCount = 0
        Dim RemovedCount = 0
        Dim RecoveredCount = 0
        Dim LostCount = 0

        ToolStripProgressBar1.Visible = True
        ToolStripProgressBar1.Maximum = Database.playlistListData.Rows.Count
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
            ToolStripProgressBar1.Value = ToolStripProgressBar1.Value + 1
        Next

        Database.InsertSqlSyncHistory(AddedCount, RemovedCount, RecoveredCount, LostCount)
        ToolStripProgressBar1.Visible = False
    End Sub
#End Region

#Region "Sub/Func"
    Private Sub SetDGVData()
        Util.InitDGV(DataGridView1)
        Util.InitDGV(DataGridView2)
        Util.InitDGV(DataGridView3)
        Util.InitDGV(DataGridView4)
        Util.InitDGV(DataGridView5)
        Util.InitDGV(DataGridView6)

        Util.SetFilterDataGridViewData(Database.playlistListData, DataGridView1, _Sort:="syncDate DESC")
        Util.SetFilterDataGridViewData(Database.playlistItemListData, DataGridView2, _Sort:="syncDate DESC")
        Util.SetFilterDataGridViewData(Database.playlistItemListRecoveredData, DataGridView3, _Sort:="syncDate DESC")
        Util.SetFilterDataGridViewData(Database.playlistItemListRemovedData, DataGridView4, _Sort:="syncDate DESC")
        Util.SetFilterDataGridViewData(Database.playlistItemListLostData, DataGridView5, _Sort:="syncDate DESC")
        Util.SetFilterDataGridViewData(Database.syncHistoryData, DataGridView6, _Sort:="syncDate DESC")

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
                If Util.CheckDGVCellValue(rows.Cells(1).Value) Then
                    rows.Cells(0).Value = "https://www.youtube.com/playlist?list=" & rows.Cells(1).Value
                End If
            Next
        End If

        If data = -1 OrElse data = DataTables.playlistItemList Then
            For Each rows As DataGridViewRow In DataGridView2.Rows
                If Util.CheckDGVCellValue(rows.Cells(2).Value) Then
                    rows.Cells(0).Value = "https://www.youtube.com/watch?v=" & rows.Cells(2).Value
                End If
            Next
        End If

        If data = -1 OrElse data = DataTables.playlistItemListRecovered Then
            For Each rows As DataGridViewRow In DataGridView3.Rows
                If Util.CheckDGVCellValue(rows.Cells(2).Value) Then
                    rows.Cells(0).Value = "https://www.youtube.com/watch?v=" & rows.Cells(2).Value
                End If
            Next
        End If

        If data = -1 OrElse data = DataTables.playlistItemListRemoved Then
            For Each rows As DataGridViewRow In DataGridView4.Rows
                If Util.CheckDGVCellValue(rows.Cells(2).Value) Then
                    rows.Cells(0).Value = "https://www.youtube.com/watch?v=" & rows.Cells(2).Value
                End If
            Next
        End If

        If data = -1 OrElse data = DataTables.playlistItemListLost Then
            For Each rows As DataGridViewRow In DataGridView5.Rows
                If Util.CheckDGVCellValue(rows.Cells(2).Value) Then
                    rows.Cells(0).Value = "https://www.youtube.com/watch?v=" & rows.Cells(2).Value
                End If
            Next
        End If
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged

        If TabControl1.SelectedIndex = 0 Then
            SetRowValue(DataTables.playlistList)
            DataGridView1.Refresh()
        ElseIf TabControl1.SelectedIndex = 1 Then
            SetRowValue(DataTables.playlistItemList)
            SetComboboxValues(DataTables.playlistItemList)
            DataGridView2.Refresh()
        ElseIf TabControl1.SelectedIndex = 2 Then
            SetRowValue(DataTables.playlistItemListRecovered)
            SetComboboxValues(DataTables.playlistItemListRecovered)
            DataGridView3.Refresh()
        ElseIf TabControl1.SelectedIndex = 3 Then
            SetRowValue(DataTables.playlistItemListRemoved)
            SetComboboxValues(DataTables.playlistItemListRemoved)
            DataGridView4.Refresh()
        ElseIf TabControl1.SelectedIndex = 4 Then
            SetRowValue(DataTables.playlistItemListLost)
            SetComboboxValues(DataTables.playlistItemListLost)
            DataGridView5.Refresh()
        ElseIf TabControl1.SelectedIndex = 5 Then
            DataGridView6.Refresh()
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
            Util.InitCombobox(ComboBox1)
            ComboBox1.DataSource = playlistList.ToList
        End If

        If data = -1 OrElse data = DataTables.playlistItemListRecovered Then
            Util.InitCombobox(ComboBox2)
            ComboBox2.DataSource = playlistList.ToList
        End If

        If data = -1 OrElse data = DataTables.playlistItemListRemoved Then
            Util.InitCombobox(ComboBox3)
            ComboBox3.DataSource = playlistList.ToList
        End If

        If data = -1 OrElse data = DataTables.playlistItemListLost Then
            Util.InitCombobox(ComboBox4)
            ComboBox4.DataSource = playlistList.ToList
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim row = DataGridView1.Rows(e.RowIndex)
            If row.Cells(0).Value Is Nothing Then Return
            Dim url = row.Cells(0).Value.ToString()
            Util.OpenLink(url)
        End If
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim row = DataGridView2.Rows(e.RowIndex)
            If row.Cells(0).Value Is Nothing Then Return
            Dim url = row.Cells(0).Value.ToString()
            Util.OpenLink(url)
        End If
    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim row = DataGridView3.Rows(e.RowIndex)
            If row.Cells(0).Value Is Nothing Then Return
            Dim url = row.Cells(0).Value.ToString()
            Util.OpenLink(url)
        End If
    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim row = DataGridView4.Rows(e.RowIndex)
            If row.Cells(0).Value Is Nothing Then Return
            Dim url = row.Cells(0).Value.ToString()
            Util.OpenLink(url)
        End If
    End Sub

    Private Sub DataGridView5_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView5.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim row = DataGridView5.Rows(e.RowIndex)
            If row.Cells(0).Value Is Nothing Then Return
            Dim url = row.Cells(0).Value.ToString()
            Util.OpenLink(url)
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim filter As String = ""
        If Not ComboBox1.SelectedItem.Value = "All" Then
            filter = String.Format("playlistID = '{0}'", ComboBox1.SelectedItem.Key)
        End If
        Util.SetFilterDataGridViewData(Database.playlistItemListData, DataGridView2, filter)
        SetRowValue(DataTables.playlistItemList)
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim filter As String = ""
        If Not ComboBox2.SelectedItem.Value = "All" Then
            filter = String.Format("playlistID = '{0}'", ComboBox2.SelectedItem.Key)
        End If
        Util.SetFilterDataGridViewData(Database.playlistItemListRecoveredData, DataGridView3, filter)
        SetRowValue(DataTables.playlistItemListRecovered)
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Dim filter As String = ""
        If Not ComboBox3.SelectedItem.Value = "All" Then
            filter = String.Format("playlistID = '{0}'", ComboBox3.SelectedItem.Key)
        End If
        Util.SetFilterDataGridViewData(Database.playlistItemListRemovedData, DataGridView4, filter)
        SetRowValue(DataTables.playlistItemListRemoved)
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Dim filter As String = ""
        If Not ComboBox4.SelectedItem.Value = "All" Then
            filter = String.Format("playlistID = '{0}'", ComboBox4.SelectedItem.Key)
        End If
        Util.SetFilterDataGridViewData(Database.playlistItemListLostData, DataGridView5, filter)
        SetRowValue(DataTables.playlistItemListLost)
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
#End Region
End Class
