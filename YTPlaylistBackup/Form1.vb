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
        SetDGVData()

        SetComboboxValues()
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
            SetDGVData()

            MsgBox("Sync Completed")
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
                        Else
                            'check if not in recovered and not in lost
                            Dim filterRowRecover = Database.playlistItemListRecoveredData.AsEnumerable.Where(Function(dr) dr(0).ToString = playlistId AndAlso dr(1).ToString = videoId).FirstOrDefault
                            Dim filterRowLost = Database.playlistItemListLostData.AsEnumerable.Where(Function(dr) dr(0).ToString = playlistId AndAlso dr(1).ToString = videoId).FirstOrDefault
                            If filterRowRecover Is Nothing AndAlso filterRowLost Is Nothing Then
                                'no backup
                                Database.InsertSqlPlaylistItemLostList(playlistId, videoId, title, desc, videoOwnerChannelId, videoOwnerChannelTitle)
                            End If
                        End If
                    Else
                        If filterRow Is Nothing Then
                            'new added
                            Database.InsertSqlPlaylistItemList(playlistId, videoId, title, desc, videoOwnerChannelId, videoOwnerChannelTitle)
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
                    End If
                Next
            End If
        Next

        Database.GetSqlPlaylistItemsList()
        Database.GetSqlPlaylistItemsRecoveredList()
    End Sub
#End Region

#Region "Sub/Func"
    Private Sub SetDGVData()
        Util.InitDGV(DataGridView1)
        Util.InitDGV(DataGridView2)
        Util.InitDGV(DataGridView3)
        Util.InitDGV(DataGridView4)
        Util.InitDGV(DataGridView5)

        Util.SetFilterDataGridViewData(Database.playlistListData, DataGridView1)
        Util.SetFilterDataGridViewData(Database.playlistItemListData, DataGridView2)
        Util.SetFilterDataGridViewData(Database.playlistItemListRecoveredData, DataGridView3)
        Util.SetFilterDataGridViewData(Database.playlistItemListRemovedData, DataGridView4)
        Util.SetFilterDataGridViewData(Database.playlistItemListLostData, DataGridView5)

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

        SetYTLinks()

        DataGridView1.Refresh()
        DataGridView2.Refresh()
        DataGridView3.Refresh()
        DataGridView4.Refresh()
        DataGridView5.Refresh()
    End Sub

    Private Sub SetYTLinks()
        For Each rows As DataGridViewRow In DataGridView1.Rows
            If Util.CheckDGVCellValue(rows.Cells(1).Value) Then
                rows.Cells(0).Value = "https://www.youtube.com/playlist?list=" & rows.Cells(1).Value
            End If
        Next

        For Each rows As DataGridViewRow In DataGridView2.Rows
            If Util.CheckDGVCellValue(rows.Cells(2).Value) Then
                rows.Cells(0).Value = "https://www.youtube.com/watch?v=" & rows.Cells(2).Value
            End If
        Next

        For Each rows As DataGridViewRow In DataGridView3.Rows
            If Util.CheckDGVCellValue(rows.Cells(2).Value) Then
                rows.Cells(0).Value = "https://www.youtube.com/watch?v=" & rows.Cells(2).Value
            End If
        Next

        For Each rows As DataGridViewRow In DataGridView4.Rows
            If Util.CheckDGVCellValue(rows.Cells(2).Value) Then
                rows.Cells(0).Value = "https://www.youtube.com/watch?v=" & rows.Cells(2).Value
            End If
        Next

        For Each rows As DataGridViewRow In DataGridView5.Rows
            If Util.CheckDGVCellValue(rows.Cells(2).Value) Then
                rows.Cells(0).Value = "https://www.youtube.com/watch?v=" & rows.Cells(2).Value
            End If
        Next
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        SetYTLinks()
        SetComboboxValues()

        DataGridView1.Refresh()
        DataGridView2.Refresh()
        DataGridView3.Refresh()
        DataGridView4.Refresh()
        DataGridView5.Refresh()
    End Sub

    Private Sub SetComboboxValues()
        Util.InitCombobox(ComboBox1)
        Util.InitCombobox(ComboBox2)
        Util.InitCombobox(ComboBox3)
        Util.InitCombobox(ComboBox4)

        Dim playlistList As New Dictionary(Of String, String) From {
            {"All Playlist", "All"}
        }

        For Each rows In Database.playlistListData.Rows
            If Not playlistList.ContainsKey(rows(0)) Then
                playlistList.Add(rows(0), rows(1))
            End If
        Next

        ComboBox1.DataSource = playlistList.ToList
        ComboBox2.DataSource = playlistList.ToList
        ComboBox3.DataSource = playlistList.ToList
        ComboBox4.DataSource = playlistList.ToList
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim row = DataGridView1.Rows(e.RowIndex)
            If row.Cells(0).Value Is Nothing Then Return
            Dim url = row.Cells(0).Value.ToString()
            Process.Start(Util.GetChromePath, url)
        End If
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim row = DataGridView2.Rows(e.RowIndex)
            If row.Cells(0).Value Is Nothing Then Return
            Dim url = row.Cells(0).Value.ToString()
            Process.Start(Util.GetChromePath, url)
        End If
    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim row = DataGridView3.Rows(e.RowIndex)
            If row.Cells(0).Value Is Nothing Then Return
            Dim url = row.Cells(0).Value.ToString()
            Process.Start(Util.GetChromePath, url)
        End If
    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim row = DataGridView4.Rows(e.RowIndex)
            If row.Cells(0).Value Is Nothing Then Return
            Dim url = row.Cells(0).Value.ToString()
            Process.Start(Util.GetChromePath, url)
        End If
    End Sub

    Private Sub DataGridView5_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView5.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim row = DataGridView5.Rows(e.RowIndex)
            If row.Cells(0).Value Is Nothing Then Return
            Dim url = row.Cells(0).Value.ToString()
            Process.Start(Util.GetChromePath, url)
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim filter As String = ""
        If Not ComboBox1.SelectedItem.Value = "All" Then
            filter = String.Format("playlistID = '{0}'", ComboBox1.SelectedItem.Key)
        End If
        Util.SetFilterDataGridViewData(Database.playlistItemListData, DataGridView2, filter)
        SetYTLinks()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim filter As String = ""
        If Not ComboBox2.SelectedItem.Value = "All" Then
            filter = String.Format("playlistID = '{0}'", ComboBox2.SelectedItem.Key)
        End If
        Util.SetFilterDataGridViewData(Database.playlistItemListRecoveredData, DataGridView3, filter)
        SetYTLinks()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Dim filter As String = ""
        If Not ComboBox3.SelectedItem.Value = "All" Then
            filter = String.Format("playlistID = '{0}'", ComboBox3.SelectedItem.Key)
        End If
        Util.SetFilterDataGridViewData(Database.playlistItemListRemovedData, DataGridView4, filter)
        SetYTLinks()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Dim filter As String = ""
        If Not ComboBox4.SelectedItem.Value = "All" Then
            filter = String.Format("playlistID = '{0}'", ComboBox4.SelectedItem.Key)
        End If
        Util.SetFilterDataGridViewData(Database.playlistItemListLostData, DataGridView5, filter)
        SetYTLinks()
    End Sub
#End Region
End Class
