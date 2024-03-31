Imports System.Threading
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Services
Imports Google.Apis.Util.Store
Imports Google.Apis.YouTube.v3
Imports Google.Apis.YouTube.v3.Data
Imports Microsoft.Data.SqlClient

Public Class Form1
    'reference https://stackoverflow.com/questions/65357223/get-youtube-channel-data-using-google-youtube-data-api-in-vb-net

#Region "Variables"
    Private credential As UserCredential
    Private ytService As YouTubeService
#End Region

#Region "Form Events"
    Private Async Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()

        Database.InitDatatables()

        Database.GetSqlPlaylistList()
        Database.GetSqlPlaylistItemsList()
        Database.GetSqlPlaylistItemsRecoveredList()

        Await OAuth()
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        If MessageBox.Show("Are you sure?", "Expensive API Call", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            ToolStripButton3.Enabled = False
            GetPlaylistList()
            GetPlaylistItemLists()
            MsgBox("Sync Completed")
            ToolStripButton3.Enabled = True
        End If
    End Sub
#End Region

#Region "Sub/Func"
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
                            'check if not in recovered
                            Dim filterTableRecover = Database.playlistItemListRecoveredData.AsEnumerable.Where(Function(dr) dr(0).ToString = playlistId AndAlso dr(1).ToString = videoId).FirstOrDefault
                            If filterTableRecover Is Nothing Then
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
End Class
