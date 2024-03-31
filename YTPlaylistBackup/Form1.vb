Imports System.Threading
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Services
Imports Google.Apis.Util.Store
Imports Google.Apis.YouTube.v3
Imports Google.Apis.YouTube.v3.Data
Imports Microsoft.Data.SqlClient

Public Class Form1
    Private credential As UserCredential
    Private ytService As YouTubeService
    Private objconnection As New SqlConnection(Util.SqlConnection)

    Dim playlistListData As New DataTable
    Dim playlistItemListData As New DataTable
    Dim playlistItemListRecoveredData As New DataTable

    'reference https://stackoverflow.com/questions/65357223/get-youtube-channel-data-using-google-youtube-data-api-in-vb-net

    Private Async Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()

        InitDatatables()

        GetSqlPlaylistList()
        GetSqlPlaylistItemsList()
        GetSqlPlaylistItemsRecoveredList()

        Await OAuth()

        ToolStripButton1.Visible = False
        ToolStripButton2.Visible = False
    End Sub

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

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        GetPlaylistList()
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If MessageBox.Show("Are you sure?", "Expensive API Call", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            GetPlaylistItemLists()
        End If
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

                Dim filterRow = playlistListData.AsEnumerable.Where(Function(dr) dr(0).ToString = id).FirstOrDefault
                If filterRow Is Nothing Then
                    InsertSqlPlaylistList(id, title, desc, itemCount)
                Else
                    If filterRow(1) <> title OrElse filterRow(2) <> desc OrElse filterRow(3) <> itemCount Then
                        UpdateSqlPlaylist(id, title, desc, itemCount)
                    End If
                End If

                If Not playlistCache.ContainsKey(id) Then
                    playlistCache.Add(id, title)
                End If
            Next

            nextPageToken = playlistListResp.NextPageToken
        End While

        'removed
        For Each rows In playlistListData.Rows
            If Not playlistCache.ContainsKey(rows(0)) Then
                DeleteSqlPlaylist(rows(0))
            End If
        Next

        GetSqlPlaylistList()
    End Sub

    Private Sub GetPlaylistItemLists()
        Dim playlistItemsCache As New Dictionary(Of String, String)
        For Each playlist In playlistListData.Rows
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

                    Dim filterRow = playlistItemListData.AsEnumerable.Where(Function(dr) dr(0).ToString = playlistId AndAlso dr(1).ToString = videoId).FirstOrDefault

                    If videoId <> "" AndAlso videoOwnerChannelId Is Nothing AndAlso
                        videoOwnerChannelTitle Is Nothing AndAlso
                        (status = "privacyStatusUnspecified" OrElse status = "private" OrElse
                        title = "Deleted video" OrElse title = "Private video" OrElse
                        desc = "This video is unavailable." OrElse desc = "This video is private.") Then
                        If filterRow IsNot Nothing Then
                            'recovered
                            InsertSqlPlaylistItemRecoveredList(playlistId, filterRow(1), filterRow(2), filterRow(3), filterRow(4), filterRow(5))
                            DeleteSqlPlaylistItem(filterRow(1), playlistId)
                        Else
                            'check if not in recovered
                            Dim filterTableRecover = playlistItemListRecoveredData.AsEnumerable.Where(Function(dr) dr(0).ToString = playlistId AndAlso dr(1).ToString = videoId).FirstOrDefault
                            If filterTableRecover Is Nothing Then
                                'no backup
                                InsertSqlPlaylistItemLostList(playlistId, videoId, title, desc, videoOwnerChannelId, videoOwnerChannelTitle)
                            End If
                        End If
                    Else
                        If filterRow Is Nothing Then
                            'new added
                            InsertSqlPlaylistItemList(playlistId, videoId, title, desc, videoOwnerChannelId, videoOwnerChannelTitle)
                        End If
                    End If

                    If Not playlistItemsCache.ContainsKey(videoId) Then
                        playlistItemsCache.Add(videoId, title)
                    End If
                Next

                nextPageToken = playlistItemListResp.NextPageToken
            End While

            'removed
            Dim filteredTable = playlistItemListData.AsEnumerable.Where(Function(dr) dr(0).ToString = playlistId).ToList
            If filteredTable IsNot Nothing Then
                For Each rows In filteredTable
                    If Not playlistItemsCache.ContainsKey(rows(1)) Then
                        InsertSqlPlaylistItemRemovedList(playlistId, rows(1), rows(2), rows(3), rows(4), rows(5))
                        DeleteSqlPlaylistItem(rows(1), playlistId)
                    End If
                Next
            End If
        Next

        GetSqlPlaylistItemsList()
        GetSqlPlaylistItemsRecoveredList()
    End Sub

    Private Sub InitDatatables()
        If playlistListData.Columns.Count = 0 Then
            playlistListData.Columns.Add("playlistID", GetType(String))
            playlistListData.Columns.Add("title", GetType(String))
            playlistListData.Columns.Add("description", GetType(String))
            playlistListData.Columns.Add("itemCount", GetType(Integer))
        End If

        If playlistItemListData.Columns.Count = 0 Then
            playlistItemListData.Columns.Add("playlistID", GetType(String))
            playlistItemListData.Columns.Add("videoID", GetType(String))
            playlistItemListData.Columns.Add("title", GetType(String))
            playlistItemListData.Columns.Add("description", GetType(String))
            playlistItemListData.Columns.Add("videoOwnerChannelId", GetType(String))
            playlistItemListData.Columns.Add("videoOwnerChannelTitle", GetType(String))
        End If

        If playlistItemListRecoveredData.Columns.Count = 0 Then
            playlistItemListRecoveredData.Columns.Add("playlistID", GetType(String))
            playlistItemListRecoveredData.Columns.Add("videoID", GetType(String))
            playlistItemListRecoveredData.Columns.Add("title", GetType(String))
            playlistItemListRecoveredData.Columns.Add("description", GetType(String))
            playlistItemListRecoveredData.Columns.Add("videoOwnerChannelId", GetType(String))
            playlistItemListRecoveredData.Columns.Add("videoOwnerChannelTitle", GetType(String))
        End If
    End Sub

    Private Sub GetSqlPlaylistList()
        Dim c As New SqlCommand
        Dim dr As SqlDataReader
        Try
            If objconnection.State = ConnectionState.Closed Then
                objconnection.Open()
            End If
            With c
                .Connection = objconnection
                .CommandText = "SELECT * from Playlists ORDER BY id"
                .CommandType = CommandType.Text
                dr = .ExecuteReader
            End With
            If playlistListData.Rows.Count > 0 AndAlso dr.HasRows Then
                playlistListData.Rows.Clear()
                playlistListData.AcceptChanges()
            End If
            While dr.Read
                playlistListData.Rows.Add(dr(1), dr(2), dr(3), dr(4))
            End While
            playlistListData.AcceptChanges()
            dr.Close()
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub

    Private Sub GetSqlPlaylistItemsList()
        Dim c As New SqlCommand
        Dim dr As SqlDataReader
        Try
            If objconnection.State = ConnectionState.Closed Then
                objconnection.Open()
            End If
            With c
                .Connection = objconnection
                .CommandText = "SELECT * from PlaylistItems ORDER BY id"
                .CommandType = CommandType.Text
                dr = .ExecuteReader
            End With
            If playlistItemListData.Rows.Count > 0 AndAlso dr.HasRows Then
                playlistItemListData.Rows.Clear()
                playlistItemListData.AcceptChanges()
            End If
            While dr.Read
                playlistItemListData.Rows.Add(dr(1), dr(2), dr(3), dr(4), dr(5), dr(6))
            End While
            playlistItemListData.AcceptChanges()
            dr.Close()
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub

    Private Sub GetSqlPlaylistItemsRecoveredList()
        Dim c As New SqlCommand
        Dim dr As SqlDataReader
        Try
            If objconnection.State = ConnectionState.Closed Then
                objconnection.Open()
            End If
            With c
                .Connection = objconnection
                .CommandText = "SELECT * from PlaylistItemsRecovered ORDER BY id"
                .CommandType = CommandType.Text
                dr = .ExecuteReader
            End With
            If playlistItemListRecoveredData.Rows.Count > 0 AndAlso dr.HasRows Then
                playlistItemListRecoveredData.Rows.Clear()
                playlistItemListRecoveredData.AcceptChanges()
            End If
            While dr.Read
                playlistItemListRecoveredData.Rows.Add(dr(1), dr(2), dr(3), dr(4), dr(5), dr(6))
            End While
            playlistItemListRecoveredData.AcceptChanges()
            dr.Close()
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub

    Private Sub InsertSqlPlaylistList(playlistId As String, title As String, description As String, itemCount As Integer)
        Dim c As New SqlCommand
        Try
            If objconnection.State = ConnectionState.Closed Then
                objconnection.Open()
            End If
            With c
                .Connection = objconnection
                .CommandText = "INSERT INTO Playlists (playlistID, title, description, itemCount, syncDate)" &
                    "Values(@playlistID, @title, @description, @itemCount, @syncDate)"
                .CommandType = CommandType.Text
                .Parameters.AddWithValue("@playlistID", playlistId)
                .Parameters.AddWithValue("@title", title)
                .Parameters.AddWithValue("@description", description)
                .Parameters.AddWithValue("@itemCount", itemCount)
                .Parameters.AddWithValue("@syncDate", Now)
                .ExecuteNonQuery()
            End With
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub

    Private Sub InsertSqlPlaylistItemList(playlistID As String, videoID As String, title As String, description As String, videoOwnerChannelId As String, videoOwnerChannelTitle As String)
        Dim c As New SqlCommand
        Try
            If objconnection.State = ConnectionState.Closed Then
                objconnection.Open()
            End If
            With c
                .Connection = objconnection
                .CommandText = "INSERT INTO PlaylistItems (playlistID, videoID, title, description, videoOwnerChannelId, videoOwnerChannelTitle, syncDate)" &
                    "Values(@playlistID, @videoID, @title, @description, @videoOwnerChannelId, @videoOwnerChannelTitle, @syncDate)"
                .CommandType = CommandType.Text
                .Parameters.AddWithValue("@playlistID", playlistID)
                .Parameters.AddWithValue("@videoID", videoID)
                .Parameters.AddWithValue("@title", title)
                .Parameters.AddWithValue("@description", description)
                .Parameters.AddWithValue("@videoOwnerChannelId", videoOwnerChannelId)
                .Parameters.AddWithValue("@videoOwnerChannelTitle", videoOwnerChannelTitle)
                .Parameters.AddWithValue("@syncDate", Now)
                .ExecuteNonQuery()
            End With
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub

    Private Sub InsertSqlPlaylistItemRecoveredList(playlistID As String, videoID As String, title As String, description As String, videoOwnerChannelId As String, videoOwnerChannelTitle As String)
        Dim c As New SqlCommand
        Try
            If objconnection.State = ConnectionState.Closed Then
                objconnection.Open()
            End If
            With c
                .Connection = objconnection
                .CommandText = "INSERT INTO PlaylistItemsRecovered (playlistID, videoID, title, description, videoOwnerChannelId, videoOwnerChannelTitle, syncDate)" &
                    "Values(@playlistID, @videoID, @title, @description, @videoOwnerChannelId, @videoOwnerChannelTitle, @syncDate)"
                .CommandType = CommandType.Text
                .Parameters.AddWithValue("@playlistID", playlistID)
                .Parameters.AddWithValue("@videoID", videoID)
                .Parameters.AddWithValue("@title", title)
                .Parameters.AddWithValue("@description", description)
                .Parameters.AddWithValue("@videoOwnerChannelId", videoOwnerChannelId)
                .Parameters.AddWithValue("@videoOwnerChannelTitle", videoOwnerChannelTitle)
                .Parameters.AddWithValue("@syncDate", Now)
                .ExecuteNonQuery()
            End With
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub

    Private Sub InsertSqlPlaylistItemRemovedList(playlistID As String, videoID As String, title As String, description As String, videoOwnerChannelId As String, videoOwnerChannelTitle As String)
        Dim c As New SqlCommand
        Try
            If objconnection.State = ConnectionState.Closed Then
                objconnection.Open()
            End If
            With c
                .Connection = objconnection
                .CommandText = "INSERT INTO PlaylistItemsRemoved (playlistID, videoID, title, description, videoOwnerChannelId, videoOwnerChannelTitle, syncDate)" &
                    "Values(@playlistID, @videoID, @title, @description, @videoOwnerChannelId, @videoOwnerChannelTitle, @syncDate)"
                .CommandType = CommandType.Text
                .Parameters.AddWithValue("@playlistID", playlistID)
                .Parameters.AddWithValue("@videoID", videoID)
                .Parameters.AddWithValue("@title", title)
                .Parameters.AddWithValue("@description", description)
                .Parameters.AddWithValue("@videoOwnerChannelId", videoOwnerChannelId)
                .Parameters.AddWithValue("@videoOwnerChannelTitle", videoOwnerChannelTitle)
                .Parameters.AddWithValue("@syncDate", Now)
                .ExecuteNonQuery()
            End With
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub

    Private Sub InsertSqlPlaylistItemLostList(playlistID As String, videoID As String, title As String, description As String, videoOwnerChannelId As String, videoOwnerChannelTitle As String)
        Dim c As New SqlCommand
        Try
            If objconnection.State = ConnectionState.Closed Then
                objconnection.Open()
            End If
            With c
                .Connection = objconnection
                .CommandText = "INSERT INTO PlaylistItemsLost (playlistID, videoID, title, description, videoOwnerChannelId, videoOwnerChannelTitle, syncDate)" &
                    "Values(@playlistID, @videoID, @title, @description, @videoOwnerChannelId, @videoOwnerChannelTitle, @syncDate)"
                .CommandType = CommandType.Text
                .Parameters.AddWithValue("@playlistID", playlistID)
                .Parameters.AddWithValue("@videoID", videoID)
                .Parameters.AddWithValue("@title", title)
                .Parameters.AddWithValue("@description", description)
                .Parameters.AddWithValue("@videoOwnerChannelId", If(videoOwnerChannelId, ""))
                .Parameters.AddWithValue("@videoOwnerChannelTitle", If(videoOwnerChannelTitle, ""))
                .Parameters.AddWithValue("@syncDate", Now)
                .ExecuteNonQuery()
            End With
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub

    Private Sub DeleteSqlPlaylistItem(videoID As String, playlistID As String)
        Dim c As New SqlCommand
        Try
            If objconnection.State = ConnectionState.Closed Then
                objconnection.Open()
            End If
            With c
                .Connection = objconnection
                .CommandText = "DELETE PlaylistItems WHERE videoID=@videoID and playlistID=@playlistID"
                .CommandType = CommandType.Text
                .Parameters.AddWithValue("@videoID", videoID)
                .Parameters.AddWithValue("@playlistID", playlistID)
                .ExecuteNonQuery()
            End With
        Catch ex As Exception
        End Try
        objconnection.Close()

        ReseedPlaylistItemsID()
    End Sub

    Private Sub DeleteSqlPlaylist(playlistID As String)
        Dim c As New SqlCommand
        Try
            If objconnection.State = ConnectionState.Closed Then
                objconnection.Open()
            End If
            With c
                .Connection = objconnection
                .CommandText = "DELETE Playlists WHERE playlistID=@playlistID"
                .CommandType = CommandType.Text
                .Parameters.AddWithValue("@playlistID", playlistID)
                .ExecuteNonQuery()
            End With
        Catch ex As Exception
        End Try
        objconnection.Close()

        ReseedPlaylistID()
    End Sub

    Private Sub UpdateSqlPlaylist(playlistID As String, title As String, description As String, itemCount As String)
        Dim c As New SqlCommand
        Try
            If objconnection.State = ConnectionState.Closed Then
                objconnection.Open()
            End If
            With c
                .Connection = objconnection
                .CommandText = "UPDATE Playlists SET title=@title, description=@description, itemCount=@itemCount, syncDate=@syncDate WHERE playlistID=@playlistID"
                .CommandType = CommandType.Text
                .Parameters.AddWithValue("@playlistID", playlistID)
                .Parameters.AddWithValue("@title", title)
                .Parameters.AddWithValue("@description", description)
                .Parameters.AddWithValue("@itemCount", itemCount)
                .Parameters.AddWithValue("@syncDate", Now)
                .ExecuteNonQuery()
            End With
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub

    Private Sub ReseedPlaylistItemsID()
        Dim c As New SqlCommand
        Try
            If objconnection.State = ConnectionState.Closed Then
                objconnection.Open()
            End If
            With c
                .Connection = objconnection
                .CommandText = "declare @max int select @max=ISNULL(max(id),0) from PlaylistItems; DBCC CHECKIDENT ('PlaylistItems', RESEED, @max );"
                .CommandType = CommandType.Text
                .ExecuteNonQuery()
            End With
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub

    Private Sub ReseedPlaylistID()
        Dim c As New SqlCommand
        Try
            If objconnection.State = ConnectionState.Closed Then
                objconnection.Open()
            End If
            With c
                .Connection = objconnection
                .CommandText = "declare @max int select @max=ISNULL(max(id),0) from Playlists; DBCC CHECKIDENT ('Playlists', RESEED, @max );"
                .CommandType = CommandType.Text
                .ExecuteNonQuery()
            End With
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub
End Class
