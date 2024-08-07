Imports Microsoft.Data.SqlClient

Public Class Database
#Region "Datatable"
    Public Shared Property playlistListData As New DataTable
    Public Shared Property playlistItemListData As New DataTable
    Public Shared Property playlistItemListRecoveredData As New DataTable
    Public Shared Property playlistItemListRemovedData As New DataTable
    Public Shared Property playlistItemListLostData As New DataTable
    Public Shared Property syncHistoryData As New DataTable
#End Region

#Region "Variables"
    Private Shared Property objconnection As New SqlConnection(SqlConn)
#End Region

#Region "Sub/Func"
    Public Shared Sub InitDatatables()
        playlistListData.Clear()
        playlistListData.AcceptChanges()
        If playlistListData.Columns.Count = 0 Then
            playlistListData.Columns.Add("playlistID", GetType(String))
            playlistListData.Columns.Add("title", GetType(String))
            playlistListData.Columns.Add("description", GetType(String))
            playlistListData.Columns.Add("itemCount", GetType(Integer))
            playlistListData.Columns.Add("syncDate", GetType(Date))
        End If

        playlistItemListData.Clear()
        playlistItemListData.AcceptChanges()
        If playlistItemListData.Columns.Count = 0 Then
            playlistItemListData.Columns.Add("playlistID", GetType(String))
            playlistItemListData.Columns.Add("videoID", GetType(String))
            playlistItemListData.Columns.Add("title", GetType(String))
            playlistItemListData.Columns.Add("description", GetType(String))
            playlistItemListData.Columns.Add("videoOwnerChannelId", GetType(String))
            playlistItemListData.Columns.Add("videoOwnerChannelTitle", GetType(String))
            playlistItemListData.Columns.Add("syncDate", GetType(Date))
        End If

        playlistItemListRecoveredData.Clear()
        playlistItemListRecoveredData.AcceptChanges()
        If playlistItemListRecoveredData.Columns.Count = 0 Then
            playlistItemListRecoveredData.Columns.Add("playlistID", GetType(String))
            playlistItemListRecoveredData.Columns.Add("videoID", GetType(String))
            playlistItemListRecoveredData.Columns.Add("title", GetType(String))
            playlistItemListRecoveredData.Columns.Add("description", GetType(String))
            playlistItemListRecoveredData.Columns.Add("videoOwnerChannelId", GetType(String))
            playlistItemListRecoveredData.Columns.Add("videoOwnerChannelTitle", GetType(String))
            playlistItemListRecoveredData.Columns.Add("syncDate", GetType(Date))
        End If

        playlistItemListRemovedData.Clear()
        playlistItemListRemovedData.AcceptChanges()
        If playlistItemListRemovedData.Columns.Count = 0 Then
            playlistItemListRemovedData.Columns.Add("playlistID", GetType(String))
            playlistItemListRemovedData.Columns.Add("videoID", GetType(String))
            playlistItemListRemovedData.Columns.Add("title", GetType(String))
            playlistItemListRemovedData.Columns.Add("description", GetType(String))
            playlistItemListRemovedData.Columns.Add("videoOwnerChannelId", GetType(String))
            playlistItemListRemovedData.Columns.Add("videoOwnerChannelTitle", GetType(String))
            playlistItemListRemovedData.Columns.Add("syncDate", GetType(Date))
        End If

        playlistItemListLostData.Clear()
        playlistItemListLostData.AcceptChanges()
        If playlistItemListLostData.Columns.Count = 0 Then
            playlistItemListLostData.Columns.Add("playlistID", GetType(String))
            playlistItemListLostData.Columns.Add("videoID", GetType(String))
            playlistItemListLostData.Columns.Add("title", GetType(String))
            playlistItemListLostData.Columns.Add("description", GetType(String))
            playlistItemListLostData.Columns.Add("videoOwnerChannelId", GetType(String))
            playlistItemListLostData.Columns.Add("videoOwnerChannelTitle", GetType(String))
            playlistItemListLostData.Columns.Add("syncDate", GetType(Date))
        End If

        syncHistoryData.Clear()
        syncHistoryData.AcceptChanges()
        If syncHistoryData.Columns.Count = 0 Then
            syncHistoryData.Columns.Add("syncDate", GetType(Date))
            syncHistoryData.Columns.Add("AddedCount", GetType(Integer))
            syncHistoryData.Columns.Add("RemovedCount", GetType(Integer))
            syncHistoryData.Columns.Add("RecoveredCount", GetType(Integer))
            syncHistoryData.Columns.Add("LostCount", GetType(Integer))
            syncHistoryData.Columns.Add("Notes", GetType(String))
        End If
    End Sub
#End Region

#Region "Read"
    Public Shared Sub GetSqlPlaylistList()
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
                playlistListData.Rows.Add(dr(1), dr(2), dr(3), dr(4), dr(5))
            End While
            playlistListData.AcceptChanges()
            dr.Close()
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub

    Public Shared Sub GetSqlPlaylistItemsList()
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
                playlistItemListData.Rows.Add(dr(1), dr(2), dr(3), dr(4), dr(5), dr(6), dr(7))
            End While
            playlistItemListData.AcceptChanges()
            dr.Close()
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub

    Public Shared Sub GetSqlPlaylistItemsRecoveredList()
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
                playlistItemListRecoveredData.Rows.Add(dr(1), dr(2), dr(3), dr(4), dr(5), dr(6), dr(7))
            End While
            playlistItemListRecoveredData.AcceptChanges()
            dr.Close()
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub

    Public Shared Sub GetSqlPlaylistItemsRemovedList()
        Dim c As New SqlCommand
        Dim dr As SqlDataReader
        Try
            If objconnection.State = ConnectionState.Closed Then
                objconnection.Open()
            End If
            With c
                .Connection = objconnection
                .CommandText = "SELECT * from PlaylistItemsRemoved ORDER BY id"
                .CommandType = CommandType.Text
                dr = .ExecuteReader
            End With
            If playlistItemListRemovedData.Rows.Count > 0 AndAlso dr.HasRows Then
                playlistItemListRemovedData.Rows.Clear()
                playlistItemListRemovedData.AcceptChanges()
            End If
            While dr.Read
                playlistItemListRemovedData.Rows.Add(dr(1), dr(2), dr(3), dr(4), dr(5), dr(6), dr(7))
            End While
            playlistItemListRemovedData.AcceptChanges()
            dr.Close()
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub

    Public Shared Sub GetSqlPlaylistItemsLostList()
        Dim c As New SqlCommand
        Dim dr As SqlDataReader
        Try
            If objconnection.State = ConnectionState.Closed Then
                objconnection.Open()
            End If
            With c
                .Connection = objconnection
                .CommandText = "SELECT * from PlaylistItemsLost ORDER BY id"
                .CommandType = CommandType.Text
                dr = .ExecuteReader
            End With
            If playlistItemListLostData.Rows.Count > 0 AndAlso dr.HasRows Then
                playlistItemListLostData.Rows.Clear()
                playlistItemListLostData.AcceptChanges()
            End If
            While dr.Read
                playlistItemListLostData.Rows.Add(dr(1), dr(2), dr(3), dr(4), dr(5), dr(6), dr(7))
            End While
            playlistItemListLostData.AcceptChanges()
            dr.Close()
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub

    Public Shared Sub GetSqlSyncHistory()
        Dim c As New SqlCommand
        Dim dr As SqlDataReader
        Try
            If objconnection.State = ConnectionState.Closed Then
                objconnection.Open()
            End If
            With c
                .Connection = objconnection
                .CommandText = "SELECT * from SyncHistory ORDER BY syncDate ASC"
                .CommandType = CommandType.Text
                dr = .ExecuteReader
            End With
            If syncHistoryData.Rows.Count > 0 AndAlso dr.HasRows Then
                syncHistoryData.Rows.Clear()
                syncHistoryData.AcceptChanges()
            End If
            While dr.Read
                syncHistoryData.Rows.Add(dr(6), dr(1), dr(2), dr(3), dr(4), dr(5))
            End While
            syncHistoryData.AcceptChanges()
            dr.Close()
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub

    Public Shared Sub GetSqlAll()
        InitDatatables()

        GetSqlPlaylistList()
        GetSqlPlaylistItemsList()
        GetSqlPlaylistItemsRecoveredList()
        GetSqlPlaylistItemsRemovedList()
        GetSqlPlaylistItemsLostList()
        GetSqlSyncHistory()
    End Sub
#End Region

#Region "Insert"
    Public Shared Sub InsertSqlPlaylistList(playlistId As String, title As String, description As String, itemCount As Integer)
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

    Public Shared Sub InsertSqlPlaylistItemList(playlistID As String, videoID As String, title As String, description As String, videoOwnerChannelId As String, videoOwnerChannelTitle As String)
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

    Public Shared Sub InsertSqlPlaylistItemRecoveredList(playlistID As String, videoID As String, title As String, description As String, videoOwnerChannelId As String, videoOwnerChannelTitle As String)
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

    Public Shared Sub InsertSqlPlaylistItemRemovedList(playlistID As String, videoID As String, title As String, description As String, videoOwnerChannelId As String, videoOwnerChannelTitle As String)
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

    Public Shared Sub InsertSqlPlaylistItemLostList(playlistID As String, videoID As String, title As String, description As String, videoOwnerChannelId As String, videoOwnerChannelTitle As String)
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

    Public Shared Sub InsertSqlSyncHistory(addedCount As Integer, removedCount As Integer, recoveredCount As Integer, lostCount As Integer, notes As String)
        Dim c As New SqlCommand
        Try
            If objconnection.State = ConnectionState.Closed Then
                objconnection.Open()
            End If
            With c
                .Connection = objconnection
                .CommandText = "INSERT INTO SyncHistory (AddedCount, RemovedCount, RecoveredCount, LostCount, Notes, syncDate)" &
                    "Values(@AddedCount, @RemovedCount, @RecoveredCount, @LostCount, @Notes, @syncDate)"
                .CommandType = CommandType.Text
                .Parameters.AddWithValue("@AddedCount", addedCount)
                .Parameters.AddWithValue("@RemovedCount", removedCount)
                .Parameters.AddWithValue("@RecoveredCount", recoveredCount)
                .Parameters.AddWithValue("@LostCount", lostCount)
                .Parameters.AddWithValue("@Notes", notes)
                .Parameters.AddWithValue("@syncDate", Now)
                .ExecuteNonQuery()
            End With
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub
#End Region

#Region "Update"
    Public Shared Sub UpdateSqlPlaylist(playlistID As String, title As String, description As String, itemCount As String)
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
#End Region

#Region "Delete"
    Public Shared Sub DeleteSqlPlaylistItem(videoID As String, playlistID As String)
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

    Public Shared Sub DeleteSqlPlaylist(playlistID As String)
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
#End Region

#Region "Reseed"
    Public Shared Sub ReseedPlaylistItemsID()
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

    Public Shared Sub ReseedPlaylistID()
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
#End Region
End Class
