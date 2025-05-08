Imports System.Data.SQLite
Imports System.IO

Public Class Database

#Region "Datatable"
    Public Shared Property playlistListData As New DataTable
    Public Shared Property playlistItemListData As New DataTable
    Public Shared Property playlistItemListRecoveredData As New DataTable
    Public Shared Property playlistItemListRemovedData As New DataTable
    Public Shared Property playlistItemListLostData As New DataTable
    Public Shared Property syncHistoryData As New DataTable
#End Region

#Region "Sub/Func"
    Public Shared Sub InitDB()
        Dim directoryPath As String = Environment.CurrentDirectory & "\db"

        Dim databaseFilePath As String = Path.Combine(directoryPath, "ytplaylist.sqlite")

        If Not Directory.Exists(directoryPath) Then
            Directory.CreateDirectory(directoryPath)
        End If

        SqlConn = $"Data Source={databaseFilePath};Version=3;"
    End Sub

    Public Shared Sub CreateTables()
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    comm.CommandText =
                      "CREATE TABLE IF NOT EXISTS
                      [SyncHistory] 
                      (
                        [id]     INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                        [AddedCount]   INTEGER NULL,
                        [RemovedCount]   INTEGER NULL,
                        [RecoveredCount]   INTEGER NULL,
                        [LostCount]   INTEGER NULL,
                        [Notes]   TEXT NULL,
                        [syncDate]   TEXT NULL
                      )"
                    comm.ExecuteNonQuery()
                End Using

                Using comm As SQLiteCommand = conn.CreateCommand()
                    comm.CommandText =
                      "CREATE TABLE IF NOT EXISTS
                      [Playlists] 
                      (
                        [id]     INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                        [playlistID]   TEXT NULL,
                        [title]   TEXT NULL,
                        [description]   TEXT NULL,
                        [itemCount]   INTEGER NULL,
                        [syncDate]   TEXT NULL
                      )"
                    comm.ExecuteNonQuery()
                End Using

                Using comm As SQLiteCommand = conn.CreateCommand()
                    comm.CommandText =
                      "CREATE TABLE IF NOT EXISTS
                      [PlaylistItems] 
                      (
                        [id]     INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                        [playlistID]   TEXT NULL,
                        [videoID]   TEXT NULL,
                        [title]   TEXT NULL,
                        [description]   TEXT NULL,
                        [videoOwnerChannelId]   TEXT NULL,
                        [videoOwnerChannelTitle]   TEXT NULL,
                        [syncDate]   TEXT NULL
                      )"
                    comm.ExecuteNonQuery()
                End Using

                Using comm As SQLiteCommand = conn.CreateCommand()
                    comm.CommandText =
                      "CREATE TABLE IF NOT EXISTS
                      [PlaylistItemsRecovered] 
                      (
                        [id]     INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                        [playlistID]   TEXT NULL,
                        [videoID]   TEXT NULL,
                        [title]   TEXT NULL,
                        [description]   TEXT NULL,
                        [videoOwnerChannelId]   TEXT NULL,
                        [videoOwnerChannelTitle]   TEXT NULL,
                        [syncDate]   TEXT NULL
                      )"
                    comm.ExecuteNonQuery()
                End Using

                Using comm As SQLiteCommand = conn.CreateCommand()
                    comm.CommandText =
                      "CREATE TABLE IF NOT EXISTS
                      [PlaylistItemsRemoved] 
                      (
                        [id]     INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                        [playlistID]   TEXT NULL,
                        [videoID]   TEXT NULL,
                        [title]   TEXT NULL,
                        [description]   TEXT NULL,
                        [videoOwnerChannelId]   TEXT NULL,
                        [videoOwnerChannelTitle]   TEXT NULL,
                        [syncDate]   TEXT NULL
                      )"
                    comm.ExecuteNonQuery()
                End Using

                Using comm As SQLiteCommand = conn.CreateCommand()
                    comm.CommandText =
                      "CREATE TABLE IF NOT EXISTS
                      [PlaylistItemsLost] 
                      (
                        [id]     INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                        [playlistID]   TEXT NULL,
                        [videoID]   TEXT NULL,
                        [title]   TEXT NULL,
                        [description]   TEXT NULL,
                        [videoOwnerChannelId]   TEXT NULL,
                        [videoOwnerChannelTitle]   TEXT NULL,
                        [syncDate]   TEXT NULL
                      )"
                    comm.ExecuteNonQuery()
                End Using

            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using
    End Sub

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
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    Dim dr As SQLiteDataReader
                    comm.CommandText = "SELECT * from Playlists ORDER BY id"
                    dr = comm.ExecuteReader

                    If playlistListData.Rows.Count > 0 AndAlso dr.HasRows Then
                        playlistListData.Rows.Clear()
                        playlistListData.AcceptChanges()
                    End If
                    While dr.Read
                        playlistListData.Rows.Add(dr(1), dr(2), dr(3), dr(4), dr(5))
                    End While
                    playlistListData.AcceptChanges()
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using
    End Sub

    Public Shared Sub GetSqlPlaylistItemsList()
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    Dim dr As SQLiteDataReader
                    comm.CommandText = "SELECT * from PlaylistItems ORDER BY id"
                    dr = comm.ExecuteReader

                    If playlistItemListData.Rows.Count > 0 AndAlso dr.HasRows Then
                        playlistItemListData.Rows.Clear()
                        playlistItemListData.AcceptChanges()
                    End If
                    While dr.Read
                        playlistItemListData.Rows.Add(dr(1), dr(2), dr(3), dr(4), dr(5), dr(6), dr(7))
                    End While
                    playlistItemListData.AcceptChanges()
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using
    End Sub

    Public Shared Sub GetSqlPlaylistItemsRecoveredList()
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    Dim dr As SQLiteDataReader
                    comm.CommandText = "SELECT * from PlaylistItemsRecovered ORDER BY id"
                    dr = comm.ExecuteReader

                    If playlistItemListRecoveredData.Rows.Count > 0 AndAlso dr.HasRows Then
                        playlistItemListRecoveredData.Rows.Clear()
                        playlistItemListRecoveredData.AcceptChanges()
                    End If
                    While dr.Read
                        playlistItemListRecoveredData.Rows.Add(dr(1), dr(2), dr(3), dr(4), dr(5), dr(6), dr(7))
                    End While
                    playlistItemListRecoveredData.AcceptChanges()
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using
    End Sub

    Public Shared Sub GetSqlPlaylistItemsRemovedList()
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    Dim dr As SQLiteDataReader
                    comm.CommandText = "SELECT * from PlaylistItemsRemoved ORDER BY id"
                    dr = comm.ExecuteReader

                    If playlistItemListRemovedData.Rows.Count > 0 AndAlso dr.HasRows Then
                        playlistItemListRemovedData.Rows.Clear()
                        playlistItemListRemovedData.AcceptChanges()
                    End If
                    While dr.Read
                        playlistItemListRemovedData.Rows.Add(dr(1), dr(2), dr(3), dr(4), dr(5), dr(6), dr(7))
                    End While
                    playlistItemListRemovedData.AcceptChanges()
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using
    End Sub

    Public Shared Sub GetSqlPlaylistItemsLostList()
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    Dim dr As SQLiteDataReader
                    comm.CommandText = "SELECT * from PlaylistItemsLost ORDER BY id"
                    dr = comm.ExecuteReader

                    If playlistItemListLostData.Rows.Count > 0 AndAlso dr.HasRows Then
                        playlistItemListLostData.Rows.Clear()
                        playlistItemListLostData.AcceptChanges()
                    End If
                    While dr.Read
                        playlistItemListLostData.Rows.Add(dr(1), dr(2), dr(3), dr(4), dr(5), dr(6), dr(7))
                    End While
                    playlistItemListLostData.AcceptChanges()
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using
    End Sub

    Public Shared Sub GetSqlSyncHistory()
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    Dim dr As SQLiteDataReader
                    comm.CommandText = "SELECT * from SyncHistory ORDER BY syncDate ASC"
                    dr = comm.ExecuteReader

                    If syncHistoryData.Rows.Count > 0 AndAlso dr.HasRows Then
                        syncHistoryData.Rows.Clear()
                        syncHistoryData.AcceptChanges()
                    End If
                    While dr.Read
                        syncHistoryData.Rows.Add(dr(6), dr(1), dr(2), dr(3), dr(4), dr(5))
                    End While
                    syncHistoryData.AcceptChanges()
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using
    End Sub

    Public Shared Sub GetSqlAll()
        InitDB()
        CreateTables()
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
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    With comm
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
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using
    End Sub

    Public Shared Sub InsertSqlPlaylistItemList(playlistID As String, videoID As String, title As String, description As String, videoOwnerChannelId As String, videoOwnerChannelTitle As String)
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    With comm
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
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using
    End Sub

    Public Shared Sub InsertSqlPlaylistItemRecoveredList(playlistID As String, videoID As String, title As String, description As String, videoOwnerChannelId As String, videoOwnerChannelTitle As String)
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    With comm
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
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using
    End Sub

    Public Shared Sub InsertSqlPlaylistItemRemovedList(playlistID As String, videoID As String, title As String, description As String, videoOwnerChannelId As String, videoOwnerChannelTitle As String)
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    With comm
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
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using
    End Sub

    Public Shared Sub InsertSqlPlaylistItemLostList(playlistID As String, videoID As String, title As String, description As String, videoOwnerChannelId As String, videoOwnerChannelTitle As String)
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    With comm
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
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using
    End Sub

    Public Shared Sub InsertSqlSyncHistory(addedCount As Integer, removedCount As Integer, recoveredCount As Integer, lostCount As Integer, notes As String)
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    With comm
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
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using
    End Sub
#End Region

#Region "Update"
    Public Shared Sub UpdateSqlPlaylist(playlistID As String, title As String, description As String, itemCount As String)
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    With comm
                        .CommandText = "UPDATE Playlists SET title=@title, description=@description, itemCount=@itemCount, syncDate=@syncDate WHERE playlistID=@playlistID"
                        .CommandType = CommandType.Text
                        .Parameters.AddWithValue("@playlistID", playlistID)
                        .Parameters.AddWithValue("@title", title)
                        .Parameters.AddWithValue("@description", description)
                        .Parameters.AddWithValue("@itemCount", itemCount)
                        .Parameters.AddWithValue("@syncDate", Now)
                        .ExecuteNonQuery()
                    End With
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using
    End Sub
#End Region

#Region "Delete"
    Public Shared Sub DeleteSqlPlaylistItem(videoID As String, playlistID As String)
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    With comm
                        .CommandText = "DELETE PlaylistItems WHERE videoID=@videoID and playlistID=@playlistID"
                        .CommandType = CommandType.Text
                        .Parameters.AddWithValue("@videoID", videoID)
                        .Parameters.AddWithValue("@playlistID", playlistID)
                        .ExecuteNonQuery()
                    End With
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using

        ReseedPlaylistItemsID()
    End Sub

    Public Shared Sub DeleteSqlPlaylist(playlistID As String)
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    With comm
                        .CommandText = "DELETE Playlists WHERE playlistID=@playlistID"
                        .CommandType = CommandType.Text
                        .Parameters.AddWithValue("@playlistID", playlistID)
                        .ExecuteNonQuery()
                    End With
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using

        ReseedPlaylistID()
    End Sub
#End Region

#Region "Reseed"
    Public Shared Sub ReseedPlaylistItemsID()
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    With comm
                        .CommandText = "UPDATE sqlite_sequence SET seq = (SELECT IFNULL(MAX(id), 0) FROM PlaylistItems) WHERE name = 'PlaylistItems';"
                        .CommandType = CommandType.Text
                        .ExecuteNonQuery()
                    End With
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using
    End Sub

    Public Shared Sub ReseedPlaylistID()
        Using conn As New SQLiteConnection(SqlConn)
            Try
                conn.Open()
                Using comm As SQLiteCommand = conn.CreateCommand()
                    With comm
                        .CommandText = "UPDATE sqlite_sequence SET seq = (SELECT IFNULL(MAX(id), 0) FROM Playlists) WHERE name = 'Playlists';"
                        .CommandType = CommandType.Text
                        .ExecuteNonQuery()
                    End With
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error")
            End Try
        End Using
    End Sub
#End Region
End Class
