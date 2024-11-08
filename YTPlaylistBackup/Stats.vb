Public Class Stats
    Private Sub Stats_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetForm(Me)
        CenterToParent()
        SetTextbox()
        SetTextBoxValue()
    End Sub

    Private Sub SetTextbox()
        For Each ctrl As Control In Me.Controls
            If TypeOf ctrl Is TextBox Then
                Dim txtBox As TextBox = CType(ctrl, TextBox)
                txtBox.ReadOnly = True
                txtBox.Text = "N/A"
            End If
        Next
    End Sub

    Private Sub SetTextBoxValue()
        Dim syncHistory = Database.syncHistoryData
        Dim playlistListData = Database.playlistListData
        Dim playlistItemListData = Database.playlistItemListData
        Dim playlistItemListRecoveredData = Database.playlistItemListRecoveredData
        Dim playlistItemListRemovedData = Database.playlistItemListRemovedData
        Dim playlistItemListLostData = Database.playlistItemListLostData

        'Last Sync Date
        Dim lastSyncDate As DataRow = syncHistory.AsEnumerable() _
                             .OrderByDescending(Function(row) row.Field(Of Date)("syncDate")) _
                             .FirstOrDefault()
        If lastSyncDate IsNot Nothing Then
            TextBox1.Text = lastSyncDate.Field(Of Date)("syncDate").ToString
        End If

        'Last Synced Playlist
        Dim lastSyncPlaylist As DataRow = playlistListData.AsEnumerable() _
                             .OrderByDescending(Function(row) row.Field(Of Date)("syncDate")) _
                             .FirstOrDefault()
        If lastSyncPlaylist IsNot Nothing Then
            TextBox2.Text = $"({lastSyncPlaylist("syncDate")}) {lastSyncPlaylist("title")}"
        End If

        Dim totalAdded As Integer = 0
        Dim totalRemoved As Integer = 0
        Dim totalRecovered As Integer = 0
        Dim totalLost As Integer = 0
        Dim totalCount As Integer = 0

        'Total Syncs (A+R+R+L)
        For Each row As DataRow In syncHistory.Rows
            totalAdded += row.Field(Of Integer)("AddedCount")
            totalRemoved += row.Field(Of Integer)("RemovedCount")
            totalRecovered += row.Field(Of Integer)("RecoveredCount")
            totalLost += row.Field(Of Integer)("LostCount")
        Next
        totalCount = totalAdded + totalRemoved + totalRecovered + totalLost
        If totalCount > 0 Then
            TextBox3.Text = totalCount.ToString
        End If

        'Total Added Sync
        If totalAdded > 0 Then
            TextBox4.Text = totalAdded.ToString
        End If

        'Total Removed Sync
        If totalRemoved > 0 Then
            TextBox5.Text = totalRemoved.ToString
        End If

        'Total Recovered Sync
        If totalRecovered > 0 Then
            TextBox6.Text = totalRecovered.ToString
        End If

        'Total Lost Sync
        If totalLost > 0 Then
            TextBox7.Text = totalLost.ToString
        End If

        Dim playlistSyncCounts As New Dictionary(Of String, Integer)
        Dim mostSyncedPlaylistID As String = Nothing
        Dim highestSyncCount As Integer = Integer.MinValue
        Dim leastSyncedPlaylistID As String = Nothing
        Dim lowestSyncCount As Integer = Integer.MaxValue
        Dim playlistRecoveredCounts As New Dictionary(Of String, Integer)
        Dim playlistRemovedCounts As New Dictionary(Of String, Integer)
        Dim playlistLostCounts As New Dictionary(Of String, Integer)

        ' Add counts from the Recovered DataTable
        For Each row As DataRow In playlistItemListRecoveredData.Rows
            Dim playlistID As String = row.Field(Of String)("playlistID")
            If Not playlistSyncCounts.ContainsKey(playlistID) Then
                playlistSyncCounts(playlistID) = 0
            End If
            playlistSyncCounts(playlistID) += 1
            If Not playlistRecoveredCounts.ContainsKey(playlistID) Then
                playlistRecoveredCounts(playlistID) = 0
            End If
            playlistRecoveredCounts(playlistID) += 1
        Next

        ' Add counts from the Removed DataTable
        For Each row As DataRow In playlistItemListRemovedData.Rows
            Dim playlistID As String = row.Field(Of String)("playlistID")
            If Not playlistSyncCounts.ContainsKey(playlistID) Then
                playlistSyncCounts(playlistID) = 0
            End If
            playlistSyncCounts(playlistID) += 1
            If Not playlistRemovedCounts.ContainsKey(playlistID) Then
                playlistRemovedCounts(playlistID) = 0
            End If
            playlistRemovedCounts(playlistID) += 1
        Next

        ' Add counts from the Lost DataTable
        For Each row As DataRow In playlistItemListLostData.Rows
            Dim playlistID As String = row.Field(Of String)("playlistID")
            If Not playlistSyncCounts.ContainsKey(playlistID) Then
                playlistSyncCounts(playlistID) = 0
            End If
            playlistSyncCounts(playlistID) += 1
            If Not playlistLostCounts.ContainsKey(playlistID) Then
                playlistLostCounts(playlistID) = 0
            End If
            playlistLostCounts(playlistID) += 1
        Next

        ' Find the playlist with the highest total
        For Each playlistID As String In playlistSyncCounts.Keys
            Dim totalSyncCount As Integer = playlistSyncCounts(playlistID)
            If totalSyncCount > highestSyncCount Then
                highestSyncCount = totalSyncCount
                mostSyncedPlaylistID = playlistID
            End If
            If totalSyncCount < lowestSyncCount Then
                lowestSyncCount = totalSyncCount
                leastSyncedPlaylistID = playlistID
            End If
        Next
        Dim mostRecoveredPlaylistID As String = Nothing
        Dim maxRecoveredCount As Integer = Integer.MinValue
        For Each playlistID As String In playlistRecoveredCounts.Keys
            Dim recoveredCount As Integer = playlistRecoveredCounts(playlistID)
            If recoveredCount > maxRecoveredCount Then
                maxRecoveredCount = recoveredCount
                mostRecoveredPlaylistID = playlistID
            End If
        Next
        Dim mostRemovedPlaylistID As String = Nothing
        Dim maxRemovedCount As Integer = Integer.MinValue
        For Each playlistID As String In playlistRemovedCounts.Keys
            Dim removedCount As Integer = playlistRemovedCounts(playlistID)
            If removedCount > maxRemovedCount Then
                maxRemovedCount = removedCount
                mostRemovedPlaylistID = playlistID
            End If
        Next
        Dim mostLostPlaylistID As String = Nothing
        Dim maxLostCount As Integer = Integer.MinValue
        For Each playlistID As String In playlistLostCounts.Keys
            Dim lostCount As Integer = playlistLostCounts(playlistID)
            If lostCount > maxLostCount Then
                maxLostCount = lostCount
                mostLostPlaylistID = playlistID
            End If
        Next

        'Most Synced Playlist (R+R+L)
        If Not String.IsNullOrWhiteSpace(mostSyncedPlaylistID) AndAlso highestSyncCount > 0 Then
            'mostSyncedPlaylistID
            Dim playlistItem As DataRow = playlistListData.AsEnumerable().Where(Function(row) row("playlistID").ToString = mostSyncedPlaylistID.ToString).FirstOrDefault()
            TextBox8.Text = $"({highestSyncCount}) {playlistItem("title")}"
        End If

        'Most Videos In A Playlist
        Dim mostVideos As DataRow = playlistListData.AsEnumerable() _
            .OrderByDescending(Function(row) row.Field(Of Integer)("itemCount")) _
            .FirstOrDefault()
        If mostVideos IsNot Nothing Then
            TextBox9.Text = $"({mostVideos("itemCount")}) {mostVideos("title")}"
        End If

        'Most Synced Month
        Dim highestTotal As Integer = Integer.MinValue
        Dim highestTotalRow As DataRow = Nothing
        For Each row As DataRow In syncHistory.Rows
            Dim total As Integer = row.Field(Of Integer)("AddedCount") +
                                   row.Field(Of Integer)("RemovedCount") +
                                   row.Field(Of Integer)("RecoveredCount") +
                                   row.Field(Of Integer)("LostCount")
            If total > highestTotal Then
                highestTotal = total
                highestTotalRow = row
            End If
        Next
        If highestTotalRow IsNot Nothing Then
            ' Access values from the highestTotalRow
            Dim syncDate As Date = highestTotalRow.Field(Of Date)("syncDate")
            Dim addedCount As Integer = highestTotalRow.Field(Of Integer)("AddedCount")
            Dim removedCount As Integer = highestTotalRow.Field(Of Integer)("RemovedCount")
            Dim recoveredCount As Integer = highestTotalRow.Field(Of Integer)("RecoveredCount")
            Dim lostCount As Integer = highestTotalRow.Field(Of Integer)("LostCount")

            TextBox10.Text = $"({highestTotal}) {syncDate}"
        End If

        'Least Videos In A Playlist
        Dim leastVideos As DataRow = playlistListData.AsEnumerable() _
            .OrderBy(Function(row) row.Field(Of Integer)("itemCount")) _
            .FirstOrDefault()
        If leastVideos IsNot Nothing Then
            TextBox11.Text = $"({leastVideos("itemCount")}) {leastVideos("title")}"
        End If

        'Least Synced Playlist
        If Not String.IsNullOrWhiteSpace(leastSyncedPlaylistID) AndAlso lowestSyncCount > 0 Then
            'mostSyncedPlaylistID
            Dim playlistItem As DataRow = playlistListData.AsEnumerable().Where(Function(row) row("playlistID").ToString = leastSyncedPlaylistID.ToString).FirstOrDefault()
            TextBox12.Text = $"({lowestSyncCount}) {playlistItem("title")}"
        End If

        'Oldest Not Synced Playlist
        Dim oldestSyncPlaylist As DataRow = playlistListData.AsEnumerable() _
                             .OrderBy(Function(row) row.Field(Of Date)("syncDate")) _
                             .FirstOrDefault()
        If oldestSyncPlaylist IsNot Nothing Then
            TextBox13.Text = $"({oldestSyncPlaylist("syncDate")}) {oldestSyncPlaylist("title")}"
        End If

        'Most Synced Playlist Removed
        If Not String.IsNullOrWhiteSpace(mostRemovedPlaylistID) AndAlso maxRemovedCount > 0 Then
            'mostSyncedPlaylistID
            Dim playlistItem As DataRow = playlistListData.AsEnumerable().Where(Function(row) row("playlistID").ToString = mostRemovedPlaylistID.ToString).FirstOrDefault()
            TextBox14.Text = $"({maxRemovedCount}) {playlistItem("title")}"
        End If

        'Most Synced Playlist Recovered
        If Not String.IsNullOrWhiteSpace(mostRecoveredPlaylistID) AndAlso maxRecoveredCount > 0 Then
            'mostSyncedPlaylistID
            Dim playlistItem As DataRow = playlistListData.AsEnumerable().Where(Function(row) row("playlistID").ToString = mostRecoveredPlaylistID.ToString).FirstOrDefault()
            TextBox15.Text = $"({maxRecoveredCount}) {playlistItem("title")}"
        End If

        'Most Synced Playlist Lost
        If Not String.IsNullOrWhiteSpace(mostLostPlaylistID) AndAlso maxLostCount > 0 Then
            'mostSyncedPlaylistID
            Dim playlistItem As DataRow = playlistListData.AsEnumerable().Where(Function(row) row("playlistID").ToString = mostLostPlaylistID.ToString).FirstOrDefault()
            TextBox16.Text = $"({maxLostCount}) {playlistItem("title")}"
        End If
    End Sub
End Class