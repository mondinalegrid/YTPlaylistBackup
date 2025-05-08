Public Enum DataTables As Integer
    playlistList = 0
    playlistItemList
    playlistItemListRecovered
    playlistItemListRemoved
    playlistItemListLost
    syncHistory
End Enum

Public Enum PlaylistListColumns As Integer
    playlistID = 0
    title
    description
    itemCount
    syncDate
End Enum

Public Enum PlaylistItemListColumns As Integer
    playlistID = 0
    videoID
    title
    description
    videoOwnerChannelId
    videoOwnerChannelTitle
    syncDate
End Enum

Public Enum PlaylistItemListRecoveredColumns As Integer
    playlistID = 0
    videoID
    title
    description
    videoOwnerChannelId
    videoOwnerChannelTitle
    syncDate
End Enum

Public Enum PlaylistItemListRemovedColumns As Integer
    playlistID = 0
    videoID
    title
    description
    videoOwnerChannelId
    videoOwnerChannelTitle
    syncDate
End Enum

Public Enum PlaylistItemListLostColumns As Integer
    playlistID = 0
    videoID
    title
    description
    videoOwnerChannelId
    videoOwnerChannelTitle
    syncDate
End Enum

Public Enum SyncHistoryColumns As Integer
    syncDate = 0
    AddedCount
    RemovedCount
    RecoveredCount
    LostCount
    notes
End Enum