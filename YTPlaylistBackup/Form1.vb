Imports System.Threading
Imports System.Windows
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Services
Imports Google.Apis.Util.Store
Imports Google.Apis.YouTube.v3
Imports Microsoft.Data.SqlClient

Public Class Form1
    Private credential As UserCredential
    Private ytService As YouTubeService
    Private objconnection As New SqlConnection(Util.SqlConnection)

    'reference https://stackoverflow.com/questions/65357223/get-youtube-channel-data-using-google-youtube-data-api-in-vb-net

    Private Async Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()

        'GetSqlPlaylistList()
        'GetSqlPlaylistItemsList()

        Await OAuth()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GetPlaylistList()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        GetPlaylistItemLists()
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
           New ClientSecrets With {
            .ClientId = Util.ClientID,
            .ClientSecret = Util.ClientSecret
            },
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
        While nextPageToken IsNot Nothing
            Dim playlistListReq = ytService.Playlists.List("id,contentDetails.itemCount,snippet.title,snippet.description")
            playlistListReq.Mine = True
            playlistListReq.MaxResults = 50
            playlistListReq.PageToken = nextPageToken
            Dim playlistListResp = playlistListReq.Execute

            For Each playlists In playlistListResp.Items
                Dim id = playlists.Id
                Dim title = playlists.Snippet.Title
                Dim desc = playlists.Snippet.Description
                Dim itemCount = playlists.ContentDetails.ItemCount
                'save to datatable

                'first get from sql server
                'then check if existing if not
                'add to datatable
            Next

            nextPageToken = playlistListResp.NextPageToken
        End While
    End Sub

    Private Sub GetPlaylistItemLists()
        Dim nextPageToken As String = ""
        While nextPageToken IsNot Nothing
            Dim playlistItemListReq = ytService.PlaylistItems.List("contentDetails.videoId,snippet.title,snippet.description,snippet.videoOwnerChannelTitle,snippet.videoOwnerChannelId")
            playlistItemListReq.PlaylistId = ""
            playlistItemListReq.MaxResults = 50
            playlistItemListReq.PageToken = nextPageToken
        End While
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
                .CommandText = "SELECT * from Playlists"
                .CommandType = CommandType.Text
                dr = .ExecuteReader
            End With
            If dr.Read Then

            Else

            End If
            dr.Close()
        Catch ex As Exception
        End Try
        objconnection.Close()
    End Sub

    Private Sub GetSqlPlaylistItemsList()

    End Sub


End Class
