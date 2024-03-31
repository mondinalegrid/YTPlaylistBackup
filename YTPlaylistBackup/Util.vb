Public Class Util

#Region "Variables"
    Public Shared Property ClientID As String = Configuration.ConfigurationManager.AppSettings("ClientID")
    Public Shared Property ClientSecret As String = Configuration.ConfigurationManager.AppSettings("ClientSecret")
    Public Shared Property ChannelId As String = Configuration.ConfigurationManager.AppSettings("ChannelId")
    Public Shared Property SqlConnection As String = Configuration.ConfigurationManager.AppSettings("SqlConnection")
#End Region

End Class
