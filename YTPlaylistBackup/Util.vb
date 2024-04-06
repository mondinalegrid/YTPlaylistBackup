Imports Microsoft.Win32

Public Class Util

#Region "Variables"
    Public Shared Property ClientID As String = Configuration.ConfigurationManager.AppSettings("ClientID")
    Public Shared Property ClientSecret As String = Configuration.ConfigurationManager.AppSettings("ClientSecret")
    Public Shared Property SqlConnection As String = Configuration.ConfigurationManager.AppSettings("SqlConnection")
#End Region

#Region "Sub/Func"
    Public Shared Sub SetForm(ByRef Form As Form, Optional Resizable As Boolean = False)
        Form.AutoScroll = True
        Form.AutoScrollMargin = New System.Drawing.Size(0, 7)
        If Not Resizable Then
            Form.AutoSize = True
            Form.FormBorderStyle = FormBorderStyle.FixedDialog
            Form.MaximizeBox = False
        End If
    End Sub

    Public Shared Sub InitDGV(ByRef DGV As DataGridView)
        DGV.ReadOnly = True
        DGV.AllowUserToAddRows = False
        DGV.AllowUserToDeleteRows = False
        DGV.AllowUserToOrderColumns = False
        DGV.AllowUserToResizeRows = False
        DGV.RowHeadersWidth = 35
        DGV.DataSource = Nothing
        DGV.Rows.Clear()
        DGV.Columns.Clear()
    End Sub

    Public Shared Sub SetFilterDataGridViewData(ByVal _Data As DataTable, ByRef _DataGridView As DataGridView, Optional _CustomFilter As String = "")
        Dim DataFiltered = _Data.DefaultView

        DataFiltered.RowFilter = If(_CustomFilter <> "", _CustomFilter, "")

        _DataGridView.DataSource = DataFiltered
    End Sub

    Public Shared Function CheckDGVCellValue(ByRef CellValue As Object) As Boolean
        If CellValue IsNot Nothing AndAlso CellValue IsNot DBNull.Value AndAlso Not String.IsNullOrEmpty(CellValue.ToString) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Sub InitCombobox(ByRef cmbBox As ComboBox)
        cmbBox.DropDownStyle = ComboBoxStyle.DropDownList
        cmbBox.FlatStyle = FlatStyle.Popup
        cmbBox.DisplayMember = "Value"
        cmbBox.ValueMember = "Key"
    End Sub

    Public Shared Sub OpenLink(link As String)
        Dim process As New Process
        process.StartInfo.UseShellExecute = True
        process.StartInfo.FileName = link
        process.Start()
    End Sub
#End Region

End Class
