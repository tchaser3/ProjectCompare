'Title:         Part Number Data Tier
'Date:          2-13-15
'Author:        Terry Holmes

'Description:   This is the data tier for all part numbers

Public Class PartNumberDataTier

    Private aPartNumberDataSetTableAdapters As PartNumberDataSetTableAdapters.partnumbersTableAdapter
    Private aPartNumberDataSet As PartNumberDataSet

    Public Function GetPartNumberInformation() As PartNumberDataSet

        'Setting up the Datatier
        Try
            aPartNumberDataSet = New PartNumberDataSet
            aPartNumberDataSetTableAdapters = New PartNumberDataSetTableAdapters.partnumbersTableAdapter
            aPartNumberDataSetTableAdapters.Fill(aPartNumberDataSet.partnumbers)
            Return aPartNumberDataSet

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Check", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return aPartNumberDataSet
        End Try
    End Function

    Public Sub UpdatePartNumberDB(ByVal aPartNumberDataSet As PartNumberDataSet)

        'This will update the database
        Try
            aPartNumberDataSetTableAdapters.Update(aPartNumberDataSet.partnumbers)
        Catch ex As Exception

        End Try

    End Sub


End Class
