'Title:         Receive Parts Data Tier
'Date:          2-16-15
'Author:        Terry Holmes

'Description:   This is the data tier for the Receive Parts

Option Strict On

Public Class ReceivePartsDataTier

    Private aReceivePartsDataSetTableAdapter As ReceivePartsDataSetTableAdapters.ReceivedPartsTableAdapter
    Private aReceivePartsDataSet As ReceivePartsDataSet

    Public Function GetReceivePartsInformation() As ReceivePartsDataSet

        'Setting up the Datatier
        Try
            aReceivePartsDataSet = New ReceivePartsDataSet
            aReceivePartsDataSetTableAdapter = New ReceivePartsDataSetTableAdapters.ReceivedPartsTableAdapter
            aReceivePartsDataSetTableAdapter.Fill(aReceivePartsDataSet.ReceivedParts)
            Return aReceivePartsDataSet

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Check", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return aReceivePartsDataSet
        End Try
    End Function

    Public Sub UpdateReceivePartsDB(ByVal aReceivePartsDataSet As ReceivePartsDataSet)

        'This will update the database
        Try
            aReceivePartsDataSetTableAdapter.Update(aReceivePartsDataSet.ReceivedParts)
        Catch ex As Exception

        End Try
    End Sub


End Class
