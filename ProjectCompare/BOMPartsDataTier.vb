'Title:         BOM Parts Data Tier
'Date:          2-17-15
'Author:        Terry Holmes

'Description:   This is the data tier class for the BOM Parts

Option Strict On

Public Class BOMPartsDataTier

    'Setting global variables
    Private aBOMPartsDataSetTableAdapter As BOMPartsDataSetTableAdapters.BOMPartsTableAdapter
    Private aBOMPartsDataSet As BOMPartsDataSet

    Public Function GetBOMPartsInformation() As BOMPartsDataSet

        'Setting up the Datatier
        Try
            aBOMPartsDataSet = New BOMPartsDataSet
            aBOMPartsDataSetTableAdapter = New BOMPartsDataSetTableAdapters.BOMPartsTableAdapter
            aBOMPartsDataSetTableAdapter.Fill(aBOMPartsDataSet.BOMParts)
            Return aBOMPartsDataSet

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Check", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return aBOMPartsDataSet
        End Try
    End Function

    Public Sub UpdateBOMPartsDB(ByVal aBOMPartsDataSet As BOMPartsDataSet)

        'This will update the database
        Try
            aBOMPartsDataSetTableAdapter.Update(aBOMPartsDataSet.BOMParts)
        Catch ex As Exception

        End Try
    End Sub
End Class
