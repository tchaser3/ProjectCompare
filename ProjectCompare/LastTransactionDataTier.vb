'Title:         Last Transaction Data Tier
'Date:          12-23-14
'Author:        Terry Holmes

'Description:   This is the data tier for the last transaction information

Option Strict On

Public Class LastTransactionDataTier

    Private aLastTransactionDataSetTableAdapter As LastTransactionDataSetTableAdapters.lasttransactionTableAdapter
    Private aLastTransactionDataSet As LastTransactionDataSet

    Private aLastTransactionIDDataSetTableAdapter As LastTransactionIDDataSetTableAdapters.lasttransactionidTableAdapter
    Private aLastTransactionIDDataSet As LastTransactionIDDataSet

    Public Function GetLastTransactionInformation() As LastTransactionDataSet

        'Setting up the Datatier
        Try
            aLastTransactionDataSet = New LastTransactionDataSet
            aLastTransactionDataSetTableAdapter = New LastTransactionDataSetTableAdapters.lasttransactionTableAdapter
            aLastTransactionDataSetTableAdapter.Fill(aLastTransactionDataSet.lasttransaction)
            Return aLastTransactionDataSet

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Check", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return aLastTransactionDataSet
        End Try
    End Function

    Public Sub UpdateLastTransactionDB(ByVal aLastTransactionDataSet As LastTransactionDataSet)

        'This will update the database
        Try
            aLastTransactionDataSetTableAdapter.Update(aLastTransactionDataSet.lasttransaction)
        Catch ex As Exception

        End Try
    End Sub

    Public Function GetLastTransactionIDInformation() As LastTransactionIDDataSet

        'Setting up the Datatier
        Try
            aLastTransactionIDDataSet = New LastTransactionIDDataSet
            aLastTransactionIDDataSetTableAdapter = New LastTransactionIDDataSetTableAdapters.lasttransactionidTableAdapter
            aLastTransactionIDDataSetTableAdapter.Fill(aLastTransactionIDDataSet.lasttransactionid)
            Return aLastTransactionIDDataSet

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Check", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return aLastTransactionIDDataSet
        End Try
    End Function

    Public Sub UpdateLastTransactionIDDB(ByVal aLastTransactionIDDataSet As LastTransactionIDDataSet)

        'This will update the database
        Try
            aLastTransactionIDDataSetTableAdapter.Update(aLastTransactionIDDataSet.lasttransactionid)
        Catch ex As Exception

        End Try
    End Sub

End Class
