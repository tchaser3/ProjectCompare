'Title:         Internal Projects Data Tier
'Data:          1-21-15
'Author:        Terry Holmes

'Description:   This is the Data Tier for the Internal Projects Form

Option Strict On

Public Class InternalProjectsDataTier

    'Setting Up Internal Projects Modular Variables
    Private aInternalProjectsDataSetTableAdapter As InternalProjectsDataSetTableAdapters.internalprojectsTableAdapter
    Private aInternalProjectsDataSet As InternalProjectsDataSet

    'Setting Up Internal Projects Modular Variables
    Private aInternalProjectsIDDataSetTableAdapter As InternalProjectsIDDataSetTableAdapters.internalprojectsidcreationTableAdapter
    Private aInternalProjectsIDDataSet As InternalProjectsIDDataSet

    Public Function GetInternalProjectsIDInformation() As InternalProjectsIDDataSet

        'Setting up this Function of the Data Tier
        Try

            aInternalProjectsIDDataSet = New InternalProjectsIDDataSet
            aInternalProjectsIDDataSetTableAdapter = New InternalProjectsIDDataSetTableAdapters.internalprojectsidcreationTableAdapter
            aInternalProjectsIDDataSetTableAdapter.Fill(aInternalProjectsIDDataSet.internalprojectsidcreation)
            Return aInternalProjectsIDDataSet

        Catch ex As Exception

            MessageBox.Show(ex.Message, "Please Check", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return aInternalProjectsIDDataSet

        End Try
    End Function

    Public Sub UpdateInternalProjectsIDDB(ByVal aInternalProjectsIDDataSet As InternalProjectsIDDataSet)

        Try

            aInternalProjectsIDDataSetTableAdapter.Update(aInternalProjectsIDDataSet.internalprojectsidcreation)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Public Function GetInternalProjectsInformation() As InternalProjectsDataSet

        'Setting up this Function of the Data Tier
        Try

            aInternalProjectsDataSet = New InternalProjectsDataSet
            aInternalProjectsDataSetTableAdapter = New InternalProjectsDataSetTableAdapters.internalprojectsTableAdapter
            aInternalProjectsDataSetTableAdapter.Fill(aInternalProjectsDataSet.internalprojects)
            Return aInternalProjectsDataSet

        Catch ex As Exception

            MessageBox.Show(ex.Message, "Please Check", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return aInternalProjectsDataSet

        End Try
    End Function

    Public Sub UpdateInternalProjectsDB(ByVal aInternalProjectsDataSet As InternalProjectsDataSet)

        Try

            aInternalProjectsDataSetTableAdapter.Update(aInternalProjectsDataSet.internalprojects)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

End Class
