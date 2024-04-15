'Title:     Employee Data Tier
'Date:      3/28/13
'Author:    Terry Holmes

Option Strict On

Public Class EmployeeDataTier
    'Setting up modular variables
    Private aEmployeesTableAdapter As EmployeesDataSetTableAdapters.employeesTableAdapter

    Private aEmployeesDataSet As EmployeesDataSet

    Public Function GetEmployeesInformation() As EmployeesDataSet

        'Setting up the Datatier
        Try
            aEmployeesDataSet = New EmployeesDataSet
            aEmployeesTableAdapter = New EmployeesDataSetTableAdapters.employeesTableAdapter
            aEmployeesTableAdapter.Fill(aEmployeesDataSet.employees)
            Return aEmployeesDataSet

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Check", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return aEmployeesDataSet
        End Try
    End Function

    Public Sub UpdateDB(ByVal aEmployeesDataSet As EmployeesDataSet)

        'This will update the database
        Try
            aEmployeesTableAdapter.Update(aEmployeesDataSet.employees)
        Catch ex As Exception

        End Try
    End Sub
End Class
