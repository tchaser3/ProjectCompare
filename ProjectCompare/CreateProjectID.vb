'Title:         Create Project ID
'Date:          3-2-15
'Author:        Terry Holmes

'Description:   This is used to create the project id

Option Strict On

Public Class CreateProjectID

    Private TheInternalProjectsIDDataSet As InternalProjectsIDDataSet
    Private TheInternalProjectsIDDataTier As InternalProjectsDataTier
    Private WithEvents TheInternalProjectsIDBindingSource As BindingSource

    Private Sub CreateProjectID_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Setting local variables
        Dim intCreatedID As Integer

        'This will load and edit thec controls
        Try

            'Loading the control
            TheInternalProjectsIDDataTier = New InternalProjectsDataTier
            TheInternalProjectsIDDataSet = TheInternalProjectsIDDataTier.GetInternalProjectsIDInformation
            TheInternalProjectsIDBindingSource = New BindingSource

            'Setting up the binding source
            With TheInternalProjectsIDBindingSource
                .DataSource = TheInternalProjectsIDDataSet
                .DataMember = "internalprojectsidcreation"
                .MoveFirst()
                .MoveLast()
            End With

            'Setting up the combo box
            With cboTransactionID
                .DataSource = TheInternalProjectsIDBindingSource
                .DisplayMember = "TransactionID"
                .DataBindings.Add("text", TheInternalProjectsIDBindingSource, "TransactionID", False, DataSourceUpdateMode.Never)
            End With

            'Setting up the rest of the controls
            txtInternalProjectID.DataBindings.Add("text", TheInternalProjectsIDBindingSource, "InternalProjectID")

            'this will create the new id
            intCreatedID = CInt(txtInternalProjectID.Text)
            intCreatedID = intCreatedID + 1
            txtInternalProjectID.Text = CStr(intCreatedID)
            Logon.mintCreatedTransactionID = CInt(intCreatedID)

            'Saving the record
            TheInternalProjectsIDBindingSource.EndEdit()
            TheInternalProjectsIDDataTier.UpdateInternalProjectsIDDB(TheInternalProjectsIDDataSet)

            'this will close the form
            Me.Close()

        Catch ex As Exception

            'this will display if there is a problem with the load
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

End Class