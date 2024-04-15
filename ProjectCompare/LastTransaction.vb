'Title:         Last Transaction
'Date:          2-27-15
'Author:        Terry Holmes

'Description:   This for is used to display/add last transaction

Option Strict On

Public Class LastTransaction

    'Setting up global variables
    Private TheLastTransactionDataSet As LastTransactionDataSet
    Private TheLastTransactionDataTier As LastTransactionDataTier
    Private WithEvents TheLastTransactionBindingSource As BindingSource

    Private addingBoolean As Boolean = False
    Private editingBoolean As Boolean = False
    Private previousSelectedIndex As Integer
    Private LogDateTime As Date = Date.Now
    Private mstrLogDate As String

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click

        'This will close the form
        Me.Close()

    End Sub

    Private Sub SetComboBoxBinding()

        With cboTransactionID
            If (addingBoolean Or editingBoolean) Then
                .DataBindings!text.DataSourceUpdateMode = DataSourceUpdateMode.OnValidation
                .DropDownStyle = ComboBoxStyle.Simple
            Else
                .DataBindings!text.DataSourceUpdateMode = DataSourceUpdateMode.Never
                .DropDownStyle = ComboBoxStyle.DropDownList
            End If
        End With

    End Sub

    Private Sub LastTransaction_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Setting local variables
        Dim intCounter As Integer
        Dim intNumberOfRecords As Integer
        Dim intSelectedIndex As Integer
        Dim intEmployeeIDForSearch As Integer
        Dim intEmployeeIDFromTable As Integer
        Dim blnEmployeeNotFound As Boolean = True

        'This will bind the controls during load up

        'Try Catch
        Try

            'Binding the data controls
            TheLastTransactionDataTier = New LastTransactionDataTier
            TheLastTransactionDataSet = TheLastTransactionDataTier.GetLastTransactionInformation
            TheLastTransactionBindingSource = New BindingSource

            'Setting up binding source
            With TheLastTransactionBindingSource
                .DataSource = TheLastTransactionDataSet
                .DataMember = "lasttransaction"
                .MoveFirst()
                .MoveLast()
            End With

            'binding the combo box
            With cboTransactionID
                .DataSource = TheLastTransactionBindingSource
                .DisplayMember = "TransactionID"
                .DataBindings.Add("text", TheLastTransactionBindingSource, "TransactionID", False, DataSourceUpdateMode.Never)
            End With

            'Binding the rest of the controls
            txtEmployeeID.DataBindings.Add("text", TheLastTransactionBindingSource, "EmployeeID")
            txtLastTransactionSummary.DataBindings.Add("text", TheLastTransactionBindingSource, "LastTransactionSummary")
            txtDate.DataBindings.Add("text", TheLastTransactionBindingSource, "Date")

            'Setting up for the loop
            intNumberOfRecords = cboTransactionID.Items.Count - 1
            intEmployeeIDForSearch = CInt(Logon.mintEmployeeID)

            'Begining loop
            For intCounter = 0 To intNumberOfRecords

                'incrementing the combo box
                cboTransactionID.SelectedIndex = intCounter

                'Getting employee id
                intEmployeeIDFromTable = CInt(txtEmployeeID.Text)

                If intEmployeeIDForSearch = intEmployeeIDFromTable Then

                    'Setting the selected index
                    intSelectedIndex = intCounter
                    blnEmployeeNotFound = False

                End If
            Next

            cboTransactionID.Visible = True

            'Beginning the process of creating a new record
            With TheLastTransactionBindingSource
                .EndEdit()
                .AddNew()
            End With

            'Setting up variables
            addingBoolean = True
            SetComboBoxBinding()
            cboTransactionID.Focus()

            'Setting the combo box
            If cboTransactionID.SelectedIndex <> -1 Then
                previousSelectedIndex = cboTransactionID.Items.Count - 1
            Else
                previousSelectedIndex = 0
            End If

            'Loading up the controls
            CreateLastTransactionID.Show()
            cboTransactionID.Text = CStr(Logon.mintCreatedTransactionID)
            txtEmployeeID.Text = CStr(Logon.mintEmployeeID)
            mstrLogDate = CStr(LogDateTime)
            txtDate.Text = mstrLogDate
            txtLastTransactionSummary.Text = Logon.mstrLastTransactionSummary

            SaveLastTransaction()
            Me.Close()


        Catch ex As Exception

            'Message to user if there is a problem
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub
    Private Sub SaveLastTransaction()

        'Try Catch to catch the exceptions
        Try

            'Saving record
            TheLastTransactionBindingSource.EndEdit()
            TheLastTransactionDataTier.UpdateLastTransactionDB(TheLastTransactionDataSet)
            addingBoolean = False
            editingBoolean = False
            SetComboBoxBinding()
            cboTransactionID.Visible = False

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class