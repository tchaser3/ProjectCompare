'Title:         Project Compare and Update
'Date:          3-2-15
'Author:        Terry Holmes

'Description:   This form will update the projects and add them if needed

Option Strict On

Public Class Form1

    'Setting up the global variables
    Private TheInternalProjectsDataSet As InternalProjectsDataSet
    Private TheInternalProjectsDataTier As InternalProjectsDataTier
    Private WithEvents TheInternalProjectsBindingSource As BindingSource

    Private TheReceivePartsDataSet As ReceivePartsDataSet
    Private TheReceivePartsDataTier As ReceivePartsDataTier
    Private WithEvents TheReceivePartsBindingSource As BindingSource

    Private TheBOMPartsDataSet As BOMPartsDataSet
    Private TheBOMPartsDataTier As BOMPartsDataTier
    Private WithEvents TheBOMPartsBindingSource As BindingSource

    Private ThePartNumberDataSet As PartNumberDataSet
    Private ThePartNumberDataTier As PartNumberDataTier
    Private WithEvents ThePartNumberBindingSource As BindingSource

    'Setting other global variables
    Private addingBoolean As Boolean = False
    Private editingBoolean As Boolean = False
    Private previousSelectedIndex As Integer

    'Setting up the arrays
    Dim mstrProjectID() As String
    Dim mintProjectCounter As Integer
    Dim mintProjectUpperLimit As Integer

    Dim mstrNewProjectID() As String
    Dim mintNewProjectCounter As Integer
    Dim mintNewProjectUpperLimit As Integer

    Dim mstrPartNumberNotFound() As String
    Dim mintPartNotFoundUpperLimit As Integer
    Dim mintPartNotFoundCounter As Integer

    Dim mstrPartNumberChanged() As String
    Dim mintPartCounter As Integer
    Dim mintPartUpperLimit As Integer

    Dim TheInputDataValidation As New InputDataValidation
    Dim TheKeywordSearch As New KeyWordSearchClass
    Dim mintNewPrintCounter As Integer
    Dim LogDate As Date = Date.Now
    Dim mstrDate As String

    Structure ReceiveInventory
        Dim mintSelectedIndex As Integer
        Dim mstrProjectID As String
        Dim mstrPartNumber As String
    End Structure

    Dim structReceiveInventory() As ReceiveInventory

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click

        'Close the program
        CloseProgram.ShowDialog()

    End Sub
    Private Sub ClearTransactionDataBindings()

        'This will clear the data bindings from the transaction controls
        cboTransactionID.DataBindings.Clear()
        txtTableProjectID.DataBindings.Clear()
        txtTablePartNumber.DataBindings.Clear()

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Setting local variables
        Dim intNumberOfRecords As Integer
        Dim intCounter As Integer

        'This will bind the controls
        Try

            'Setting up the data variables
            TheInternalProjectsDataTier = New InternalProjectsDataTier
            TheInternalProjectsDataSet = TheInternalProjectsDataTier.GetInternalProjectsInformation
            TheInternalProjectsBindingSource = New BindingSource

            'Setting up the part data variable
            ThePartNumberDataTier = New PartNumberDataTier
            ThePartNumberDataSet = ThePartNumberDataTier.GetPartNumberInformation
            ThePartNumberBindingSource = New BindingSource

            'Setting up the binding source
            With TheInternalProjectsBindingSource
                .DataSource = TheInternalProjectsDataSet
                .DataMember = "internalprojects"
                .MoveFirst()
                .MoveLast()
            End With

            With ThePartNumberBindingSource
                .DataSource = ThePartNumberDataSet
                .DataMember = "partnumbers"
                .MoveFirst()
                .MoveLast()
            End With

            'Setting up the combo box
            With cboInternalProjectID
                .DataSource = TheInternalProjectsBindingSource
                .DisplayMember = "internalProjectID"
                .DataBindings.Add("text", TheInternalProjectsBindingSource, "internalProjectID", False, DataSourceUpdateMode.Never)
            End With

            With cboPartID
                .DataSource = ThePartNumberBindingSource
                .DisplayMember = "PartID"
                .DataBindings.Add("text", ThePartNumberBindingSource, "PartID", False, DataSourceUpdateMode.Never)
            End With

            'Setting up the rest of the controls
            txtTWCProjectID.DataBindings.Add("text", TheInternalProjectsBindingSource, "TWCControlNumber")

            txtPartDescription.DataBindings.Add("text", ThePartNumberBindingSource, "Description")
            txtPartNumber.DataBindings.Add("text", ThePartNumberBindingSource, "PartNumber")

            SetPartControlsVisible(False)

            'Setting up to copy the array
            intNumberOfRecords = cboInternalProjectID.Items.Count - 1
            ReDim mstrProjectID(intNumberOfRecords)
            mintProjectUpperLimit = intNumberOfRecords
            mintProjectCounter = 0

            'running the array
            For intCounter = 0 To intNumberOfRecords

                'Incrementing the combo box
                cboInternalProjectID.SelectedIndex = intCounter

                'Loading the Project ID into array
                mstrProjectID(intCounter) = txtTWCProjectID.Text
            Next

            'This will clear the transactional data bindings
            ClearTransactionDataBindings()
            SetControlsVisible(False)

            If PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                PartNumberDocument.PrinterSettings = PrintDialog1.PrinterSettings
                ProjectIDDocument.PrinterSettings = PrintDialog1.PrinterSettings
                PartNumberNotFoundDocument.PrinterSettings = PrintDialog1.PrinterSettings
            End If

            LoadReceivedDataBindings()

            intNumberOfRecords = cboTransactionID.Items.Count - 1
            ReDim structReceiveInventory(intNumberOfRecords)

            'running loop
            For intCounter = 0 To intNumberOfRecords

                'Setting up the combo box
                cboTransactionID.SelectedIndex = intCounter

                'loading up the structure
                structReceiveInventory(intCounter).mintSelectedIndex = cboTransactionID.SelectedIndex
                structReceiveInventory(intCounter).mstrPartNumber = txtTablePartNumber.Text
                structReceiveInventory(intCounter).mstrProjectID = txtTableProjectID.Text

            Next

        Catch ex As Exception

            'Message to user
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub
    Private Sub ClearProjectDataBindings()

        'This will clear the main data bindings
        cboInternalProjectID.DataBindings.Clear()
        txtTWCProjectID.DataBindings.Clear()

    End Sub
    Private Sub LoadReceivedDataBindings()

        'This will load the received databindings
        Try

            'Setting up the bindings
            TheReceivePartsDataTier = New ReceivePartsDataTier
            TheReceivePartsDataSet = TheReceivePartsDataTier.GetReceivePartsInformation
            TheReceivePartsBindingSource = New BindingSource

            'Setting up the binding source
            With TheReceivePartsBindingSource
                .DataSource = TheReceivePartsDataSet
                .DataMember = "ReceivedParts"
                .MoveFirst()
                .MoveLast()
            End With

            'Setting up the combo box
            With cboTransactionID
                .DataSource = TheReceivePartsBindingSource
                .DisplayMember = "TransactionID"
                .DataBindings.Add("text", TheReceivePartsBindingSource, "TransactionID", False, DataSourceUpdateMode.Never)
            End With

            txtTableProjectID.DataBindings.Add("text", TheReceivePartsBindingSource, "ProjectID")
            txtTablePartNumber.DataBindings.Add("text", TheReceivePartsBindingSource, "PartNumber")

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub LoadBOMPartsDataBindings()


        'This will load the received databindings
        Try

            'Setting up the bindings
            TheBOMPartsDataTier = New BOMPartsDataTier
            TheBOMPartsDataSet = TheBOMPartsDataTier.GetBOMPartsInformation
            TheBOMPartsBindingSource = New BindingSource

            'Setting up the binding source
            With TheBOMPartsBindingSource
                .DataSource = TheBOMPartsDataSet
                .DataMember = "BOMParts"
                .MoveFirst()
                .MoveLast()
            End With

            'Setting up the combo box
            With cboTransactionID
                .DataSource = TheBOMPartsBindingSource
                .DisplayMember = "TransactionID"
                .DataBindings.Add("text", TheBOMPartsBindingSource, "TransactionID", False, DataSourceUpdateMode.Never)
            End With

            txtTableProjectID.DataBindings.Add("text", TheBOMPartsBindingSource, "ProjectID")

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub SetComboBoxBinding()
        With cboInternalProjectID
            If (addingBoolean Or editingBoolean) Then
                .DataBindings!text.DataSourceUpdateMode = DataSourceUpdateMode.OnValidation
                .DropDownStyle = ComboBoxStyle.Simple
            Else
                .DataBindings!text.DataSourceUpdateMode = DataSourceUpdateMode.Never
                .DropDownStyle = ComboBoxStyle.DropDownList
            End If
        End With
    End Sub
    Private Sub CompareProjectID()

        'setting local variables
        Dim intMainCounter As Integer
        Dim intNewProjectArray As Integer
        Dim intTransactionCounter As Integer
        Dim intTransactionNumberOfRecords As Integer
        Dim strProjectIDForSearch As String
        Dim strProjectIDFromTable As String
        Dim blnItemNotFound As Boolean

        'Setting up for the loop
        intTransactionNumberOfRecords = cboTransactionID.Items.Count - 1

        'Performing Loop
        For intTransactionCounter = 0 To intTransactionNumberOfRecords

            'Setting up the boolean modifier
            blnItemNotFound = True

            'incrementing the combo box
            cboTransactionID.SelectedIndex = intTransactionCounter

            'Getting the project ID
            strProjectIDForSearch = txtTableProjectID.Text

            'performing the second loop
            For intMainCounter = 0 To mintProjectUpperLimit

                'getting part number
                strProjectIDFromTable = mstrProjectID(intMainCounter)

                'If statements
                If strProjectIDForSearch = strProjectIDFromTable Then
                    blnItemNotFound = False
                End If
            Next

            'third loop
            For intNewProjectArray = 0 To mintNewProjectUpperLimit

                'Getting part number
                strProjectIDFromTable = mstrNewProjectID(intNewProjectArray)

                'If statements
                If strProjectIDForSearch = strProjectIDFromTable Then
                    blnItemNotFound = False
                End If
            Next

            'Setting up for the item to be saved
            If blnItemNotFound = True Then
                mstrNewProjectID(mintNewProjectCounter) = strProjectIDForSearch
                mintNewProjectCounter += 1
                mintNewProjectUpperLimit += 1
                CreateNewProjectID()
            End If
        Next

    End Sub
    Private Sub CreateNewProjectID()

        'This will create a new Project
        Try

            With TheInternalProjectsBindingSource
                .EndEdit()
                .AddNew()
            End With

            'Setting up the variables
            addingBoolean = True
            SetComboBoxBinding()
            CreateProjectID.Show()
            cboInternalProjectID.Text = CStr(Logon.mintCreatedTransactionID)
            txtTWCProjectID.Text = txtTableProjectID.Text

            'saving the record
            TheInternalProjectsBindingSource.EndEdit()
            TheInternalProjectsDataTier.UpdateInternalProjectsDB(TheInternalProjectsDataSet)
            addingBoolean = False
            editingBoolean = False
            SetComboBoxBinding()

        Catch ex As Exception

            'Message to user if failure
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnProcess_Click(sender As Object, e As EventArgs) Handles btnProcess.Click

        'This will process the transactions
        'Setting local variables
        Dim intNumberOfRecords As Integer

        mintNewProjectCounter = 0
        mintNewProjectUpperLimit = 0
        SetControlsVisible(True)
        ClearTransactionDataBindings()
        LoadReceivedDataBindings()
        intNumberOfRecords = cboTransactionID.Items.Count - 1
        ReDim mstrNewProjectID(intNumberOfRecords)
        CompareProjectID()
        ClearTransactionDataBindings()
        LoadBOMPartsDataBindings()
        CompareProjectID()
        ClearTransactionDataBindings()
        mintNewPrintCounter = 0
        mintNewProjectUpperLimit -= 1
        ProjectIDDocument.Print()
        SetControlsVisible(False)
        MessageBox.Show("Process Is Complete", "Thank You", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub
    Private Sub SetControlsVisible(ByVal ValueBoolean As Boolean)

        cboInternalProjectID.Visible = ValueBoolean
        cboTransactionID.Visible = ValueBoolean
        txtTableProjectID.Visible = ValueBoolean
        txtTWCProjectID.Visible = ValueBoolean
        txtTablePartNumber.Visible = ValueBoolean

    End Sub
    Private Sub EditSaveBOMTransaction()

        Try

            'This will change the project id
            CheckPartNumber()
            editingBoolean = True
            SetComboBoxBinding()
            txtTableProjectID.Text = "WHS"
            TheBOMPartsBindingSource.EndEdit()
            TheBOMPartsDataTier.UpdateBOMPartsDB(TheBOMPartsDataSet)
            editingBoolean = False
            SetComboBoxBinding()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub
    Private Sub EditSaveReceiveTransaction()

        Try

            'This will change the project id
            CheckPartNumber()
            editingBoolean = True
            SetComboBoxBinding()
            txtTableProjectID.Text = "WHS"
            TheReceivePartsBindingSource.EndEdit()
            TheReceivePartsDataTier.UpdateReceivePartsDB(TheReceivePartsDataSet)
            editingBoolean = False
            SetComboBoxBinding()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub

    Private Sub btnConvert_Click(sender As Object, e As EventArgs) Handles btnConvert.Click

        'This will change all the Decimal project id to WHS project

        'Setting local variables
        Dim intCounter As Integer
        Dim intNumberOfRecords As Integer
        Dim strValueForValidation As String
        Dim decInputValue As Decimal
        Dim decInputValueTruncated As Decimal
        Dim blnKeywordFailed As Boolean
        Dim strKeywordForSearch As String
        Dim strKeywordFromTable As String
        Dim intNewNumber As Integer

        'Setting the bindings for the receive table
        SetControlsVisible(True)
        ClearTransactionDataBindings()
        LoadReceivedDataBindings()

        'Setting up for the loop
        intNumberOfRecords = cboTransactionID.Items.Count - 1
        ReDim mstrPartNumberChanged(intNumberOfRecords)
        ReDim mstrPartNumberNotFound(intNumberOfRecords)
        mintPartNotFoundUpperLimit = 0
        mintPartUpperLimit = 0
        mintPartNotFoundCounter = 0

        'Performing Loop
        For intCounter = 0 To intNumberOfRecords


            'Loading up the variable
            strValueForValidation = structReceiveInventory(intCounter).mstrProjectID

            If IsNumeric(strValueForValidation) Then

                'Setting up the decimal variables
                decInputValue = Convert.ToDecimal(strValueForValidation)
                decInputValueTruncated = Decimal.Truncate(decInputValue)
                intNewNumber = CInt(strValueForValidation)

                If intNewNumber < 80000 Then
                    cboTransactionID.SelectedIndex = structReceiveInventory(intCounter).mintSelectedIndex
                    EditSaveReceiveTransaction()
                End If
                If intNewNumber > 130000 Then
                    cboTransactionID.SelectedIndex = structReceiveInventory(intCounter).mintSelectedIndex
                    EditSaveReceiveTransaction()
                End If
                If decInputValue <> decInputValueTruncated Then

                    'This will change the project name
                    cboTransactionID.SelectedIndex = structReceiveInventory(intCounter).mintSelectedIndex
                    EditSaveReceiveTransaction()

                End If
            Else

                'Setting up for keyword search
                strKeywordFromTable = strValueForValidation

                'beginning the first validation
                strKeywordForSearch = "ICK"
                blnKeywordFailed = TheKeywordSearch.FindKeyWord(strKeywordForSearch, strKeywordFromTable)

                'Checking the second variable
                If blnKeywordFailed = True Then
                    strKeywordForSearch = "299-133"
                    blnKeywordFailed = TheKeywordSearch.FindKeyWord(strKeywordForSearch, strKeywordFromTable)
                End If

                If blnKeywordFailed = True Then
                    strKeywordForSearch = "DB"
                    blnKeywordFailed = TheKeywordSearch.FindKeyWord(strKeywordForSearch, strKeywordFromTable)
                End If

                If blnKeywordFailed = True Then
                    strKeywordForSearch = "TECH OPS"
                    blnKeywordFailed = TheKeywordSearch.FindKeyWord(strKeywordForSearch, strKeywordFromTable)
                End If

                If blnKeywordFailed = True Then
                    strKeywordForSearch = "NO DID"
                    blnKeywordFailed = TheKeywordSearch.FindKeyWord(strKeywordForSearch, strKeywordFromTable)
                End If

                If blnKeywordFailed = True Then
                    strKeywordForSearch = "MAIN"
                    blnKeywordFailed = TheKeywordSearch.FindKeyWord(strKeywordForSearch, strKeywordFromTable)
                End If

                If blnKeywordFailed = True Then
                    strKeywordForSearch = "YARD"
                    blnKeywordFailed = TheKeywordSearch.FindKeyWord(strKeywordForSearch, strKeywordFromTable)
                End If

                If blnKeywordFailed = True Then
                    strKeywordForSearch = "ADF"
                    blnKeywordFailed = TheKeywordSearch.FindKeyWord(strKeywordForSearch, strKeywordFromTable)
                End If

                If blnKeywordFailed = True Then
                    strKeywordForSearch = "REPLACEMENT"
                    blnKeywordFailed = TheKeywordSearch.FindKeyWord(strKeywordForSearch, strKeywordFromTable)
                End If

                'changing the information
                If blnKeywordFailed = False Then
                    cboTransactionID.SelectedIndex = structReceiveInventory(intCounter).mintSelectedIndex
                    EditSaveReceiveTransaction()
                End If

            End If
        Next

        SetControlsVisible(False)
        SetPartControlsVisible(True)

        Logon.mstrLastTransactionSummary = "RAN ON HAND INVENTORY REPORT"

        'Setting up for multiple pages
        mintNewPrintCounter = 0
        PartNumberDocument.Print()
        mintNewPrintCounter = 0
        mintPartNotFoundUpperLimit = mintPartNotFoundCounter - 1
        PartNumberNotFoundDocument.Print()
        SetPartControlsVisible(False)
        MessageBox.Show("Process Is Complete", "Thank You", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub
    Private Sub CheckPartNumber()

        'this sub routine will check to see if the part number is in array

        'Setting local variables
        Dim intCounter As Integer
        Dim strPartNumberFromTable As String
        Dim strPartNumberForSearch As String
        Dim blnItemNotFound As Boolean = True

        'Setting up for loop
        strPartNumberForSearch = txtTablePartNumber.Text

        'running loop
        For intCounter = 0 To mintPartUpperLimit

            'Getting part number
            strPartNumberFromTable = mstrPartNumberChanged(intCounter)

            If strPartNumberForSearch = strPartNumberFromTable Then
                blnItemNotFound = False
            End If

        Next

        If blnItemNotFound = True Then
            mintPartUpperLimit += 1
            mstrPartNumberChanged(mintPartUpperLimit) = strPartNumberForSearch
        End If


    End Sub
    Private Function FindPartNumber(ByVal strPartNumberForSearch As String) As Boolean

        'Setting local variables
        Dim intCounter As Integer
        Dim intNumberOfRecords As Integer
        Dim strPartNumberFromTable As String
        Dim intSelectedIndex As Integer
        Dim blnItemNotFound As Boolean = True

        'Setting up for loop
        intNumberOfRecords = cboPartID.Items.Count - 1

        'Performing loop
        For intCounter = 0 To intNumberOfRecords

            'incrementing the combo box
            cboPartID.SelectedIndex = intCounter

            'Getting part number
            strPartNumberFromTable = txtPartNumber.Text

            'setting the selected index
            If strPartNumberForSearch = strPartNumberFromTable Then
                intSelectedIndex = intCounter
                blnItemNotFound = False
            End If
        Next

        'Setting the combo box to right location
        If blnItemNotFound = False Then

            cboPartID.SelectedIndex = intSelectedIndex

        End If

        Return blnItemNotFound

    End Function
    Private Sub SetPartControlsVisible(ByVal ValueBoolean As Boolean)

        cboPartID.Visible = ValueBoolean
        txtPartDescription.Visible = ValueBoolean
        txtPartNumber.Visible = ValueBoolean

    End Sub

    Private Sub PartNumberDocument_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PartNumberDocument.PrintPage

        'This will print the document

        'Setting local variables
        Dim intCounter As Integer
        Dim blnItemNotFound As Boolean
        Dim intStartingPageCounter As Integer
        Dim blnNewPage As Boolean = False

        'Setting up variables for the reports
        Dim PrintHeaderFont As New Font("Arial", 18, FontStyle.Bold)
        Dim PrintSubHeaderFont As New Font("Arial", 14, FontStyle.Bold)
        Dim PrintItemsFont As New Font("Arial", 10, FontStyle.Regular)
        Dim PrintX As Single = e.MarginBounds.Left
        Dim PrintY As Single = e.MarginBounds.Top
        Dim HeadingLineHeight As Single = PrintHeaderFont.GetHeight + 18
        Dim ItemLineHeight As Single = PrintItemsFont.GetHeight + 10

        'Getting the date
        mstrDate = CStr(LogDate)

        'Setting up for default position
        PrintY = 100

        'Setting up for the print header
        PrintX = 150
        e.Graphics.DrawString("Blue Jay Communications Inventory Report", PrintHeaderFont, Brushes.Black, PrintX, PrintY)
        PrintY = PrintY + HeadingLineHeight
        PrintX = 162
        e.Graphics.DrawString("Part Numbers To Counted:  " + mstrDate, PrintSubHeaderFont, Brushes.Black, PrintX, PrintY)
        PrintY = PrintY + HeadingLineHeight

        'Setting up the columns
        PrintX = 100
        e.Graphics.DrawString("Part Number", PrintItemsFont, Brushes.Black, PrintX, PrintY)
        PrintX = 250
        e.Graphics.DrawString("Part Description", PrintItemsFont, Brushes.Black, PrintX, PrintY)
        PrintY = PrintY + HeadingLineHeight

        'Performing Loop
        For intCounter = mintNewPrintCounter To mintPartUpperLimit

            blnItemNotFound = FindPartNumber(mstrPartNumberChanged(intCounter))

            If blnItemNotFound = False Then

                PrintX = 100
                e.Graphics.DrawString(txtPartNumber.Text, PrintItemsFont, Brushes.Black, PrintX, PrintY)
                PrintX = 250
                e.Graphics.DrawString(txtPartDescription.Text, PrintItemsFont, Brushes.Black, PrintX, PrintY)
                PrintY = PrintY + ItemLineHeight

            Else

                mstrPartNumberNotFound(mintPartNotFoundCounter) = mstrPartNumberChanged(intCounter)
                mintPartNotFoundCounter += 1

            End If

            If PrintY > 900 Then
                If intStartingPageCounter = mintPartUpperLimit Then
                    e.HasMorePages = False
                Else
                    e.HasMorePages = True
                    blnNewPage = True
                End If
            End If

            If blnNewPage = True Then
                mintNewPrintCounter = intCounter + 1
                intCounter = mintPartUpperLimit
            End If

        Next

    End Sub

    Private Sub PartNumberNotFoundDocument_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PartNumberNotFoundDocument.PrintPage

        'This will print the document

        'Setting local variables
        Dim intCounter As Integer
        Dim intStartingPageCounter As Integer
        Dim blnNewPage As Boolean = False

        'Setting up variables for the reports
        Dim PrintHeaderFont As New Font("Arial", 18, FontStyle.Bold)
        Dim PrintSubHeaderFont As New Font("Arial", 14, FontStyle.Bold)
        Dim PrintItemsFont As New Font("Arial", 10, FontStyle.Regular)
        Dim PrintX As Single = e.MarginBounds.Left
        Dim PrintY As Single = e.MarginBounds.Top
        Dim HeadingLineHeight As Single = PrintHeaderFont.GetHeight + 18
        Dim ItemLineHeight As Single = PrintItemsFont.GetHeight + 10

        'Getting the date
        mstrDate = CStr(LogDate)

        'Setting up for default position
        PrintY = 100

        'Setting up for the print header
        PrintX = 150
        e.Graphics.DrawString("Blue Jay Communications Inventory Report", PrintHeaderFont, Brushes.Black, PrintX, PrintY)
        PrintY = PrintY + HeadingLineHeight
        PrintX = 162
        e.Graphics.DrawString("Part Numbers Not Found:  " + mstrDate, PrintSubHeaderFont, Brushes.Black, PrintX, PrintY)
        PrintY = PrintY + HeadingLineHeight

        'Setting up the columns
        PrintX = 100
        e.Graphics.DrawString("Part Number", PrintItemsFont, Brushes.Black, PrintX, PrintY)
        PrintY = PrintY + HeadingLineHeight

        'Performing Loop
        For intCounter = mintNewPrintCounter To mintPartNotFoundUpperLimit


            PrintX = 100
            e.Graphics.DrawString(mstrPartNumberNotFound(intCounter), PrintItemsFont, Brushes.Black, PrintX, PrintY)
            PrintY = PrintY + ItemLineHeight

            If PrintY > 900 Then
                If intStartingPageCounter = mintPartNotFoundUpperLimit Then
                    e.HasMorePages = False
                Else
                    e.HasMorePages = True
                    blnNewPage = True
                End If
            End If

            If blnNewPage = True Then
                mintNewPrintCounter = intCounter + 1
                intCounter = mintPartNotFoundUpperLimit
            End If

        Next

    End Sub

    Private Sub ProjectIDDocument_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles ProjectIDDocument.PrintPage

        'This Subroutine is used to print the projects that have been created
        'Setting local variables
        Dim intCounter As Integer
        Dim intStartingPageCounter As Integer
        Dim blnNewPage As Boolean = False

        'Setting up variables for the reports
        Dim PrintHeaderFont As New Font("Arial", 18, FontStyle.Bold)
        Dim PrintSubHeaderFont As New Font("Arial", 14, FontStyle.Bold)
        Dim PrintItemsFont As New Font("Arial", 10, FontStyle.Regular)
        Dim PrintX As Single = e.MarginBounds.Left
        Dim PrintY As Single = e.MarginBounds.Top
        Dim HeadingLineHeight As Single = PrintHeaderFont.GetHeight + 18
        Dim ItemLineHeight As Single = PrintItemsFont.GetHeight + 10

        'Getting the date
        mstrDate = CStr(LogDate)

        'Setting up for default position
        PrintY = 100

        'Setting up for the print header
        PrintX = 150
        e.Graphics.DrawString("Blue Jay Communications Inventory Report", PrintHeaderFont, Brushes.Black, PrintX, PrintY)
        PrintY = PrintY + HeadingLineHeight
        PrintX = 162
        e.Graphics.DrawString("Time Warner Projects Added:  " + mstrDate, PrintSubHeaderFont, Brushes.Black, PrintX, PrintY)
        PrintY = PrintY + HeadingLineHeight

        'Setting up the columns
        PrintX = 100
        e.Graphics.DrawString("ProjectID", PrintItemsFont, Brushes.Black, PrintX, PrintY)
        PrintY = PrintY + HeadingLineHeight

        'Performing Loop
        For intCounter = mintNewPrintCounter To mintNewProjectUpperLimit


            PrintX = 100
            e.Graphics.DrawString(mstrNewProjectID(intCounter), PrintItemsFont, Brushes.Black, PrintX, PrintY)
            PrintY = PrintY + ItemLineHeight

            If PrintY > 900 Then
                If intStartingPageCounter = mintNewProjectUpperLimit Then
                    e.HasMorePages = False
                Else
                    e.HasMorePages = True
                    blnNewPage = True
                End If
            End If

            If blnNewPage = True Then
                mintNewPrintCounter = intCounter + 1
                intCounter = mintNewProjectUpperLimit
            End If

        Next

    End Sub
End Class
