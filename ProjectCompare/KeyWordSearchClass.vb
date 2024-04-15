'Title:         Key Word Search Class
'Date:          1-27-15
'Author:        Terry Holmes

'Description:   This class is used to find the a keyword

Option Strict On

Public Class KeyWordSearchClass

    'Creating Class Function
    Public Function FindKeyWord(ByVal strKeywordForSearch As String, ByVal strKeywordFromTable As String) As Boolean

        'Setting Class Local Variables
        Dim blnItemNotFound As Boolean = True
        Dim intKeywordLengthFromTable As Integer = 0
        Dim chaKeywordInputArray() As Char
        Dim intKeywordCounter As Integer = 0
        Dim intKeywordLengthForSearch As Integer = 0
        Dim intCharacterCounter As Integer = 0
        Dim intCounter As Integer = 0
        Dim strTempString As String = ""

        'Setting up the character length variables
        intKeywordLengthForSearch = strKeywordForSearch.Length - 1
        intKeywordLengthFromTable = strKeywordFromTable.Length - 1
        intCharacterCounter = intKeywordLengthForSearch

        'Loading up the Character Array
        chaKeywordInputArray = strKeywordFromTable.ToCharArray

        'If statement to see if the table length is larger or equal to
        If intKeywordLengthForSearch <= intKeywordLengthFromTable Then

            'Loop to begin the process
            For intCounter = 0 To intKeywordLengthFromTable

                'This will clear the string
                strTempString = ""

                'Second Loop for the character counter
                For intKeywordCounter = intCounter To intCharacterCounter

                    'Creating the string
                    strTempString = strTempString + CStr(chaKeywordInputArray(intKeywordCounter))

                Next

                'Setting up for the next loop
                If intCharacterCounter < intKeywordLengthFromTable Then
                    intCharacterCounter = intCharacterCounter + 1
                End If

                'Comparing the completed string
                If strTempString = strKeywordForSearch Then
                    blnItemNotFound = False
                End If
            Next
        Else

            'Checking to see if two match
            If strKeywordForSearch = strKeywordFromTable Then
                blnItemNotFound = False
            End If

        End If

        'Returning value to call form
        Return blnItemNotFound

    End Function
End Class
