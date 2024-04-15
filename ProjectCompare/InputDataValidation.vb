'Title:         Input Data Validation
'Date:          7/23/13
'Author:        Terry Holmes

'Description:   This Class performs the Data Validation replacing on the form validation

Option Strict On

Public Class InputDataValidation

    Public Function VerifyData(ByVal strValueInput As String) As Boolean

        'Setting Local Variables
        Dim blnFatalError As Boolean = False
        Dim decInputValue As Decimal
        Dim decInputValueTruncated As Decimal
        Dim strErrorMessage As String

        strErrorMessage = ""
        Logon.mstrErrorMessage = ""

        'Perparing For Data Validation

        If Not IsNumeric(strValueInput) Then
            blnFatalError = True
            strErrorMessage = strErrorMessage + "The Value entered is Not Numeric" + vbNewLine
        Else

            decInputValue = Convert.ToDecimal(strValueInput)
            decInputValueTruncated = Decimal.Truncate(decInputValue)

            If decInputValue <> decInputValueTruncated Then
                blnFatalError = True
                strErrorMessage = strErrorMessage + "The Value Entered is not An Integer" + vbNewLine
            ElseIf decInputValue < 0 Then
                blnFatalError = True
                strErrorMessage = strErrorMessage + "The Value Entered is Less Than 1" + vbNewLine
            ElseIf decInputValue > 999999999 Then
                blnFatalError = True
                strErrorMessage = strErrorMessage + "The Value Entered is Too Large" + vbNewLine
            End If
        End If

        If blnFatalError = True Then
            Logon.mstrErrorMessage = strErrorMessage
        End If

        'Returning value to calling sub-routing
        Return blnFatalError

    End Function
    Public Function VerifyYesNoData(ByVal strValueInput As String) As Boolean

        Dim blnFatalError As Boolean = False
        Dim strErrorMessage As String = ""

        If strValueInput = "" Then
            blnFatalError = True
            strErrorMessage = strErrorMessage + "The Value Entered is Less Blank" + vbNewLine
        Else
            If strValueInput = "YES" Or strValueInput = "NO" Then

            Else
                blnFatalError = True
                strErrorMessage = strErrorMessage + "You can only put either a YES or NO for this Value" + vbNewLine
            End If
        End If

        If blnFatalError = True Then
            Logon.mstrErrorMessage = strErrorMessage
        End If

        'Returning value to calling sub-routing
        Return blnFatalError

    End Function
    Public Function VerifyTextData(ByVal strValueInput As String) As Boolean

        Dim blnFatalError As Boolean = False
        Dim strErrorMessage As String = ""

        If strValueInput = "" Then
            blnFatalError = True
            strErrorMessage = strErrorMessage + "The Value Entered is Less Blank" + vbNewLine
        End If

        If blnFatalError = True Then
            Logon.mstrErrorMessage = strErrorMessage
        End If

        'Returning value to calling sub-routing
        Return blnFatalError

    End Function

    Public Function VerifyStartingEndingDatesAsDates(ByVal strValueInput As String) As Boolean

        Dim blnIsValidDate As Boolean = False

        If Not IsDate(strValueInput) Then
            blnIsValidDate = True
            Logon.mstrErrorMessage = "The information Entered is Not a Date"
        End If

        Return blnIsValidDate

    End Function

    Public Function VerifyStartingEndingDates(ByVal datValueInput As Date) As Boolean

        Dim blnIsValidDate As Boolean = False
        Dim datTestingDate As Date = CDate("06/01/2013")

        If datValueInput < datTestingDate Then

            blnIsValidDate = True
            Logon.mstrErrorMessage = "The Date Entered is Outside of the Range"
        End If

        Return blnIsValidDate

    End Function
    Public Function VerifyMoney(ByVal strValueInput As String) As Boolean

        Dim blnFatalError As Boolean = False
        Dim strErrorMessage As String = ""
        Dim decValueInput As Decimal

        If Not IsNumeric(strValueInput) Then
            blnFatalError = True
            strErrorMessage = "The Value Entered is not Numeric"
        Else
            decValueInput = CDec(strValueInput)
            If decValueInput < 0.0 Then
                blnFatalError = True
                strErrorMessage = "The Value is Less Than 0.00"
            End If
        End If

        Logon.mstrErrorMessage = strErrorMessage

        Return blnFatalError

    End Function
    Public Function VerifyYear(ByVal strValueInput As String) As Boolean

        Dim blnFatalError As Boolean = False
        Dim strErrorMessage As String = ""
        Dim intYear As Integer

        If Not IsNumeric(strValueInput) Then
            blnFatalError = True
            strErrorMessage = "The Value Entered is not Numeric"
        Else
            intYear = CInt(strValueInput)
            If intYear < 1980 Or intYear > 2099 Then
                blnFatalError = True
                strErrorMessage = "Year Is Out of Range"
            End If
        End If

        Logon.mstrErrorMessage = strErrorMessage

        Return blnFatalError

    End Function
    Public Function VerifyVehicleNumber(ByVal strValueInput As String) As Boolean

        Dim intVehicleNumber As Integer
        Dim blnFatalError As Boolean = True

        intVehicleNumber = CInt(strValueInput)

        If intVehicleNumber > 1099 And intVehicleNumber < 1600 Then

            blnFatalError = False

        End If

        If blnFatalError = True Then
            Logon.mstrErrorMessage = "Vehicle Number Out of Range"
        End If

        Return blnFatalError


    End Function
    Public Function VerifyCableCategory(ByVal strValueInput As String) As Boolean

        'Creating Local Variable
        Dim blnCableCategoryCorrect As Boolean = False
        Dim blnFatalError As Boolean = False
        Dim strErrorMessage As String

        'If Statements to determine the correct Tool Type
        If strValueInput = "COAX" Then
            blnCableCategoryCorrect = True
        ElseIf strValueInput = "FIBER" Then
            blnCableCategoryCorrect = True
        ElseIf strValueInput = "DROP CABLE" Then
            blnCableCategoryCorrect = True
        ElseIf strValueInput = "TWISTED PAIR" Then
            blnCableCategoryCorrect = True
        ElseIf strValueInput = "SPECIALTY CABLE" Then
            blnCableCategoryCorrect = True

        End If

        If blnCableCategoryCorrect = False Then
            strErrorMessage = "The Cable Type Entered is not Correct"
            Logon.mstrErrorMessage = strErrorMessage
            blnFatalError = True
        End If

        Return blnFatalError

    End Function
    Public Function VerifyPartType(ByVal strValueInput As String) As Boolean

        'Creating Local Variable
        Dim blnPartTypeNotCorrect As Boolean = True
        Dim blnFatalError As Boolean = False
        Dim strErrorMessage As String

        'If Statements to determine the correct Tool Type
        If strValueInput = "CABLE" Then
            blnPartTypeNotCorrect = False
        ElseIf strValueInput = "BOLT" Then
            blnPartTypeNotCorrect = False
        ElseIf strValueInput = "CLAMP" Then
            blnPartTypeNotCorrect = False
        ElseIf strValueInput = "PAD" Then
            blnPartTypeNotCorrect = False
        ElseIf strValueInput = "TAP" Then
            blnPartTypeNotCorrect = False
        ElseIf strValueInput = "FIBER" Then
            blnPartTypeNotCorrect = False
        ElseIf strValueInput = "ENCLOSURE" Then
            blnPartTypeNotCorrect = False
        ElseIf strValueInput = "CONNECTOR" Then
            blnPartTypeNotCorrect = False
        ElseIf strValueInput = "EQ" Then
            blnPartTypeNotCorrect = False
        ElseIf strValueInput = "SPLICE" Then
            blnPartTypeNotCorrect = False
        End If

        If blnPartTypeNotCorrect = True Then
            strErrorMessage = "The Part Type Entered is not Correct"
            Logon.mstrErrorMessage = strErrorMessage
            blnFatalError = True
        End If

        Return blnFatalError

    End Function
    Public Function VerifyPhoneNumberFormat(ByVal strValueInput As String) As Boolean

        'Creating Local Variable
        Dim blnCharNotCorrect As Boolean = True
        Dim blnFatalError As Boolean = False
        Dim chaCharacterCheck As Char
        Dim intCounter As Integer
        Dim intStringLength As Integer

        'Making sure the information entered meets the length requirement
        intStringLength = strValueInput.Length

        If intStringLength <> 14 Then
            blnFatalError = True
        End If

        'Setting the character array
        Dim charArray() As Char = strValueInput.ToCharArray

        'For loop for to seeing if the phone number is in the correct order
        For intCounter = 0 To 13

            If blnFatalError = False Then
                blnCharNotCorrect = True
                chaCharacterCheck = charArray(intCounter)

                If intCounter = 0 Then
                    If chaCharacterCheck = "(" Then
                        blnCharNotCorrect = False
                    Else
                        blnCharNotCorrect = True
                    End If
                ElseIf intCounter = 4 Then
                    If chaCharacterCheck = ")" Then
                        blnCharNotCorrect = False
                    Else
                        blnCharNotCorrect = True
                    End If
                ElseIf intCounter = 5 Then
                    If chaCharacterCheck = " " Then
                        blnCharNotCorrect = False
                    Else
                        blnCharNotCorrect = True
                    End If
                ElseIf intCounter = 9 Then
                    If chaCharacterCheck = "-" Then
                        blnCharNotCorrect = False
                    Else
                        blnCharNotCorrect = True
                    End If
                ElseIf chaCharacterCheck = "1" Then
                    blnCharNotCorrect = False
                ElseIf chaCharacterCheck = "2" Then
                    blnCharNotCorrect = False
                ElseIf chaCharacterCheck = "3" Then
                    blnCharNotCorrect = False
                ElseIf chaCharacterCheck = "4" Then
                    blnCharNotCorrect = False
                ElseIf chaCharacterCheck = "5" Then
                    blnCharNotCorrect = False
                ElseIf chaCharacterCheck = "6" Then
                    blnCharNotCorrect = False
                ElseIf chaCharacterCheck = "7" Then
                    blnCharNotCorrect = False
                ElseIf chaCharacterCheck = "8" Then
                    blnCharNotCorrect = False
                ElseIf chaCharacterCheck = "9" Then
                    blnCharNotCorrect = False
                ElseIf chaCharacterCheck = "0" Then
                    blnCharNotCorrect = False
                End If

                'Setting up for output
                If blnCharNotCorrect = True Then
                    blnFatalError = True
                End If

            End If

        Next

        'Setting up the global variable
        Logon.mstrErrorMessage = "The Phone Number Entered Is In The Correct Format"

        'Returning the value to calling sub-routine
        Return blnFatalError

    End Function

End Class
