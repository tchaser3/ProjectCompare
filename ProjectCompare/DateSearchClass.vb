'Title:         Date Search Class
'Date:          05-20-15
'Author:        Terry Holmes

'Description:   This is the date class

Public Class DateSearchClass

    Public Function RemoveTime(ByVal StartDate As Date) As Date

        'setting local variables
        Dim intDay As Integer
        Dim intMonth As Integer
        Dim intYear As Integer
        Dim strDate As String
        Dim datReturnDate As Date

        'Getting the date integers
        intMonth = StartDate.Month
        intDay = StartDate.Day
        intYear = StartDate.Year

        'getting the string
        strDate = CStr(intMonth) + "/" + CStr(intDay) + "/" + CStr(intYear)

        'Setting the date
        datReturnDate = CDate(strDate)

        'Returning value to main
        Return datReturnDate

    End Function
    Public Function AddingDay(ByVal StartingDate As Date, ByVal intNumberOfDays As Integer) As Date

        'setting local variables
        Dim intMonth As Integer
        Dim intDay As Integer
        Dim intYear As Integer
        Dim strDate As String
        Dim datReturnDate As Date
        Dim intLeapYear As Integer
        Dim blnDateTooBig As Boolean = True

        'Getting initial values
        intMonth = StartingDate.Month
        intDay = StartingDate.Day
        intYear = StartingDate.Year

        'Incrementing the day
        intDay = intDay + intNumberOfDays

        'If Statements to determine new day
        While blnDateTooBig = True

            intLeapYear = intYear Mod 4

            If intMonth = 1 Or intMonth = 3 Or intMonth = 5 Or intMonth = 7 Or intMonth = 8 Or intMonth = 10 Or intMonth = 12 Then
                If intDay > 31 Then
                    intDay = intDay - 31
                    intMonth += 1
                    If intMonth = 13 Then
                        intMonth = 1
                        intYear += 1
                    End If
                    If intMonth = 2 Then
                        If intLeapYear = 0 Then
                            If intDay <= 29 Then
                                blnDateTooBig = False
                            End If
                        Else
                            If intDay <= 28 Then
                                blnDateTooBig = False
                            End If
                        End If
                    Else
                        If intDay <= 30 Then
                            blnDateTooBig = False
                        End If
                    End If
                ElseIf intDay <= 31 Then
                    blnDateTooBig = False
                End If
            ElseIf intMonth = 4 Or intMonth = 6 Or intMonth = 9 Or intMonth = 11 Then
                If intDay > 30 Then
                    intDay = intDay - 30
                    intMonth += 1
                    If intDay <= 31 Then
                        blnDateTooBig = False
                    End If
                ElseIf intDay <= 30 Then
                    blnDateTooBig = False
                End If
            ElseIf intMonth = 2 Then
                intLeapYear = intYear Mod 4
                If intLeapYear = 0 Then
                    If intDay > 29 Then
                        intDay = intDay - 29
                        intMonth += 1
                    ElseIf intDay <= 29 Then
                        blnDateTooBig = False
                    End If
                Else
                    If intDay > 28 Then
                        intDay = intDay - 28
                        intMonth += 1
                    ElseIf intDay <= 28 Then
                        blnDateTooBig = False
                    End If
                End If
                If intDay <= 31 Then
                    blnDateTooBig = False
                End If
            End If

        End While

        'to catch a date exception
        Try

            'getting the string
            strDate = CStr(intMonth) + "/" + CStr(intDay) + "/" + CStr(intYear)

            'Setting the date
            datReturnDate = CDate(strDate)

            'returning value to main
            Return datReturnDate

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
            datReturnDate = StartingDate
            Return datReturnDate
        End Try

    End Function
    Public Function SubtractDays(ByVal StartingDate As Date, ByVal intNumberOfDays As Integer) As Date

        'Setting local variables
        Dim intMonth As Integer
        Dim intDay As Integer
        Dim intYear As Integer
        Dim strDate As String
        Dim datReturnDate As Date
        Dim intLeapYear As Integer
        Dim blnNegativeNumber As Boolean = True

        'Getting the date
        intMonth = StartingDate.Month
        intDay = StartingDate.Day
        intYear = StartingDate.Year

        'Getting the math factor
        intDay = intDay - intNumberOfDays

        While blnNegativeNumber = True

            intLeapYear = intYear Mod 4

            'Beginning If Statements
            If intDay < 1 Then
                If intMonth = 1 Or intMonth = 3 Or intMonth = 5 Or intMonth = 7 Or intMonth = 8 Or intMonth = 10 Or intMonth = 12 Then

                    'Doing the math
                    intDay = 31 + intDay
                    intMonth -= 1
                    If intMonth < 1 Then
                        intMonth = 12
                        intYear -= 1
                    End If
                    If intMonth = 12 Then
                        If intDay > 0 And intDay < 32 Then
                            blnNegativeNumber = False
                        End If
                    ElseIf intMonth = 2 Then
                        If intDay > 0 Then
                            If intLeapYear = 0 Then
                                If intDay < 30 Then
                                    blnNegativeNumber = False
                                End If
                            Else
                                If intDay < 29 Then
                                    blnNegativeNumber = False
                                End If
                            End If
                        End If
                    ElseIf intDay > 0 And intDay < 31 Then
                        blnNegativeNumber = False
                    End If
                ElseIf intMonth = 4 Or intMonth = 6 Or intMonth = 9 Or intMonth = 11 Then

                    'doing the math
                    intDay = 30 + intDay
                    intMonth -= 1
                    If intDay > 0 And intDay < 31 Then
                        blnNegativeNumber = False
                    End If
                ElseIf intMonth = 2 Then
                    If intLeapYear = 0 Then
                        intDay = 29 + intDay
                        intMonth -= 1
                    Else
                        intDay = 28 + intDay
                        intMonth -= 1
                    End If
                    If intDay > 0 And intDay < 32 Then
                        blnNegativeNumber = False
                    End If
                End If
            ElseIf intMonth = 1 Or intMonth = 3 Or intMonth = 5 Or intMonth = 7 Or intMonth = 8 Or intMonth = 10 Or intMonth = 12 Then
                If intDay > 0 And intDay < 32 Then
                    blnNegativeNumber = False
                End If
            ElseIf intMonth = 4 Or intMonth = 6 Or intMonth = 9 Or intMonth = 11 Then
                If intDay > 0 And intDay < 31 Then
                    blnNegativeNumber = False
                End If
            ElseIf intMonth = 2 And intLeapYear = 0 Then
                If intDay > 0 And intDay < 30 Then
                    blnNegativeNumber = False
                End If
            ElseIf intMonth = 2 And intLeapYear <> 0 Then
                If intDay > 0 And intDay < 29 Then
                    blnNegativeNumber = False
                End If
            End If
        End While
        

        Try

            'getting the string
            strDate = CStr(intMonth) + "/" + CStr(intDay) + "/" + CStr(intYear)

            'Setting the date
            datReturnDate = CDate(strDate)

            'returning value to main
            Return datReturnDate

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
            datReturnDate = StartingDate
            Return datReturnDate
        End Try

    End Function
    Public Function DateDifference(ByVal datStartingDate As Date, ByVal datEndingDate As Date) As String

        Dim intStartingMonth As Integer
        Dim intStartingDay As Integer
        Dim intStartingYear As Integer
        Dim intStartingHours As Integer
        Dim intStartingMinutes As Integer
        Dim intStartingSecond As Integer
        Dim intEndingMonth As Integer
        Dim intEndingDay As Integer
        Dim intEndingYear As Integer
        Dim intEndingHours As Integer
        Dim intEndingMinutes As Integer
        Dim intEndingSecond As Integer
        Dim intTotalDays As Integer = 0
        Dim intTotalHours As Integer = 0
        Dim intTotalMinutes As Integer = 0
        Dim intTotalSeconds As Integer = 0
        Dim strReturnDateTime As String = ""
        Dim intTotalYear As Integer
        Dim intLeapYear As Integer
        Dim intTotalMonth As Integer
        Dim blnFirstRotation As Boolean
        Dim blnDateNotMatching As Boolean = True

        'Getting Starting Information
        intStartingMonth = datStartingDate.Month
        intStartingDay = datStartingDate.Day
        intStartingYear = datStartingDate.Year
        intStartingHours = datStartingDate.Hour
        intStartingMinutes = datStartingDate.Minute
        intStartingSecond = datStartingDate.Second

        'Getting ending information
        intEndingMonth = datEndingDate.Month
        intEndingDay = datEndingDate.Day
        intEndingYear = datEndingDate.Year
        intEndingHours = datEndingDate.Hour
        intEndingMinutes = datEndingDate.Minute
        intEndingSecond = datEndingDate.Second

        'Performing math for seconds
        If intStartingSecond > intEndingSecond Then
            intEndingMinutes -= 1
            intEndingSecond += 60
        End If
        intTotalSeconds = intEndingSecond - intStartingSecond

        'Performing Math for Minutes
        If intStartingMinutes > intEndingMinutes Then
            intEndingHours -= 1
            intEndingMinutes += 60
        End If
        intTotalMinutes = intEndingMinutes - intStartingMinutes

        'Performing Math for Hours
        If intStartingHours > intEndingHours Then
            intEndingDay -= 1
            intEndingHours += 24
        End If

        intTotalHours = intEndingHours - intStartingHours

        'Setting up to figure the days
        blnFirstRotation = True
        intTotalYear = intStartingYear
        intTotalMonth = intStartingMonth

        If intStartingMonth = intEndingMonth And intStartingYear = intEndingYear And intEndingDay >= intStartingDay Then
            intTotalDays = intEndingDay - intStartingDay
        Else
            'Creating the number of days
            While blnDateNotMatching = True

                If intTotalMonth = 1 Or intTotalMonth = 3 Or intTotalMonth = 5 Or intTotalMonth = 7 Or intTotalMonth = 8 Or intTotalMonth = 10 Or intTotalMonth = 12 Then
                    If blnFirstRotation = True Then
                        intTotalDays = intTotalDays + 31 - intStartingDay
                        blnFirstRotation = False
                        intTotalMonth += 1
                        If intTotalMonth = 13 Then
                            intTotalMonth = 1
                            intTotalYear += 1
                        End If
                    ElseIf intTotalMonth = intEndingMonth And intTotalYear = intEndingYear Then
                        intTotalDays += intEndingDay
                        blnDateNotMatching = False
                    Else
                        intTotalDays += 31
                        intTotalMonth += 1
                    End If
                ElseIf intTotalMonth = 4 Or intTotalMonth = 6 Or intTotalMonth = 9 Or intTotalMonth = 11 Then
                    If blnFirstRotation = True Then
                        intTotalDays = intTotalDays + 30 - intStartingDay
                        blnFirstRotation = False
                        intTotalMonth += 1
                    ElseIf intTotalMonth = intEndingMonth And intTotalYear = intEndingYear Then
                        intTotalDays += intEndingDay
                        blnDateNotMatching = False
                    Else
                        intTotalDays += 30
                        intTotalMonth += 1
                    End If
                ElseIf intTotalMonth = 2 Then
                    intLeapYear = intTotalYear Mod 4
                    If intLeapYear = 0 Then
                        If blnFirstRotation = True Then
                            intTotalDays = intTotalDays + 29 - intStartingDay
                            intTotalMonth += 1
                            blnFirstRotation = False
                        ElseIf intTotalMonth = intEndingMinutes And intTotalYear = intEndingYear Then
                            intTotalDays += intEndingDay
                            blnDateNotMatching = False
                        Else
                            intTotalDays += 29
                            intTotalMonth += 1
                        End If
                    Else
                        If blnFirstRotation = True Then
                            intTotalDays = intTotalDays + 28 - intStartingDay
                            intTotalMonth += 1
                            blnFirstRotation = False
                        ElseIf intTotalMonth = intEndingMinutes And intTotalYear = intEndingYear Then
                            intTotalDays += intEndingDay
                            blnDateNotMatching = False
                        Else
                            intTotalDays += 28
                            intTotalMonth += 1
                        End If
                    End If
                End If

            End While
        End If

        'Creating Output
        strReturnDateTime = CStr(intTotalDays) + " Days, " + CStr(intTotalHours) + " Hours, " + CStr(intTotalMinutes) + " Minutes and " + CStr(intTotalSeconds) + " Seconds"

        Return strReturnDateTime

    End Function
    Public Function DaysDifference(ByVal datStartingDate As Date, ByVal datEndingDate As Date) As Integer

        Dim intStartingMonth As Integer
        Dim intStartingDay As Integer
        Dim intStartingYear As Integer
        Dim intStartingHours As Integer
        Dim intStartingMinutes As Integer
        Dim intStartingSecond As Integer
        Dim intEndingMonth As Integer
        Dim intEndingDay As Integer
        Dim intEndingYear As Integer
        Dim intEndingHours As Integer
        Dim intEndingMinutes As Integer
        Dim intEndingSecond As Integer
        Dim intTotalDays As Integer = 0
        Dim intTotalHours As Integer = 0
        Dim intTotalMinutes As Integer = 0
        Dim intTotalSeconds As Integer = 0
        Dim strReturnDateTime As String = ""
        Dim intTotalYear As Integer
        Dim intLeapYear As Integer
        Dim intTotalMonth As Integer
        Dim blnFirstRotation As Boolean
        Dim blnDateNotMatching As Boolean = True

        'Getting Starting Information
        intStartingMonth = datStartingDate.Month
        intStartingDay = datStartingDate.Day
        intStartingYear = datStartingDate.Year
        intStartingHours = datStartingDate.Hour
        intStartingMinutes = datStartingDate.Minute
        intStartingSecond = datStartingDate.Second

        'Getting ending information
        intEndingMonth = datEndingDate.Month
        intEndingDay = datEndingDate.Day
        intEndingYear = datEndingDate.Year
        intEndingHours = datEndingDate.Hour
        intEndingMinutes = datEndingDate.Minute
        intEndingSecond = datEndingDate.Second

        'Performing math for seconds
        If intStartingSecond > intEndingSecond Then
            intEndingMinutes -= 1
            intEndingSecond += 60
        End If
        intTotalSeconds = intEndingSecond - intStartingSecond

        'Performing Math for Minutes
        If intStartingMinutes > intEndingMinutes Then
            intEndingHours -= 1
            intEndingMinutes += 60
        End If
        intTotalMinutes = intEndingMinutes - intStartingMinutes

        'Performing Math for Hours
        If intStartingHours > intEndingHours Then
            intEndingDay -= 1
            intEndingHours += 24
        End If

        intTotalHours = intEndingHours - intStartingHours

        'Setting up to figure the days
        blnFirstRotation = True
        intTotalYear = intStartingYear
        intTotalMonth = intStartingMonth

        If intStartingMonth = intEndingMonth And intStartingYear = intEndingYear And intEndingDay >= intStartingDay Then
            intTotalDays = intEndingDay - intStartingDay
        Else
            'Creating the number of days
            While blnDateNotMatching = True

                If intTotalMonth = 1 Or intTotalMonth = 3 Or intTotalMonth = 5 Or intTotalMonth = 7 Or intTotalMonth = 8 Or intTotalMonth = 10 Or intTotalMonth = 12 Then
                    If blnFirstRotation = True Then
                        intTotalDays = intTotalDays + 31 - intStartingDay
                        blnFirstRotation = False
                        intTotalMonth += 1
                        If intTotalMonth = 13 Then
                            intTotalMonth = 1
                            intTotalYear += 1
                        End If
                    ElseIf intTotalMonth = intEndingMonth And intTotalYear = intEndingYear Then
                        intTotalDays += intEndingDay
                        blnDateNotMatching = False
                    Else
                        intTotalDays += 31
                        intTotalMonth += 1
                    End If
                ElseIf intTotalMonth = 4 Or intTotalMonth = 6 Or intTotalMonth = 9 Or intTotalMonth = 11 Then
                    If blnFirstRotation = True Then
                        intTotalDays = intTotalDays + 30 - intStartingDay
                        blnFirstRotation = False
                        intTotalMonth += 1
                    ElseIf intTotalMonth = intEndingMonth And intTotalYear = intEndingYear Then
                        intTotalDays += intEndingDay
                        blnDateNotMatching = False
                    Else
                        intTotalDays += 30
                        intTotalMonth += 1
                    End If
                ElseIf intTotalMonth = 2 Then
                    intLeapYear = intTotalYear Mod 4
                    If intLeapYear = 0 Then
                        If blnFirstRotation = True Then
                            intTotalDays = intTotalDays + 29 - intStartingDay
                            intTotalMonth += 1
                            blnFirstRotation = False
                        ElseIf intTotalMonth = intEndingMinutes And intTotalYear = intEndingYear Then
                            intTotalDays += intEndingDay
                            blnDateNotMatching = False
                        Else
                            intTotalDays += 29
                            intTotalMonth += 1
                        End If
                    Else
                        If blnFirstRotation = True Then
                            intTotalDays = intTotalDays + 28 - intStartingDay
                            intTotalMonth += 1
                            blnFirstRotation = False
                        ElseIf intTotalMonth = intEndingMonth And intTotalYear = intEndingYear Then
                            intTotalDays += intEndingDay
                            blnDateNotMatching = False
                        Else
                            intTotalDays += 28
                            intTotalMonth += 1
                        End If
                    End If
                End If

            End While
        End If

        'returning value
        Return intTotalDays

    End Function

End Class
