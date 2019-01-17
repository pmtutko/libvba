Attribute VB_Name = "Lib_CalendarSupport"
Attribute VB_Description = "Support functions for calendar calculations and day/date determination."
'@Folder("Libraries")
Option Explicit
Option Compare Text

'--- from: http://www.cpearson.com/excel/betternetworkdays.aspx

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' EDaysOfWeek
' Days of the week to exclude. This is a bit-field
' enum, so that its values can be added or OR'd
' together to specify more than one day. E.g,.
' to exclude Tuesday and Saturday, use
' (Tuesday+Saturday), or (Tuesday OR Saturday)
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Enum EDaysOfWeek
    Sunday = 1                                   ' 2 ^ (vbSunday - 1)
    Monday = 2                                   ' 2 ^ (vbMonday - 1)
    Tuesday = 4                                  ' 2 ^ (vbTuesday - 1)
    Wednesday = 8                                ' 2 ^ (vbWednesday - 1)
    Thursday = 16                                ' 2 ^ (vbThursday - 1)
    Friday = 32                                  ' 2 ^ (vbFriday - 1)
    Saturday = 64                                ' 2 ^ (vbSaturday - 1)
End Enum

Public Function IsAWorkDay(ByRef thisDay As Date) As Boolean
    Dim calArray As Variant
    calArray = Calendars.GetCalendarAsArray()
    IsAWorkDay = (Workday2(thisDay - 1, 1, Sunday + Saturday, calArray) = thisDay)
End Function

Public Function NetWorkdays2(ByVal StartDate As Date, _
                             ByVal EndDate As Date, _
                             ByVal ExcludeDaysOfWeek As Long, _
                             Optional ByRef Holidays As Variant) As Variant
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' NetWorkdays2
    ' This function calcluates the number of days between StartDate and EndDate
    ' excluding those days of the week specified by ExcludeDaysOfWeek and
    ' optionally excluding dates in Holidays. ExcludeDaysOfWeek is a
    ' value from the table below.
    '       1  = Sunday     = 2 ^ (vbSunday - 1)
    '       2  = Monday     = 2 ^ (vbMonday - 1)
    '       4  = Tuesday    = 2 ^ (vbTuesday - 1)
    '       8  = Wednesday  = 2 ^ (vbWednesday - 1)
    '       16 = Thursday   = 2 ^ (vbThursday - 1)
    '       32 = Friday     = 2 ^ (vbFriday - 1)
    '       64 = Saturday   = 2 ^ (vbSaturday - 1)
    ' To exclude multiple days, add the values in the table together. For example,
    ' to exclude Mondays and Wednesdays, set ExcludeDaysOfWeek to 10 = 8 + 2 =
    ' Monday + Wednesday.
    ' If StartDate is less than or equal to EndDate, the result is positive. If
    ' StartDate is greater than EndDate, the result is negative. If either
    ' StartDate or EndDate is less than or equal to 0, the result is a
    ' #NUM error. If ExcludeDaysOfWeek is less than 0 or greater than or
    ' equal to 127 (all days excluded), the result is a #NUM error.
    ' Holidays is optional and may be a single constant value, an array of values,
    ' or a worksheet range of cells.
    ' This function can be used as a replacement for the NETWORKDAYS worksheet
    ' function. With NETWORKDAYS, the excluded days of week are hard coded
    ' as Saturday and Sunday. You cannot exlcude other days of the week. This
    ' function allows you to exclude any number of days of the week (with the
    ' exception of excluding all days of week), from 0 to 6 days. If
    ' ExcludeDaysOfWeek = 65 (Sunday + Saturday), the result is the same as
    ' NETWORKDAYS.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim TestDayOfWeek As Long
    Dim TestDate As Date
    Dim Count As Long
    Dim Stp As Long
    Dim Holiday As Variant
    Dim Exclude As Boolean

    If ExcludeDaysOfWeek < 0 Or ExcludeDaysOfWeek >= 127 Then
        ' invalid value for ExcludeDaysOfWeek. get out with error.
        NetWorkdays2 = CVErr(xlErrNum)
        Exit Function
    End If

    If StartDate <= 0 Or EndDate <= 0 Then
        ' invalid date. get out with error.
        NetWorkdays2 = CVErr(xlErrNum)
        Exit Function
    End If

    ' set the value used for the Step in
    ' the For loop.
    If StartDate <= EndDate Then
        Stp = 1
    Else
        Stp = -1
    End If

    For TestDate = StartDate To EndDate Step Stp
        ' get the bit pattern of the weekday of TestDate
        TestDayOfWeek = 2 ^ (Weekday(TestDate, vbSunday) - 1)
        If (TestDayOfWeek And ExcludeDaysOfWeek) = 0 Then
            ' do not exclude this day of week
            If IsMissing(Holidays) = True Then
                ' count day
                Count = Count + 1
            Else
                Exclude = False
                ' holidays provided. test date for holiday.
                If IsObject(Holidays) = True Then
                    ' assume Excel.Range
                    For Each Holiday In Holidays
                        If Holiday.value = TestDate Then
                            Exclude = True
                            Exit For
                        End If
                    Next Holiday
                Else
                    ' not an Excel.Range
                    If IsArray(Holidays) = True Then
                        For Each Holiday In Holidays
                            If Int(Holiday) = TestDate Then
                                Exclude = True
                                Exit For
                            End If
                        Next Holiday
                    Else
                        ' not an array or range, assume single value
                        If TestDate = Holidays Then
                            Exclude = True
                        End If
                    End If
                End If
                If Exclude = False Then
                    Count = Count + 1
                End If
            End If
        Else
            ' excluded day of week. do nothing
        End If
    Next TestDate
    ' return the result, positive or negative based on Stp.
    NetWorkdays2 = Count * Stp

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Workday2
' This is a replacement for the ATP WORKDAY function. It
' expands on WORKDAY by allowing you to specify any number
' of days of the week to exclude.
'   StartDate       The date on which the period starts.
'   DaysRequired    The number of workdays to include
'                   in the period.
'   ExcludeDOW      The sum of the values in EDaysOfWeek
'                   to exclude. E..g, to exclude Tuesday
'                   and Saturday, pass Tuesday+Saturday in
'                   this parameter.
'   Holidays        an array or range of dates to exclude
'                   from the period.
' RESULT:           A date that is DaysRequired past
'                   StartDate, excluding holidays and
'                   excluded days of the week.
' Because it is possible that combinations of holidays and
' excluded days of the week could make an end date impossible
' to determine (e.g., exclude all days of the week), the latest
' date that will be calculated is StartDate + (10 * DaysRequired).
' This limit is controlled by the RunawayLoopControl variable.
' If DaysRequired is less than zero, the result is #VALUE. If
' the RunawayLoopControl value is exceeded, the result is #VALUE.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Workday2(ByVal StartDate As Date, _
                         ByVal DaysRequired As Long, _
                         ByVal ExcludeDOW As EDaysOfWeek, _
                         Optional ByRef Holidays As Variant) As Variant
    Dim N As Long                                ' generic counter
    Dim C As Long                                ' days actually worked
    Dim TestDate As Date                         ' incrementing date
    Dim HNdx As Long                             ' holidays index
    Dim CurDOW As EDaysOfWeek                    ' day of week of TestDate
    Dim IsHoliday As Boolean                     ' is TestDate a holiday?
    Dim RunawayLoopControl As Long               ' prevent infinite looping
    Dim V As Variant                             ' For Each loop variable for Holidays.

    If DaysRequired < 0 Then
        ' day required must be greater than or equal
        ' to zero.
        Workday2 = CVErr(xlErrValue)
        Exit Function
    ElseIf DaysRequired = 0 Then
        Workday2 = StartDate
        Exit Function
    End If

    If ExcludeDOW >= (Sunday + Monday + Tuesday + Wednesday + _
                      Thursday + Friday + Saturday) Then
        ' all days of week excluded. get out with error.
        Workday2 = CVErr(xlErrValue)
        Exit Function
    End If

    ' this prevents an infinite loop which is possible
    ' under certain circumstances.
    RunawayLoopControl = DaysRequired * 10000
    N = 0
    C = 0
    ' loop until the number of actual days worked (C)
    ' is equal to the specified DaysRequired.
    Do Until C = DaysRequired
        N = N + 1
        TestDate = StartDate + N
        CurDOW = 2 ^ (Weekday(TestDate) - 1)
        If (CurDOW And ExcludeDOW) = 0 Then
            ' not excluded day of week. continue.
            IsHoliday = False
            ' test for holidays
            If IsMissing(Holidays) = False Then
                For Each V In Holidays
                    If V = TestDate Then
                        IsHoliday = True
                        ' TestDate is a holiday. get out and
                        ' don't count it.
                        Exit For
                    End If
                Next V
            End If
            If IsHoliday = False Then
                ' TestDate is not a holiday. Include the date.
                C = C + 1
            End If
        End If
        If N > RunawayLoopControl Then
            ' out of control loop. get out with #VALUE
            Workday2 = CVErr(xlErrValue)
            Exit Function
        End If
    Loop
    ' return the result
    Workday2 = StartDate + N
End Function


