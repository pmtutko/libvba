# libvba
Library of helpful VBA functions and subs for use in any project.

## Lib_CalendarSupport
This module copies functions directly from Craig Pearson's excellent website discussing VBA code, specifically the describing a [Better NetWorkdays](http://www.cpearson.com/excel/betternetworkdays.aspx) function.

### Public Interface
```VBA
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

Public Function NetWorkdays2(ByVal StartDate As Date, _
                             ByVal EndDate As Date, _
                             ByVal ExcludeDaysOfWeek As Long, _
                             Optional ByRef Holidays As Variant) As Variant
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

```

## Lib_GeneralSupport
Variety of support functions operating on common VBA variables and system calls.

### Public Interface
```VBA
Public Function FileExists(ByVal fName As String) As Boolean
    '--- from: https://stackoverflow.com/a/28237845/4717755
    'Returns TRUE if the provided name points to an existing file.
    'Returns FALSE if not existing, or if it's a folder

Public Function FileTimestamp(ByVal fName As String) As Date
    '--- returns the last modified timestamp for the given file

Public Function FolderExists(ByVal path As String) As Boolean
    '--- checks the given path and verifies that it exists and is a folder

Public Function CreateFolder(ByVal folderPath As String) As String
    '--- recursively creates all folders in the given path
    '    from:  http://www.freevbcode.com/ShowCode.asp?ID=257

Public Function GetPathOnly(ByVal folderPath As String) As String
    '--- returns the directory path portion of the given fully qualified
    '    pathname

Public Function IsFileOpen(ByVal Filename As String) As Boolean
    '--- checks if the given file is already open by another application
    '    https://stackoverflow.com/a/9373914/4717755

Public Sub DebugClearImmediateWindow()
    '--- Convenience method to clear the debugger immediate window

Public Sub CopyTextToClipboard(ByVal text As String)
    '--- from: https://stackoverflow.com/a/25336423/4717755
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014

Public Function GetNetworkPath(ByVal driveName As String) As String
    '--- Converts a drive letter, e.g. 'W:\', to its fully qualified network
    '    path useful for saving a network folder location without any user-
    '    specific custom mapping
    '--- from https://www.mrexcel.com/forum/excel-questions/
    '            658830-show-full-file-paths-unc-not-mapped-drive-letters-print.html

Public Function SelectFile(ByVal thisTitle As String, _
                           ByVal thisFilterDescription As String, _
                           ByVal thisFilter As String, _
                           Optional ByVal thisPath As String = vbNullString) As String
    '--- NOTE: this function works within Excel, but doesn't work in MS Project
    '          if you need this within Project, you'll have to open an instance
    '          of Excel or Word and change Application.FileDialog to xlApp.FileDialog
    '          (where Dim xlApp As Excel.Application; Set xlApp = AttachToExcelApplication())

Public Function PadLeft(ByRef text As String, _
                        ByVal totalLength As Long, _
                        Optional ByVal padCharacter As String = " ") As String
    '--- adds padding to the left of the given string
    '    from: https://stackoverflow.com/a/12060429/4717755

Public Function PadRight(ByVal text As String, _
                  ByVal totalLength As Long, _
                  Optional ByVal padCharacter As String = " ") As String
    '--- adds padding to the right of the given string
    '    from: https://stackoverflow.com/a/12060429/4717755

Public Function PadCenter(ByVal text As String, _
                          ByVal totalLength As Long, _
                          Optional ByVal padCharacter As String = " ") As String

Public Function ArraysMatch(ByRef array1 As Variant, _
                            ByRef array2 As Variant) As Boolean
    '--- basically an element-by-element check comparing two arrays.
    '    returns TRUE if all values (and dimensions) are identical.
    '    returns FALSE if anything is different
    '    LIMITATION: can handle up to five-dimensional arrays
```

## Lib_ListObjectSupport
Variety of support functions operating on Excel ListObjects (tables).

### Public Interface
```VBA
'~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Sub:  SaveListObjectFilters
' Purpose:  Save filter on worksheet
' Returns:  wks.AutoFilterMode when function entered
' Source: http://stackoverflow.com/questions/9489126/
'                      in-excel-vba-how-do-i-save-restore-a-user-defined-filter
'
' Arguments:
'  [Name]  [Type]  [Description]
'  wks  I/P  Worksheet that filter may reside on
'  FilterRange O/P  Range on which filter is applied as string; "" if no filter
'  FilterCache O/P  Variant dynamic array in which to save filter
'
' Author:  Based on MS Excel AutoFilter Object help file
'
' Modifications:
' 2006/12/11 Phil Spencer: Adapted as general purpose routine
' 2007/03/23 PJS: Now turns off .AutoFilterMode
' 2013/03/13 PJS: Initial mods for XL14, which has more operators
' 2013/05/31 P.H.: Changed to save list-object filters
Public Function SaveListObjectFilters(ByRef lo As ListObject, _
                                      ByRef FilterCache() As Variant) As Boolean

'~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Sub:  RestoreListObjectFilters
' Purpose:  Restore filter on listobject
' Source: http://stackoverflow.com/questions/9489126/in-excel-vba-how-do-i-save-restore-a-user-defined-filter
' Arguments:
'  [Name]  [Type]  [Description]
'  wks  I/P  Worksheet that filter resides on
'  FilterRange I/P  Range on which filter is applied
'  FilterCache I/P  Variant dynamic array containing saved filter
'
' Author:  Based on MS Excel AutoFilter Object help file
'
' Modifications:
' 2006/12/11 Phil Spencer: Adapted as general purpose routine
' 2013/03/13 PJS: Initial mods for XL14, which has more operators
' 2013/05/31 P.H.: Changed to restore list-object filters
'
' Comments:
'----------------------------
Public Sub RestoreListObjectFilters(ByRef lo As ListObject, _
                                    ByRef FilterCache() As Variant)

Public Sub TableCopy(ByRef srcTable As ListObject, _
                     ByRef dstTable As ListObject, _
                     Optional copySrcFilters As Boolean = True)
    '--- copies the data from the source table to the destination table
    '    in order to guarantee a clean copy, ALL of the data in the
    '    destination table is destroyed and replaced with data from the source
    '    optionally copies any existing filters from the source table to
    '    the destination table and applies them
```

## Lib_MSProjectSupport
Variety of support functions operating on MS Project

### Public Interface
```VBA
Public Function IsMSProjectRunning() As Boolean
    '--- quick check to see if an instance of MS Project is running

Public Function AttachToMSProjectApplication() As MSProject.Application
    '--- finds an existing and running instance of MS Project, or starts
    '    the application if one is not already running

Public Function ProjectGetCustomFieldItems(ByVal fieldId As Long) As Dictionary
    '--- returns a collection of the lookup items assigned to the given field

Public Function ProjectCustomFieldHasItems(ByVal fieldId As Long) As Boolean
    '--- determines if the given field has any lookup values

Public Function ProjectCustomFieldItemCount(ByVal fieldId As Long) As Long
    '--- determines the number of lookup items in the given field
```

## Lib_PerformanceSupport
Methods to control disabling/enabling of the Application level screen updates. Supports call nesting and debug messaging, plus high precision timer calls.

### Public Interface
```VBA
Public Sub ReportUpdateState()
    '--- Prints to the immediate window the current state and values of the Application update controls

Public Sub DisableUpdates(Optional ByVal debugMsg As String = vbNullString, _
                          Optional ByVal forceZero As Boolean = False)
    '--- Disables Application level updates and events and saves their initial state to be 
    '    restored later. Supports nested calls. Displays debug messages according to the 
    '    module-global DEBUG_MODE flag.

Public Sub EnableUpdates(Optional ByVal debugMsg As String = vbNullString, _
                         Optional ByVal forceZero As Boolean = False)
    '--- Restores Application level updates and events to their state, prior to the *first* 
    '    DisableUpdates call. Supports nested calls. Displays debug messages according to 
    '    the module-global DEBUG_MODE flag.

' Precision Timer Controls
Public Sub StartCounter()
    '--- Captures the high precision counter value to use as a starting
    '    reference time.

Public Function TimeElapsed() As Double
    '--- Returns the time elapsed since the call to StartCounter in microseconds
```

## Lib_PivotTableSupport
Variety of support functions operating on Excel Pivot Tables.

### Public Interface
```VBA
Public Function PivotFieldIsHidden(ByRef pTable As PivotTable, _
                                   ByVal ptFieldName As String) As Boolean
    '--- Determines if the given field is hidden or visible in the pivot table

Public Function PivotFieldIsData(ByRef pTable As PivotTable, _
                                 ByVal ptFieldName As String) As Boolean
    '--- Determines if the given field is data field in the pivot table

Public Function PivotFieldIsRow(ByRef pTable As PivotTable, _
                                ByVal ptFieldName As String) As Boolean
    '--- Determines if the given field is displayed as a row field
    '    in the pivot table

Public Function PivotFieldIsColumn(ByRef pTable As PivotTable, _
                                   ByVal ptFieldName As String) As Boolean
    '--- Determines if the given field is displayed as a column field
    '    in the pivot table

Public Function PivotFieldPosition(ByRef pTable As PivotTable, _
                                   ByVal ptFieldName As String) As Long
    '--- Determines the position of the given field in the pivot table

Public Function AnyPivotItemsExpanded(ByRef pField As PivotField) As Boolean
    '--- Determines if any displayed items in the pivot table have
    '    been expanded

Public Function PivotItemExpanded(ByRef pField As PivotField, _
                                   ByVal itemPosition As Long) As Boolean
    '--- Determines if the specific field in the pivot table has been expanded

Public Sub PivotSetShowDetail(ByRef pTable As PivotTable, _
                              ByVal showFlag As Boolean)
    '--- Expands or collapses all visible fields in the given pivot table

Public Function PivotItemsShown(pf As PivotField) As Boolean
    '--- Determines if any items are shown in the given pivot field

Public Sub PivotCollapseAllItems(ByRef thisPivot As PivotTable)
    '--- Collapses all rows and items in the given pivot table

Public Sub PivotExpandAllItems(ByRef thisPivot As PivotTable)
    '--- Expands all rows and items in the given pivot table
```

## Lib_RangeSupport
Variety of support functions operating on Excel ranges and arrays.

### Public Interface
```VBA
Public Sub ColorizeDataRange(ByRef data As Variant, _
                              ByRef interiorColor As Variant, _
                              ByRef fontColor As Variant)
    '--- If the given 'data' parameter is a Range, then the given interior
    '    and font colors are set

Public Function RangeToStringArray(ByRef dataRange As Range) As String()
    '--- converts a range of cells into an array of strings containing the
    '    cell data. the array dimensions are the same as the rows and columns
    '    of the source data

Public Sub DataCopy(ByRef srcArea As Range, ByRef dstArea As Range)
    '--- copies the source data to the destination
    '    for a guarantee of a clear data copy, the destination
    '    area is erased first, then the source copied into the
    '    resized area

Public Function SuspendDataFiltering(ByRef dataSheet As Worksheet, _
                                     ByRef filterArray() As Variant, _
                                     ByRef currentFiltRange As String) As Boolean
    '--- If the AutoFilterMode is enabled on the given worksheet, the filters
    '    are captured in the supplied filterArray and returned, and
    '    AutoFilterMode is disabled.

Public Sub RestoreDataFiltering(ByRef dataSheet As Worksheet, _
                                ByRef filterArray() As Variant, _
                                ByRef currentFiltRange As String, _
                                ByVal dataWasFiltered As Boolean)
    '--- Using the filterArray created by SuspendDataFiltering, the filters
    '    are applied to the given worksheet and AutoFilterMode is re-enabled.
```

## Lib_StringSupport
Variety of support functions operating on VBA Strings.

### Public Interface
```VBA
Public Function CleanString(ByVal inString As String) As String
    '--- strips out ANY non-printable character and the designated
    '    whitespace from the string and returns the result

Public Sub CopyTextToClipboard(text As String)
    '--- from: https://stackoverflow.com/a/25336423/4717755
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
```

