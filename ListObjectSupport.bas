Attribute VB_Name = "ListObjectSupport"
Attribute VB_Description = "Variety of support functions operating on Excel ListObjects (tables)."
Option Explicit

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
Attribute SaveListObjectFilters.VB_Description = "Detects any filters applied to the given ListObject table and saves them to the FilterCache array parameter."
    Dim ii As Long
    Dim filterRange As String

    filterRange = ""
    With lo.AutoFilter
        filterRange = .Range.Address
        With .Filters
            ReDim FilterCache(1 To .count, 1 To 3)
            For ii = 1 To .count
                With .item(ii)
                    If .On Then
                        #If False Then           ' XL11 code
                            FilterCache(ii, 1) = .Criteria1
                            If .Operator Then
                                FilterCache(ii, 2) = .Operator
                                FilterCache(ii, 3) = .Criteria2
                            End If
                        #Else                    ' first pass XL14
                            Select Case .Operator
  
                            Case 1, 2            'xlAnd, xlOr
                                FilterCache(ii, 1) = .Criteria1
                                FilterCache(ii, 2) = .Operator
                                FilterCache(ii, 3) = .Criteria2
  
                            Case 0, 3 To 7       ' no operator, xlTop10Items,
                                                 ' xlBottom10Items, xlTop10Percent,
                                                 ' xlBottom10Percent, xlFilterValues
                                FilterCache(ii, 1) = .Criteria1
                                FilterCache(ii, 2) = .Operator
  
                            Case Else            ' These are not correctly restored; 
                                                 ' there's someting in Criteria1 but
                                                 ' can't save it.
                                FilterCache(ii, 2) = .Operator
                                ' FilterCache(ii, 1) = .Criteria1  ' <-- Generates an error
                                ' No error in next statement, but couldn't do restore operation
                                ' Set FilterCache(ii, 1) = .Criteria1
 
                            End Select
                        #End If
                    End If
                End With                         ' .Item(ii)
            Next
        End With                                 ' .Filters
    End With                                     ' wks.AutoFilter

End Function

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
Attribute RestoreListObjectFilters.VB_Description = "Uses the FilterCache array as an input parameter and (re)applies the saved filters to the given ListObject table."
    Dim col As Long

    If lo.Range.Address <> "" Then
        For col = 1 To UBound(FilterCache(), 1)
  
            #If False Then                       ' XL11
                If Not IsEmpty(FilterCache(col, 1)) Then
                    If FilterCache(col, 2) Then
                        lo.AutoFilter field:=col, _
                                      Criteria1:=FilterCache(col, 1), _
                                      Operator:=FilterCache(col, 2), _
                                      Criteria2:=FilterCache(col, 3)
                    Else
                        lo.AutoFilter field:=col, _
                                      Criteria1:=FilterCache(col, 1)
                    End If
                End If
            #Else
  
                If Not IsEmpty(FilterCache(col, 2)) Then
                    Select Case FilterCache(col, 2)
  
                    Case 0                       ' no operator
                        lo.Range.AutoFilter field:=col, _
                                            Criteria1:=FilterCache(col, 1) ' Do NOT reload 'Operator'
 
                    Case 1, 2                    'xlAnd, xlOr
                        lo.Range.AutoFilter field:=col, _
                                            Criteria1:=FilterCache(col, 1), _
                                            Operator:=FilterCache(col, 2), _
                                            Criteria2:=FilterCache(col, 3)
  
                    Case 3 To 6                  ' xlTop10Items, xlBottom10Items, xlTop10Percent,  xlBottom10Percent
                        #If True Then
                            lo.Range.AutoFilter field:=col, _
                                                Criteria1:=FilterCache(col, 1) ' Do NOT reload 'Operator' , it doesn't work
                            ' wks.AutoFilter.Filters.Item(col).Operator = FilterCache(col, 2)
                        #Else                    ' Trying to restore Operator as well as Criteria ..
                            ' Including the 'Operator:=' arguement leads to error.
                            ' Criteria1 is expressed as if for a FALSE .Operator
                            lo.Range.AutoFilter field:=col, _
                                                Criteria1:=FilterCache(col, 1), _
                                                Operator:=FilterCache(col, 2)
                        #End If
  
                    Case 7                       'xlFilterValues
                        lo.Range.AutoFilter field:=col, _
                                            Criteria1:=FilterCache(col, 1), _
                                            Operator:=FilterCache(col, 2)
  
                        #If False Then           ' Switch on filters on cell formats
                            ' These statements restore the filter, but cannot reset the pass Criteria, so the filter hides all data.
                            ' Leave it off instead.
                        Case Else                ' (Various filters on data format)
                            lo.RangeAutoFilter field:=col, _
                                               Operator:=FilterCache(col, 2)
                        #End If                  ' Switch on filters on cell formats
 
                    End Select
                End If
  
            #End If                              ' XL11 / XL14
        Next col
    End If

End Sub

Public Sub TableCopy(ByRef srcTable As ListObject, _
                     ByRef dstTable As ListObject, _
                     Optional copySrcFilters As Boolean = True)
Attribute TableCopy.VB_Description = "Copies the data from the source table to the destination table in order to guarantee a clean copy, ALL of the data in the destination table is destroyed and replaced with data from the source optionally copies any existing filters from the source table to the destination table and applies them."
    '--- copies the data from the source table to the destination table
    '    in order to guarantee a clean copy, ALL of the data in the
    '    destination table is destroyed and replaced with data from the source
    '    optionally copies any existing filters from the source table to
    '    the destination table and applies them
    
    If copySrcFilters Then
        '--- save any filters applied to the source table...
        Dim savedTableFilters() As Variant
        SaveListObjectFilters srcTable, savedTableFilters
    End If
    
    '--- save the state of the application to restore after this block
    Dim appStateAlerts As Boolean
    appStateAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    '--- guarantee that ALL of the data in the table is fresh and
    '    accurate by deleting all rows of the destination, then
    '    copying all rows from the source
    dstTable.AutoFilter.ShowAllData
    dstTable.DataBodyRange.Delete
    dstTable.Range(2, 1) = "temp"
    Dim srcData As Variant
    srcData = srcTable.DataBodyRange      'transfer to memory array
    dstTable.DataBodyRange.Resize(UBound(srcData, 1), _
                                  UBound(srcData, 2)) = srcData
    
    If copySrcFilters Then
        '--- apply any existing filters to the table so it looks the same
        RestoreListObjectFilters dstTable, savedTableFilters
    End If
    
    '--- restore application state
    Application.DisplayAlerts = appStateAlerts
End Sub

