Attribute VB_Name = "Lib_PivotTableSupport"
Attribute VB_Description = "Variety of support functions operating on Excel Pivot Tables."
Option Explicit

Public Function PivotFieldIsHidden(ByRef pTable As PivotTable, _
                                   ByVal ptFieldName As String) As Boolean
Attribute PivotFieldIsHidden.VB_Description = "Determines if the given field is hidden or visible in the pivot table."
    '--- Determines if the given field is hidden or visible in the pivot table
    Dim field As PivotField
    PivotFieldIsHidden = False
    For Each field In pTable.HiddenFields
        If field.Name = ptFieldName Then
            PivotFieldIsHidden = True
            Exit Function
        End If
    Next field
End Function

Public Function PivotFieldIsData(ByRef pTable As PivotTable, _
                                 ByVal ptFieldName As String) As Boolean
Attribute PivotFieldIsData.VB_Description = "Determines if the given field is data field in the pivot table."
    '--- Determines if the given field is data field in the pivot table
    Dim field As PivotField
    PivotFieldIsData = False
    For Each field In pTable.DataFields
        If field.Name = "Sum of " & ptFieldName Then
            PivotFieldIsData = True
            Exit Function
        End If
    Next field
End Function

Public Function PivotFieldIsRow(ByRef pTable As PivotTable, _
                                ByVal ptFieldName As String) As Boolean
Attribute PivotFieldIsRow.VB_Description = "Determines if the given field is displayed as a row field in the pivot table."
    '--- Determines if the given field is displayed as a row field 
    '    in the pivot table
    Dim field As PivotField
    PivotFieldIsRow = False
    For Each field In pTable.RowFields
        If field.Name = ptFieldName Then
            PivotFieldIsRow = True
            Exit Function
        End If
    Next field
End Function

Public Function PivotFieldIsColumn(ByRef pTable As PivotTable, _
                                   ByVal ptFieldName As String) As Boolean
Attribute PivotFieldIsColumn.VB_Description = "Determines if the given field is displayed as a column field in the pivot table."
    '--- Determines if the given field is displayed as a column field 
    '    in the pivot table
    Dim field As PivotField
    PivotFieldIsColumn = False
    For Each field In pTable.ColumnFields
        If field.Name = ptFieldName Then
            PivotFieldIsColumn = True
            Exit Function
        End If
    Next field
End Function

Public Function PivotFieldPosition(ByRef pTable As PivotTable, _
                                   ByVal ptFieldName As String) As Long
Attribute PivotFieldPosition.VB_Description = "Determines the position of the given field in the pivot table."
    '--- Determines the position of the given field in the pivot table
    Dim field As PivotField
    PivotFieldPosition = 0
    For Each field In pTable.PivotFields
        If field.Name = ptFieldName Then
            If TypeName(field.position) = "Error" Then
                '--- we'll get an error if the field is not included as a
                '    row or column. this isn't a problem, but there's no
                '    real position in this case, so return 0
            Else
                PivotFieldPosition = field.position
            End If
            Exit Function
        End If
    Next field
End Function

Public Function AnyPivotItemsExpanded(ByRef pField As PivotField) As Boolean
Attribute AnyPivotItemsExpanded.VB_Description = "Determines if any displayed items in the pivot table have been expanded"
    '--- Determines if any displayed items in the pivot table have 
    '    been expanded
    Dim pItem As PivotItem
    AnyPivotItemsExpanded = False
    Dim i As Long
    i = 0
    For Each pItem In pField.PivotItems
        i = i + 1
        If pItem.ShowDetail Then
            AnyPivotItemsExpanded = True
            Exit Function
        End If
    Next pItem
End Function

Public Function PivotItemExpanded(ByRef pField As PivotField, _
                                   ByVal itemPosition As Long) As Boolean
Attribute PivotItemExpanded.VB_Description = "Determines if the specific field in the pivot table has been expanded"
    '--- Determines if the specific field in the pivot table has been expanded
    Dim pItem As PivotItem
    PivotItemExpanded = False
    Set pItem = pField.PivotItems(itemPosition)
    If pItem.ShowDetail Then
        PivotItemExpanded = True
        Exit Function
    End If
End Function

Public Sub PivotSetShowDetail(ByRef pTable As PivotTable, _
                              ByVal showFlag as Boolean)
Attribute PivotSetShowDetail.VB_Description = "Expands or collapses all visible fields in the given pivot table."
    '--- Expands or collapses all visible fields in the given pivot table
    Dim rowField As PivotField
    For Each rowField In pTable.RowFields
        If rowField.position <> pTable.RowFields.count Then
            rowField.ShowDetail = showFlag
        End If
    Next rowField
End Sub

public Function PivotItemsShown(pf As PivotField) As Boolean
Attribute PivotItemsShown.VB_Description = "Determines if any items are shown in the given pivot field."
    '--- Determines if any items are shown in the given pivot field
    Dim pi As PivotItem
    For Each pi In pf.PivotItems
        If pi.ShowDetail Then
            PivotItemsShown = True
            Exit For
        End If
    Next pi
End Function

