Attribute VB_Name = "Lib_RangeSupport"
Attribute VB_Description = "Variety of support functions operating on Excel ranges and arrays."
'@Folder("Libraries")
Option Explicit

Public Sub ColorizeDataRange(ByRef data As Variant, _
                              ByRef interiorColor As Variant, _
                              ByRef fontColor As Variant)
Attribute ColorizeDataRange.VB_Description = "If the given 'data' parameter is a Range, then the given interior and font colors are set"
    '--- If the given 'data' parameter is a Range, then the given interior
    '    and font colors are set
    If TypeName(data) = "Range" Then
        data.Interior.Color = interiorColor
        data.Font.Color = fontColor
    End If
End Sub

Public Function RangeToStringArray(ByRef dataRange As Range) As String()
Attribute RangeToStringArray.VB_Description = "Converts a range of cells into an array of strings containing the cell data. the array dimensions are the same as the rows and columns of the source data"
    '--- converts a range of cells into an array of strings containing the
    '    cell data. the array dimensions are the same as the rows and columns
    '    of the source data

    ' create a memory array from the data area
    Dim dataArray As Variant
    dataArray = dataRange.value

    '--- Set up a string array for them
    Dim stringArray() As String
    ReDim stringArray(1 To UBound(dataArray, 1), 1 To UBound(dataArray, 2))

    '--- Put them in there!
    Dim columnCounter As Long
    Dim rowCounter As Long
    For rowCounter = UBound(dataArray, 1) To 1 Step -1
        For columnCounter = UBound(dataArray, 2) To 1 Step -1
            stringArray(rowCounter, columnCounter) = _
                                 CStr(dataArray(rowCounter, columnCounter))
        Next columnCounter
    Next rowCounter

    '--- Return the string array
    RangeToStringArray = stringArray
End Function

Public Sub DataCopy(ByRef srcArea As Range, ByRef dstArea As Range)
Attribute DataCopy.VB_Description = "Copies the source data to the destination for a guarantee of a clear data copy, the destination area is erased first, then the source copied into the resized area."
    '--- copies the source data to the destination
    '    for a guarantee of a clear data copy, the destination
    '    area is erased first, then the source copied into the
    '    resized area
    Dim srcData As Variant
    srcData = srcArea          'source data to a memory array
    dstArea.ClearContents
    dstArea.Resize(UBound(srcData, 1), UBound(srcData, 2)) = srcData
End Sub

Public Function SuspendDataFiltering(ByRef dataSheet As Worksheet, _
                                     ByRef filterArray() As Variant, _
                                     ByRef currentFiltRange As String) As Boolean
Attribute SuspendDataFiltering.VB_Description = "If the AutoFilterMode is enabled on the given worksheet, the filters are captured in the supplied filterArray and returned, and AutoFilterMode is disabled."
    '--- If the AutoFilterMode is enabled on the given worksheet, the filters
    '    are captured in the supplied filterArray and returned, and
    '    AutoFilterMode is disabled.
    SuspendDataFiltering = False
    If dataSheet.AutoFilterMode = True Then
        SuspendDataFiltering = True
        ' Capture AutoFilter settings
        With dataSheet.AutoFilter
            currentFiltRange = .Range.Address
            With .Filters
                ReDim filterArray(1 To .Count, 1 To 3)
                Dim f As Long
                For f = 1 To .Count
                    With .Item(f)
                        If .On Then
                            filterArray(f, 1) = .Criteria1
                            If .Operator Then
                                filterArray(f, 2) = .Operator
                                filterArray(f, 3) = .Criteria2 'simply delete this line to make it work in Excel 2010
                            End If
                        End If
                    End With
                Next f
            End With
        End With
        'Remove AutoFilter
        dataSheet.AutoFilterMode = False
    End If
End Function

Public Sub RestoreDataFiltering(ByRef dataSheet As Worksheet, _
                                ByRef filterArray() As Variant, _
                                ByRef currentFiltRange As String, _
                                ByVal dataWasFiltered As Boolean)
Attribute RestoreDataFiltering.VB_Description = "Using the filterArray created by SuspendDataFiltering, the filters are applied to the given worksheet and AutoFilterMode is re-enabled."
    '--- Using the filterArray created by SuspendDataFiltering, the filters
    '    are applied to the given worksheet and AutoFilterMode is re-enabled.
    Dim col As Long
    If dataWasFiltered Then
        ' Restore Filter settings
        dataSheet.Range("A1").AutoFilter
        For col = 1 To UBound(filterArray(), 1)
            If Not IsEmpty(filterArray(col, 1)) Then
                If filterArray(col, 2) Then
                    dataSheet.Range(currentFiltRange).AutoFilter field:=col, _
                                                                 Criteria1:=filterArray(col, 1), _
                                                                 Operator:=filterArray(col, 2), _
                                                                 Criteria2:=filterArray(col, 3)
                Else
                    dataSheet.Range(currentFiltRange).AutoFilter field:=col, _
                                                                 Criteria1:=filterArray(col, 1)
                End If
            End If
        Next col
    End If
End Sub

