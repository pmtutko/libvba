Attribute VB_Name = "RangeSupport"
Attribute VB_Description = "Variety of support functions operating on Excel ranges and arrays."
Option Explicit

Public Function RangeToStringArray(dataRange As Range) As String()
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
    Dim columnCounter As Long, rowCounter As Long
    For rowCounter = UBound(dataArray, 1) To 1 Step -1
        For columnCounter = UBound(dataArray, 2) To 1 Step -1
            stringArray(rowCounter, columnCounter) = _ 
                                 CStr(dataArray(rowCounter, columnCounter))
        Next columnCounter
    Next rowCounter

    '--- Return the string array
    RangeToStringArray = stringArray
End Function

Public Function ArraysMatch(ByRef array1 As Variant, _
                            ByRef array2 As Variant) As Boolean
Attribute ArraysMatch.VB_Description = "Element-by-element check comparing two arrays. Returns TRUE if all values (and dimensions) are identical. Returns FALSE if anything is different. LIMITATION: can only handle up to five-dimensional arrays"
    '--- basically an element-by-element check comparing two arrays.
    '    returns TRUE if all values (and dimensions) are identical.
    '    returns FALSE if anything is different
    '    LIMITATION: can only handle up to five-dimensional arrays
    
    '--- make sure each array has the same number of dimensions first...
    Dim numDimensions1 As Long
    Dim numValues As Long
    numDimensions1 = 1
    On Error Resume Next
    Err.Clear
    Do While True
        numValues = UBound(array1, numDimensions1)
        If Err.Number > 0 Then
            '--- subtract one because we've gone too far
            numDimensions1 = numDimensions1 - 1
            Err.Clear
            Exit Do
        Else
            numDimensions1 = numDimensions1 + 1
        End If
    Loop
    Dim numDimensions2 As Long
    numDimensions2 = 1
    Do While True
        numValues = UBound(array2, numDimensions2)
        If Err.Number > 0 Then
            '--- subtract one because we've gone too far
            numDimensions2 = numDimensions2 - 1
            Err.Clear
            Exit Do
        Else
            numDimensions2 = numDimensions2 + 1
        End If
    Loop
    If numDimensions1 <> numDimensions2 Then
        ArraysMatch = False
        Exit Function
    End If
    
    '--- now check if these dimensions have the same number of values 
    '    in each dimension...
    On Error GoTo 0
    Dim i As Long
    For i = 1 To numDimensions1
        If UBound(array1, i) <> UBound(array2, i) Then
            ArraysMatch = False
            Exit Function
        End If
    Next i
    
    '--- finally, check each element in every dimension against each 
    '    other to make sure all values are identical
    For i = LBound(array1, 1) To UBound(array1, 1)
        If numDimensions1 = 1 Then
            If array1(i) <> array2(i) Then
                ArraysMatch = False
                Exit Function
            End If
        Else
            Dim j As Long
            For j = LBound(array1, 2) To UBound(array1, 2)
                If numDimensions1 = 2 Then
                    If array1(i, j) <> array2(i, j) Then
                        ArraysMatch = False
                        Exit Function
                    End If
                Else
                    Dim k As Long
                    For k = LBound(array1, 3) To UBound(array1, 3)
                        If numDimensions1 = 3 Then
                            If array1(i, j, k) <> array2(i, j, k) Then
                                ArraysMatch = False
                                Exit Function
                            End If
                        Else
                            Dim m As Long
                            For m = LBound(array1, 4) To UBound(array1, 4)
                                If numDimensions1 = 4 Then
                                    If array1(i, j, k, m) <> array2(i, j, k, m) Then
                                        ArraysMatch = False
                                        Exit Function
                                    End If
                                Else
                                    Dim n As Long
                                    For n = LBound(array1, 5) To UBound(array1, 5)
                                        If numDimensions1 = 5 Then
                                            If array1(i, j, k, m, n) <> array2(i, j, k, m, n) Then
                                                ArraysMatch = False
                                                Exit Function
                                            End If
                                        Else
                                        End If
                                    Next n
                                End If
                            Next m
                        End If
                    Next k
                End If
            Next j
        End If
    Next i
    
    '--- if we get here, everything matches!
    ArraysMatch = True
End Function

Function GetNetworkPath(ByVal DriveName As String) As String
Attribute GetNetworkPath.VB_Description = "Converts a drive letter, e.g. 'W:\', to its fully qualified network path. Useful for saving a network folder location without any user-specific custom mapping"
    '--- Converts a drive letter, e.g. 'W:\', to its fully qualified network 
    '    path useful for saving a network folder location without any user-
    '    specific custom mapping
    '--- from https://www.mrexcel.com/forum/excel-questions/ 
    '            658830-show-full-file-paths-unc-not-mapped-drive-letters-print.html
    Dim objNtWork  As Object
    Dim objDrives  As Object
    Dim lngLoop    As Long
    
    Set objNtWork = CreateObject("WScript.Network")
    Set objDrives = objNtWork.enumnetworkdrives
    
    GetNetworkPath = vbNullString
    For lngLoop = 0 To objDrives.count - 1 Step 2
        If UCase(objDrives.item(lngLoop)) = UCase(DriveName) Then
            GetNetworkPath = objDrives.item(lngLoop + 1)
            Exit For
        End If
    Next
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
