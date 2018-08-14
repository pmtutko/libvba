Attribute VB_Name = "Lib_GeneralSupport"
Attribute VB_Description = "Variety of support functions operating on Excel ranges and arrays."
Option Explicit

Public Sub DebugClearImmediateWindow()
Attribute DebugClearImmediateWindow.VB_Description = "Convenience method to clear the debugger immediate window"
    '--- Convenience method to clear the debugger immediate window
    Application.SendKeys "^g ^a {DEL}"
End Sub

Public Sub CopyTextToClipboard(Text As String)
Attribute CopyTextToClipboard.VB_Description = "Copies the given text to the system clipboard"
    '--- from: https://stackoverflow.com/a/25336423/4717755
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText Text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub

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

Public Function ArraysMatch(ByRef array1 As Variant, _
                            ByRef array2 As Variant) As Boolean
Attribute ArraysMatch.VB_Description = "Element-by-element check comparing two arrays. Returns TRUE if all values (and dimensions) are identical. Returns FALSE if anything is different. LIMITATION: can handle up to five-dimensional arrays"
    '--- basically an element-by-element check comparing two arrays.
    '    returns TRUE if all values (and dimensions) are identical.
    '    returns FALSE if anything is different
    '    LIMITATION: can handle up to five-dimensional arrays
    
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
    
    '--- now check if these dimensions have the same number of values in each dimension...
    On Error GoTo 0
    Dim i As Long
    For i = 1 To numDimensions1
        If UBound(array1, i) <> UBound(array2, i) Then
            ArraysMatch = False
            Exit Function
        End If
    Next i
    
    '--- finally, check each element in every dimension against each other to make sure
    '    all values are identical
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

Private Sub TestArraysMatch()
    Dim a2 As Variant
    Dim b2 As Variant
    Dim c2 As Variant
    Dim d2 As Variant
    a2 = [{1,1;2,2;3,3;4,4}]
    b2 = [{1,1;2,2;3,3;4,4}]
    d2 = [{1,1;2,2;3,3;4,4;5,5}]
    c2 = [{1,1,1,1;2,2,2,2;3,3,3,3;4,4,4,4}]
    Debug.Print "compare a2(" & UBound(a2, 1) & ","; UBound(a2, 2) & ") and c2(" & _
                UBound(c2, 1) & ","; UBound(c2, 2) & ") [FALSE]:" & _
                ArraysMatch(a2, c2)
    Debug.Print "compare a2(" & UBound(a2, 1) & ","; UBound(a2, 2) & ") and b2(" & _
                UBound(b2, 1) & ","; UBound(b2, 2) & ") [TRUE ]:" & _
                ArraysMatch(a2, b2)
    'Debug.Print "compare v(2, 4) and y(2, 7)    should be FALSE: " & ArraysMatch(v, y); ""
    'Debug.Print "compare v(2, 4) and w(2, 4)    should be TRUE : " & ArraysMatch(v, w); ""
    
    Dim t1 As Variant
    Dim t2 As Variant
    t1 = Resources.ListObjects("STSResources").DataBodyRange
    t2 = WorkingData.ListObjects("STSResourcesShadow").DataBodyRange
    Debug.Print "resources tables = " & ArraysMatch(t1, t2)
    
    WorkingData.ListObjects("STSResourcesShadow").DataBodyRange.Delete
    WorkingData.ListObjects("STSResourcesShadow").Range(2, 1) = "temp"
    WorkingData.ListObjects("STSResourcesShadow").DataBodyRange.Resize(UBound(t1, 1), UBound(t1, 2)) = t1
    t2 = WorkingData.ListObjects("STSResourcesShadow").DataBodyRange
    Debug.Print "resources tables = " & ArraysMatch(t1, t2)
End Sub
