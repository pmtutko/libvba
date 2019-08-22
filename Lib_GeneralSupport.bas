Attribute VB_Name = "Lib_GeneralSupport"
Attribute VB_Description = "Variety of support functions operating on Excel ranges and arrays."
Option Explicit

Public Sub DPrintf(ByVal base As String, ParamArray args())
Attribute DPrintf.VB_Description = "Performs a Debug.Print of the base string, replacing the format token '%s' with the given arguments"
    If UBound(args, 1) = -1 Then
        Debug.Print SPrintf(base)
    Else
        Debug.Print SPrintf(base, args())
    End If
End Sub

Public Function SPrintf(ByVal base As String, ParamArray args()) As String
Attribute SPrintf.VB_Description = "Returns a formatted string combining the base string, replacing the format token '%s' with the given arguments"
    '--- print a formatted string to the immediate window (using Debug.Print)
    '    replacing format tokens in the base string with variables in the args
    Const TOKEN As String = "%s"
    Dim pos1 As Long
    pos1 = InStr(1, base, TOKEN)
    
    If UBound(args, 1) = -1 Then
        '--- if there are no argmuents, then there should be no tokens
        If pos1 = 0 Then
            '--- no arguments, no tokens, just return the base string
            SPrintf = base
        Else
            '--- but we do have tokens, so it's an error
            SPrintf = "DPrint ERROR: no format tokens '" & TOKEN & _
                      "' found in '" & base & "'"
        End If
    
    Else
        '--- there are arguments, so we have to potentially unwind nested
        '    args arrays (see: https://stackoverflow.com/a/30243700/4717755)
        args = ParamArrayDelegated(args)
        
        If pos1 = 0 Then
            '--- if there are no tokens in the string
            SPrintf = "DPrint ERROR: no format tokens '" & TOKEN & _
                      "' found in '" & base & "'"
        Else
            '--- array base index offset calculation: allows 0 or 1 indexing
            Dim offset As Long
            offset = IIf(LBound(args, 1) = 0, 1, 0)
            
            '--- make sure the number of format tokens match the number of arguments
            '    in the parameter array
            Dim totalTokens As Long
            Do While pos1 > 0
                totalTokens = totalTokens + 1
                pos1 = InStr(pos1 + 1, base, TOKEN)
            Loop
            If totalTokens <> (UBound(args, 1) + offset) Then
                SPrintf = "DPrint ERROR: mismatch in number of tokens to " & _
                          "arguments in '" & base & "'"
                Exit Function
            End If
            
            Const TEMP_TOKEN As String = "}~{"
            Dim i As Long
            For i = LBound(args, 1) To UBound(args, 1)
                '--- quick check to see if the TOKEN exists in this input argument
                '    if it does, then the caller wants to see that in the output
                '    so we'll make a temporary replacement until the end
                args(i) = Replace$(args(i), TOKEN, TEMP_TOKEN)
                base = Replace$(base, TOKEN, args(i), , 1)
            Next i
            base = Replace$(base, TEMP_TOKEN, TOKEN)
            SPrintf = base
        End If
    End If
End Function

Public Function ParamArrayDelegated(ParamArray prms() As Variant) As Variant
Attribute ParamArrayDelegated.VB_Description = "Unwinds nested parameter arrays"
    Dim arrPrms() As Variant, arrWrk() As Variant
    'When prms(0) is Array, supposed is delegated from another function
    arrPrms = prms
    If UBound(arrPrms) <> -1 Then
        Do While VarType(arrPrms(0)) >= vbArray And UBound(arrPrms) < 1
            arrWrk = arrPrms(0)
            arrPrms = arrWrk
        Loop
    End If
    ParamArrayDelegated = arrPrms
End Function

Public Function FolderExists(ByVal path As String) As Boolean
Attribute FolderExists.VB_Description = "checks the given path and verifies that it exists and is a folder"
    '--- checks the given path and verifies that it exists and is a folder
    If Right$(path, 1) <> "\" Then
        path = path & "\"
    End If
    FolderExists = Dir(path, vbDirectory) <> vbNullString
End Function

Public Function CreateFolder(sFolder As String) As String
Attribute CreateFolder.VB_Description = "recursively creates all folders in the given path"
    '--- recursively creates all folders in the given path
    '    from:  http://www.freevbcode.com/ShowCode.asp?ID=257
    On Error GoTo ErrorHandler
    Dim s As String
    s = GetPathOnly(sFolder)
    If Dir(s, vbDirectory) = "" Then
        s = CreateFolder(s)
        If Len(s) > 0 Then
            MkDir s
            Debug.Print "mkdir: " & s
        End If
    End If
    CreateFolder = sFolder
    Exit Function

ErrorHandler:
    Exit Function
End Function

Public Function GetPathOnly(sPath As String) As String
Attribute GetPathOnly.VB_Description = "returns the directory path portion of the given fully qualified pathname"
    '--- returns the directory path portion of the given fully qualified
    '    pathname
    '    from:  http://www.freevbcode.com/ShowCode.asp?ID=257
    GetPathOnly = Left(sPath, InStrRev(sPath, "\", Len(sPath)) - 1)
End Function

Public Function IsFileOpen(Filename As String) As Boolean
Attribute IsFileOpen.VB_Description = "checks if the given file is already open by another application"
    '--- checks if the given file is already open by another application
    '    https://stackoverflow.com/a/9373914/4717755
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open Filename For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsFileOpen = False
    Case 70:   IsFileOpen = True    'permission denied
    Case Else: Error ErrNo
    End Select
End Function

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
Attribute GetNetworkPath.VB_Description = "Converts a drive letter, e.g. 'W:', to its fully qualified network path. Useful for saving a network folder location without any user-specific custom mapping"
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
