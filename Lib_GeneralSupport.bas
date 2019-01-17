Attribute VB_Name = "Lib_GeneralSupport"
Attribute VB_Description = "Variety of support functions operating on common VBA variables and system calls."
'@Folder("Libraries")
Option Explicit

Public Function FileExists(ByVal fName As String) As Boolean
    '--- from: https://stackoverflow.com/a/28237845/4717755
    'Returns TRUE if the provided name points to an existing file.
    'Returns FALSE if not existing, or if it's a folder
    On Error Resume Next
    FileExists = ((GetAttr(fName) And vbDirectory) <> vbDirectory)
    On Error GoTo 0
End Function

Public Function FileTimestamp(ByVal fName As String) As Date
    '--- returns the last modified timestamp for the given file
    If FileExists(fName) Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        FileTimestamp = fso.GetFile(fName).DateLastModified
    Else
        FileTimestamp = CVErr(xlErrName)
    End If
End Function

Public Function FolderExists(ByVal path As String) As Boolean
Attribute FolderExists.VB_Description = "checks the given path and verifies that it exists and is a folder"
    '--- checks the given path and verifies that it exists and is a folder
    Dim folderPath As String
    folderPath = path
    If Right$(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    FolderExists = Dir(folderPath, vbDirectory) <> vbNullString
End Function

Public Function CreateFolder(ByVal folderPath As String) As String
Attribute CreateFolder.VB_Description = "recursively creates all folders in the given path"
    '--- recursively creates all folders in the given path
    '    from:  http://www.freevbcode.com/ShowCode.asp?ID=257
    On Error GoTo ErrorHandler
    Dim s As String
    s = GetPathOnly(folderPath)
    If Dir(s, vbDirectory) = vbNullString Then
        s = CreateFolder(s)
        If Len(s) > 0 Then
            MkDir s
        End If
    End If
    CreateFolder = folderPath
    Exit Function

ErrorHandler:
    Exit Function
End Function

Public Function GetPathOnly(ByVal folderPath As String) As String
Attribute GetPathOnly.VB_Description = "returns the directory path portion of the given fully qualified pathname"
    '--- returns the directory path portion of the given fully qualified
    '    pathname
    '    from:  http://www.freevbcode.com/ShowCode.asp?ID=257
    GetPathOnly = left$(folderPath, InStrRev(folderPath, "\", Len(folderPath)) - 1)
End Function

Public Function IsFileOpen(ByVal Filename As String) As Boolean
Attribute IsFileOpen.VB_Description = "checks if the given file is already open by another application"
    '--- checks if the given file is already open by another application
    '    https://stackoverflow.com/a/9373914/4717755
    Dim ff As Long
    Dim ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open Filename For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsFileOpen = False
    Case 70:   IsFileOpen = True    'permission denied
    Case Else: Err.Raise ErrNo
    End Select
End Function

Public Sub DebugClearImmediateWindow()
Attribute DebugClearImmediateWindow.VB_Description = "Convenience method to clear the debugger immediate window"
    '--- Convenience method to clear the debugger immediate window
    Application.SendKeys "^g ^a {DEL}"
End Sub

Public Sub CopyTextToClipboard(ByVal text As String)
Attribute CopyTextToClipboard.VB_Description = "Copies the given text to the system clipboard"
    '--- from: https://stackoverflow.com/a/25336423/4717755
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub

Public Function GetNetworkPath(ByVal driveName As String) As String
Attribute GetNetworkPath.VB_Description = "Converts a drive letter, e.g. 'W:', to its fully qualified network path. Useful for saving a network folder location without any user-specific custom mapping"
    '--- Converts a drive letter, e.g. 'W:\', to its fully qualified network
    '    path useful for saving a network folder location without any user-
    '    specific custom mapping
    '--- from https://www.mrexcel.com/forum/excel-questions/
    '            658830-show-full-file-paths-unc-not-mapped-drive-letters-print.html
    Dim networkObject  As Object
    Dim networkDrives As Object
    
    Set networkObject = CreateObject("WScript.Network")
    Set networkDrives = networkObject.enumnetworkdrives
    
    Dim i As Long
    GetNetworkPath = vbNullString
    For i = 0 To networkDrives.Count - 1 Step 2
        If UCase$(networkDrives.Item(i)) = UCase$(driveName) Then
            GetNetworkPath = networkDrives.Item(i + 1)
            Exit For
        End If
    Next
End Function

Public Function SelectFile(ByVal thisTitle As String, _
                           ByVal thisFilterDescription As String, _
                           ByVal thisFilter As String, _
                           Optional ByVal thisPath As String = vbNullString) As String
    '--- NOTE: this function works within Excel, but doesn't work in MS Project
    '          if you need this within Project, you'll have to open an instance
    '          of Excel or Word and change Application.FileDialog to xlApp.FileDialog
    '          (where Dim xlApp As Excel.Application; Set xlApp = AttachToExcelApplication())
    SelectFile = vbNullString
    Dim filePicker As Office.FileDialog
    Set filePicker = Application.FileDialog(MsoFileDialogType.msoFileDialogOpen)
    With filePicker
        .Title = thisTitle
        .Filters.Clear
        .Filters.Add thisFilterDescription, thisFilter, 1
        .AllowMultiSelect = False
        If thisPath <> vbNullString Then
            .InitialFileName = thisPath
        Else
            .InitialFileName = ThisWorkbook.path
        End If
        If .Show Then
            '--- if the selected file is located on a network path, that path may be
            '    mapped to a user-specified local drive letter. in order to maintain
            '    access to the file for other users, we have to convert that mapped
            '    drive letter to the network path. any files on a user's local drive
            '    will obviously not be mapped and will not be accessible by other users
            Dim selectedProjectFilepath As String
            selectedProjectFilepath = .SelectedItems(1)
            If Mid$(selectedProjectFilepath, 2, 1) = ":" Then
                '--- this means the first two characters are the drive letter
                Dim networkPath As String
                networkPath = GetNetworkPath(left$(selectedProjectFilepath, 2))
                If Len(networkPath) = 0 Then
                    '--- the drive letter is not mapped or is a local drive,
                    '    so there's nothing to do
                Else
                    selectedProjectFilepath = networkPath & "\" & _
                                              Right$(selectedProjectFilepath, _
                                                     Len(selectedProjectFilepath) - 2)
                End If
            End If
            SelectFile = selectedProjectFilepath
        End If
    End With
End Function

Public Function PadLeft(ByRef text As String, _
                        ByVal totalLength As Long, _
                        Optional ByVal padCharacter As String = " ") As String
Attribute PadLeft.VB_Description = "adds padding to the left of the given string"
    '--- adds padding to the left of the given string
    '    from: https://stackoverflow.com/a/12060429/4717755
    If totalLength <= 0 Then
        Err.Raise Number:=9, Source:="PadLeft Function", _
                                       Description:="illegal length: totalLength <= 0"
    End If
    If Len(CStr(text)) >= totalLength Then
        PadLeft = left$(CStr(text), totalLength)
    Else
        PadLeft = String(totalLength - Len(CStr(text)), padCharacter) & CStr(text)
    End If
End Function

Public Function PadRight(ByVal text As String, _
                  ByVal totalLength As Long, _
                  Optional ByVal padCharacter As String = " ") As String
Attribute PadRight.VB_Description = "adds padding to the right of the given string"
    '--- adds padding to the right of the given string
    '    from: https://stackoverflow.com/a/12060429/4717755
    If totalLength <= 0 Then
        Err.Raise Number:=9, Source:="PadRight Function", _
                                       Description:="illegal length: totalLength <= 0"
    End If
    If Len(CStr(text)) >= totalLength Then
        PadRight = Right$(CStr(text), totalLength)
    Else
        PadRight = CStr(text) & String(totalLength - Len(CStr(text)), padCharacter)
    End If
End Function

Public Function PadCenter(ByVal text As String, _
                          ByVal totalLength As Long, _
                          Optional ByVal padCharacter As String = " ") As String
Attribute PadCenter.VB_Description = "adds padding to the center the given string"
    If totalLength <= 0 Then
        Err.Raise Number:=9, Source:="PadCenter Function", _
                                       Description:="illegal length: totalLength <= 0"
    End If
    '--- adds padding to the center the given string
    Dim leftPad As Long
    Dim rightPad As Long
    If Len(CStr(text)) > totalLength Then
        '--- the original text is longer than the requested length,
        '    so truncate the text to that length
        PadCenter = left$(CStr(text), totalLength)
    Else
        leftPad = (totalLength / 2) - (Len(CStr(text)) / 2)
        rightPad = totalLength - Len(CStr(text)) - leftPad
        PadCenter = String(leftPad, padCharacter) & text & String(rightPad, padCharacter)
    End If
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
    Dim numValues As Variant
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
                                    Dim N As Long
                                    For N = LBound(array1, 5) To UBound(array1, 5)
                                        If numDimensions1 = 5 Then
                                            If array1(i, j, k, m, N) <> array2(i, j, k, m, N) Then
                                                ArraysMatch = False
                                                Exit Function
                                            End If
                                        Else
                                        End If
                                    Next N
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
