Attribute VB_Name = "Lib_MSExcelSupport"
Attribute VB_Description = "Variety of support functions operating on Excel ranges and arrays."
'@Folder("Libraries")
Option Explicit

Public Function IsExcelRunning() As Boolean
Attribute IsExcelRunning.VB_Description = "quick check to see if an instance of MS Excel is running"
    '--- quick check to see if an instance of MS Excel is running
    Dim xlApp As Object
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    IsExcelRunning = True        'assumes it's running
    If Err > 0 Then
        IsExcelRunning = False   'unless it's not running
    End If
End Function

Public Function AttachToExcelApplication() As Excel.Application
Attribute AttachToExcelApplication.VB_Description = "finds an existing and running instance of MS Excel, or starts the application if one is not already running"
    '--- finds an existing and running instance of MS Excel, or starts
    '    the application if one is not already running
    Dim xlApp As Excel.Application
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If Err > 0 Then
        '--- we have to start one
        '    an exception will be raised if the application is not installed
        Set xlApp = CreateObject("Excel.Application")
    End If
    Set AttachToExcelApplication = xlApp
End Function

Public Function GetWorkbook(ByVal sFullName As String, _
                            Optional ByRef wasAlreadyOpen As Boolean) As Workbook
    '--- credit to: https://stackoverflow.com/a/9382034/4717755
    Dim sFile As String
    Dim wbReturn As Workbook

    sFile = Dir(sFullName)

    On Error Resume Next
        Set wbReturn = Workbooks(sFile)

        If wbReturn Is Nothing Then
            Set wbReturn = Workbooks.Open(sFullName)
            wasAlreadyOpen = False
        Else
            wasAlreadyOpen = True
        End If
    On Error GoTo 0

    Set GetWorkbook = wbReturn

End Function

Public Sub DumpDefinedNames(ByRef startInThisCell As Range)
    startInThisCell.ListNames
End Sub


