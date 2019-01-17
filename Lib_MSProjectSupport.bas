Attribute VB_Name = "Lib_MSProjectSupport"
Attribute VB_Description = "Variety of support functions operating on MS Project"
'@Folder("Libraries")
Option Explicit

Public Function IsMSProjectRunning() As Boolean
Attribute IsMSProjectRunning.VB_Description = "quick check to see if an instance of MS Project is running"
    '--- quick check to see if an instance of MS Project is running
    Dim mspApp As Object
    On Error Resume Next
    Set mspApp = GetObject(, "MSProject.Application")
    IsMSProjectRunning = True         'assumes it's running
    If Err > 0 Then
        IsMSProjectRunning = False    'unless it's not running
    End If
End Function

Public Function AttachToMSProjectApplication() As MSProject.Application
Attribute AttachToMSProjectApplication.VB_Description = "finds an existing and running instance of MS Project, or starts the application if one is not already running"
    '--- finds an existing and running instance of MS Project, or starts
    '    the application if one is not already running
    Dim mspApp As MSProject.Application
    On Error Resume Next
    Set mspApp = GetObject(, "MSProject.Application")
    If Err > 0 Then
        '--- we have to start one
        '    an exception will be raised if the application is not installed
        Set mspApp = CreateObject("MSProject.Application")
    End If
    Set AttachToMSProjectApplication = mspApp
End Function

Public Function ProjectGetCustomFieldItems(ByVal fieldId As Long) As Dictionary
Attribute ProjectGetCustomFieldItems.VB_Description = "returns a collection of the lookup items assigned to the given field"
    '--- returns a collection of the lookup items assigned to the given field
    If ProjectCustomFieldHasItems(fieldId) Then
        Dim mspApp As MSProject.Application
        Set mspApp = AttachToMSProjectApplication()
        Dim items As Dictionary
        Set items = New Dictionary
        Dim i As Long
        For i = 1 To 100
            Dim Value As String
            On Error Resume Next
            Value = mspApp.CustomFieldValueListGetItem(fieldId, pjValueListValue, i)
            If Err > 0 Then
                '--- we're done
                If items.Count > 0 Then
                    Set ProjectGetCustomFieldItems = items
                Else
                    Set ProjectGetCustomFieldItems = Nothing
                End If
                Exit Function
            Else
                items.Add Value, Value
            End If
        Next i
    End If
End Function

Public Function ProjectCustomFieldHasItems(ByVal fieldId As Long) As Boolean
Attribute ProjectCustomFieldHasItems.VB_Description = "determines if the given field has any lookup values"
    '--- determines if the given field has any lookup values
    Dim mspApp As MSProject.Application
    Set mspApp = AttachToMSProjectApplication()
    Dim item As Variant
    On Error Resume Next
    item = mspApp.CustomFieldValueListGetItem(fieldId, pjValueListValue, 1)
    If Err > 0 Then
        '--- no items
        ProjectCustomFieldHasItems = False
    Else
        ProjectCustomFieldHasItems = True
    End If
End Function

Public Function ProjectCustomFieldItemCount(ByVal fieldId As Long) As Long
Attribute ProjectCustomFieldItemCount.VB_Description = "determines the number of lookup items in the given field"
    '--- determines the number of lookup items in the given field
    Dim mspApp As MSProject.Application
    Set mspApp = AttachToMSProjectApplication()
    Dim item As Variant
    On Error Resume Next
    Dim i As Long
    i = 0
    For i = 0 To 10000
        item = mspApp.CustomFieldValueListGetItem(fieldId, pjValueListValue, i + 1)
        If Err > 0 Then
            '--- no items
            ProjectCustomFieldItemCount = i
            Exit Function
        End If
    Next i
End Function

