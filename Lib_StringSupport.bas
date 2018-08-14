Attribute VB_Name = "Lib_StringSupport"
Attribute VB_Description = "Variety of support functions operating on VBA Strings."
Option Explicit

Public Function CleanString(ByVal inString As String) As String
Attribute CleanString.VB_Description = "Strips out ANY non-printable character and the designated whitespace from the string and returns the result"
    '--- strips out ANY non-printable character and the designated
    '    whitespace from the string and returns the result
    Dim i As Long
    Dim outString As String
    outString = vbNullString
    For i = 1 To Len(inString)
        Dim charValue As Long
        charValue = Asc(Mid$(inString, i, 1))
        If (charValue >= 33) And (charValue <= 122) Then
            outString = outString & Mid$(inString, i, 1)
        End If
    Next i
    CleanString = outString
End Function

Public Sub CopyTextToClipboard(Text As String)
Attribute CopyTextToClipboard.VB_Description = "Copies the given string parameter to the Windows clipboard."
    '--- from: https://stackoverflow.com/a/25336423/4717755
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText Text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub

