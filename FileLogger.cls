VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Libraries")
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetTempPath Lib "kernel32" _
                             Alias "GetTempPathA" _
                             (ByVal nBufferLength As Long, _
                              ByVal lpBuffer As String) As Long
    Private Declare PtrSafe Function GetTempFileName Lib "kernel32" _
                             Alias "GetTempFileNameA" _
                             (ByVal lpszPath As String, _
                              ByVal lpPrefixString As String, _
                              ByVal wUnique As Long, _
                              ByVal lpTempFileName As String) As Long
#Else
    Private Declare Function GetTempPath Lib "kernel32" _
                             Alias "GetTempPathA" _
                             (ByVal nBufferLength As Long, _
                              ByVal lpBuffer As String) As Long
    Private Declare Function GetTempFileName Lib "kernel32" _
                             Alias "GetTempFileNameA" _
                             (ByVal lpszPath As String, _
                              ByVal lpPrefixString As String, _
                              ByVal wUnique As Long, _
                              ByVal lpTempFileName As String) As Long
#End If

'------------------------------------------------------------------------------
'  Class Internal Data
'------------------------------------------------------------------------------
Private Type InternalData
    WriteFileOnTermination As Boolean
    TextShouldBeAppended As Boolean
    AutoAddCrLf As Boolean
    PrependTimestamp As Boolean
    Filename As String
    Buffer As String
End Type

Private this As InternalData

'------------------------------------------------------------------------------
'  Class Properties
'------------------------------------------------------------------------------
Public Property Let Filename(ByVal newName As String)
    this.Filename = newName
End Property

Public Property Get Filename() As String
    Filename = this.Filename
End Property

Public Property Let Buffer(ByVal newBuffer As String)
    this.Buffer = newBuffer
End Property

Public Property Get Buffer() As String
    Buffer = this.Buffer
End Property

Public Sub ClearBuffer()
    this.Buffer = vbNullString
End Sub

Public Property Let AutoAddCrLf(ByVal newFlag As Boolean)
    this.AutoAddCrLf = newFlag
End Property

Public Property Get AutoAddCrLf() As Boolean
    AutoAddCrLf = this.AutoAddCrLf
End Property

Public Property Let WriteFileOnTermination(ByVal newFlag As Boolean)
    this.WriteFileOnTermination = newFlag
End Property

Public Property Get WriteFileOnTermination() As Boolean
    WriteFileOnTermination = this.WriteFileOnTermination
End Property

Public Property Let PrependTimestamp(ByVal newFlag As Boolean)
    this.PrependTimestamp = newFlag
End Property

Public Property Get PrependTimestamp() As Boolean
    PrependTimestamp = this.PrependTimestamp
End Property

'------------------------------------------------------------------------------
'  Class Public Methods
'------------------------------------------------------------------------------
Public Sub Append(ByVal moreText As String, _
                  Optional ByVal AutoAddCrLf As Variant, _
                  Optional ByVal PrependTimestamp As Variant)
    '--- the optional parameter is a Boolean and will override the class
    '    setting if present
    If Not IsMissing(PrependTimestamp) Then
        If PrependTimestamp Then
            this.Buffer = this.Buffer & Format$(Now(), "dd-mmm-yyyy ") & _
                                  TimeInMS() & ": " & moreText
        Else
            this.Buffer = this.Buffer & moreText
        End If
    ElseIf this.PrependTimestamp Then
        this.Buffer = this.Buffer & Format$(Now(), "dd-mmm-yyyy ") & _
                      TimeInMS() & ": " & moreText
    Else
        this.Buffer = this.Buffer & moreText
    End If
    
    '--- the optional parameter overrides the class setting
    If Not IsMissing(AutoAddCrLf) Then
        If AutoAddCrLf Then
            this.Buffer = this.Buffer & vbCrLf
        End If
    ElseIf this.AutoAddCrLf Then
        this.Buffer = this.Buffer & vbCrLf
    End If
End Sub

Public Sub WriteToFile()
    '--- opens the given filename and writes the entire contents of the
    '    current buffer to the file, then closes the file
    If (Len(this.Filename) > 0) And (Len(this.Buffer) > 0) Then
        Dim freeFileNumber As Long
        
        freeFileNumber = FreeFile
        If this.TextShouldBeAppended Then
            Open this.Filename For Append As #freeFileNumber
        Else
            Open this.Filename For Output As #freeFileNumber
        End If
    
        Print #freeFileNumber, this.Buffer
        Close #freeFileNumber
        
    Else
        Err.Raise vbObject + 600, Source:="In FileLogger:WriteBuffer", _
                  Description:="No file and path specified."
    End If
End Sub

Public Function CreateTempFileName(Optional prefix As String = "VBA", _
                                   Optional extension As String = "", _
                                   Optional setAsObjectFilename As Boolean = True) As String
    Dim tempPath As String * 512
    Dim tempName As String * 576
    Dim returnValue As Long
    Dim tempFilename As String
    
    CreateTempFileName = vbNullString
    
    returnValue = GetTempPath(512, tempPath)
    If (returnValue > 0 And returnValue < 512) Then
        returnValue = GetTempFileName(tempPath, prefix, 0, tempName)
        If returnValue <> 0 Then
            tempFilename = left$(tempName, InStr(tempName, vbNullChar) - 1)
        End If
        If extension > "" Then
            Kill tempFilename
            If Right(tempFilename, 4) = ".tmp" Then
                tempFilename = left$(tempFilename, Len(tempFilename) - 4)
            End If
            tempFilename = tempFilename & extension
        End If
        If setAsObjectFilename Then
            this.Filename = tempFilename
        End If
        CreateTempFileName = tempFilename
    End If
End Function

'------------------------------------------------------------------------------
'  Class Private Methods
'------------------------------------------------------------------------------
Private Function TimeInMS() As String
    TimeInMS = Format$(Now, "HH:nn:ss") & "." & Right$(Format$(Timer, "#0.00"), 2)
End Function

Private Sub Class_Initialize()
    this.TextShouldBeAppended = False
    this.PrependTimestamp = True
    this.AutoAddCrLf = False
    this.WriteFileOnTermination = True
End Sub

Private Sub Class_Terminate()
    If this.WriteFileOnTermination Then
        WriteToFile
    End If
End Sub




