VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ScrollingStatusForm 
   Caption         =   "Process Status"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10845
   OleObjectBlob   =   "ScrollingStatusForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ScrollingStatusForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
' Private Internal Data
'------------------------------------------------------------------------------
Enum EntryAction
    [_First] = 1
    AddNew = 1
    Latest = 2
    Existing = 3
    [_Last] = 3
End Enum

'------------------------------------------------------------------------------
' Public Class Properties (READ ONLY)
'------------------------------------------------------------------------------
Public Property Get Count() As Long
    Count = Me.StatusBox.ListCount
End Property

'------------------------------------------------------------------------------
' Public Class Properties (READ/WRITE)
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' Public Class Methods
'------------------------------------------------------------------------------
Public Sub UpdateEntry(action As EntryAction, _
                       Optional msg As String = vbNullString, _
                       Optional statusComplete As Boolean = False, _
                       Optional index As Long = -1)
    '--- the caller MUST specify the type of entry action to be performed
    '    other fields are optional as dictated by the entry action
    With Me.StatusBox
        Dim listIndex As Long
        Select Case action
            Case AddNew
                .AddItem
                listIndex = .ListCount - 1
            Case Latest
                listIndex = .ListCount - 1
            Case Existing
                If (index > -1) And (index < .ListCount) Then
                    listIndex = index
                Else
                    listIndex = .ListCount - 1
                End If
        End Select
        If statusComplete Then
            .List(listIndex, 0) = "Complete"
        Else
            .List(listIndex, 0) = "Working..."
        End If
        If Len(msg) > 0 Then
            .List(listIndex, 1) = msg
        End If
        .TopIndex = .ListCount - 1
    End With
    Me.Repaint
End Sub

'------------------------------------------------------------------------------
' Private Class Methods
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Me.StatusBox.ColumnWidths = "50;120"
    '--- make this userform pop-up centered over the Excel application, even if
    '    it's shown on a secondary display monitor
    With Me
        .StartUpPosition = 0
        .left = Application.left + (0.5 * Application.width) - (0.5 * .width)
        .top = Application.top + (0.5 * Application.height) - (0.5 * .height)
    End With
    '--- now show it
    Me.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Me.Hide
End Sub

Private Sub UserForm_Terminate()
    Me.Hide
End Sub
