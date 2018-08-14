Attribute VB_Name = "Lib_PerformanceSupport"
Attribute VB_Description = "Methods to control disabling/enabling of the Application level screen updates. Supports call nesting and debug messaging, plus high precision timer calls."
Option Explicit

'------------------------------------------------------------------------------
' For Update methods
'
Private Type SavedState
    screenUpdate As Boolean
    calculationType As XlCalculation
    eventsFlag As Boolean
    callCounter As Long
End Type

Private previousState As SavedState

Private Const DEBUG_MODE As Boolean = False 'COMPILE TIME ONLY!!

'------------------------------------------------------------------------------
' For Precision Counter methods
'
Private Type LargeInteger
    lowpart As Long
    highpart As Long
End Type

Private Declare Function QueryPerformanceCounter Lib _
                         "kernel32" (lpPerformanceCount As LargeInteger) As Long
Private Declare Function QueryPerformanceFrequency Lib _
                         "kernel32" (lpFrequency As LargeInteger) As Long

Private counterStart As LargeInteger
Private counterEnd As LargeInteger
Private crFrequency As Double

Private Const TWO_32 = 4294967296#               ' = 256# * 256# * 256# * 256#

'==============================================================================
' Screen and Event Update Controls
'
Public Sub ReportUpdateState()
Attribute ReportUpdateState.VB_Description = "Prints to the immediate window the current state and values of the Application update controls."
    Debug.Print ":::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Debug.Print "Application.ScreenUpdating      = " & Application.ScreenUpdating
    Debug.Print "Application.Calculation         = " & Application.Calculation
    Debug.Print "Application.EnableEvents        = " & Application.EnableEvents
    Debug.Print "--previousState.screenUpdate    = " & previousState.screenUpdate
    Debug.Print "--previousState.calculationType = " & previousState.calculationType
    Debug.Print "--previousState.eventsFlag      = " & previousState.eventsFlag
    Debug.Print "--previousState.callCounter     = " & previousState.callCounter
    Debug.Print "--DEBUG_MODE is currently " & DEBUG_MODE
End Sub

Public Sub DisableUpdates(Optional debugMsg As String = vbNullString, _
                          Optional forceZero As Boolean = False)
Attribute DisableUpdates.VB_Description = "Disables Application level updates and events and saves their initial state to be restored later. Supports nested calls. Displays debug messages according to the module-global DEBUG_MODE flag."
    With Application
        '--- capture previous state if this is the first time
        If forceZero Or (previousState.callCounter = 0) Then
            previousState.screenUpdate = .ScreenUpdating
            previousState.calculationType = .Calculation
            previousState.eventsFlag = .EnableEvents
            previousState.callCounter = 0
        End If

        '--- now turn it all off and count
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        previousState.callCounter = previousState.callCounter + 1

        '--- optional stuff
        If DEBUG_MODE Then
            Debug.Print "Updates disabled (" & previousState.callCounter & ")";
            If Len(debugMsg) > 0 Then
                Debug.Print debugMsg
            Else
                Debug.Print vbCrLf
            End If
        End If
    End With
End Sub

Public Sub EnableUpdates(Optional debugMsg As String = vbNullString, _
                         Optional forceZero As Boolean = False)
Attribute EnableUpdates.VB_Description = "Restores Application level updates and events to their state, prior to the *first* DisableUpdates call. Supports nested calls. Displays debug messages according to the module-global DEBUG_MODE flag."
    With Application
        '--- countdown!
        If previousState.callCounter >= 1 Then
            previousState.callCounter = previousState.callCounter - 1
        ElseIf forceZero = False Then
            '--- shouldn't get here
            Debug.Print "EnableUpdates ERROR: reached callCounter = 0"
        End If

        '--- only re-enable updates if the counter gets to zero
        '    or we're forcing it
        If forceZero Or (previousState.callCounter = 0) Then
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .EnableEvents = True
        End If

        '--- optional stuff
        If DEBUG_MODE Then
            Debug.Print "Updates enabled (" & previousState.callCounter & ")";
            If Len(debugMsg) > 0 Then
                Debug.Print debugMsg
            Else
                Debug.Print vbCrLf
            End If
        End If
    End With
End Sub

'==============================================================================
' Precision Timer Controls
'
Private Function LI2Double(lgInt As LargeInteger) As Double
Attribute LI2Double.VB_Description = "Converts LARGE_INTEGER to Double"
    '--- converts LARGE_INTEGER to Double
    Dim low As Double
    low = lgInt.lowpart
    If low < 0 Then
        low = low + TWO_32
    End If
    LI2Double = lgInt.highpart * TWO_32 + low
End Function

Public Sub StartCounter()
Attribute StartCounter.VB_Description = "Captures the high precision counter value to use as a starting reference time."
    '--- Captures the high precision counter value to use as a starting
    '    reference time.
    Dim perfFrequency As LargeInteger
    QueryPerformanceFrequency perfFrequency
    crFrequency = LI2Double(perfFrequency)
    QueryPerformanceCounter counterStart
End Sub

Public Function TimeElapsed() As Double
Attribute TimeElapsed.VB_Description = "Returns the time elapsed since the call to StartCounter in microseconds."
    '--- Returns the time elapsed since the call to StartCounter in microseconds
    If crFrequency = 0# Then
        Err.Raise Number:=11, _
                  Description:="Must call 'StartCounter' in order to avoid " & _
                                "divide by zero errors."
    End If
    Dim crStart As Double
    Dim crStop As Double
    QueryPerformanceCounter counterEnd
    crStart = LI2Double(counterStart)
    crStop = LI2Double(counterEnd)
    TimeElapsed = 1000# * (crStop - crStart) / crFrequency
End Function

