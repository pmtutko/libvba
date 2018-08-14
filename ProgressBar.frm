VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Consolidating Project Resources"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   2370
   ClientWidth     =   8205
   OleObjectBlob   =   "ProgressBar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'--------------------------------------------------------------------------
' in order for the "Stop" button to work, the calling module must set up
' a loop with "DoEvents" and monitor the (object).UserRequestedStop flag
'--------------------------------------------------------------------------


'--------------------------------------------------------------------------
'Public Variables and Objects
'--------------------------------------------------------------------------

'----- direct access to labels:
'  ProjectNumber    - user formatted to "1 of 3"
'  ProjectName      - user formatted string
'  ProjectTasks     - user formatted integer
'  ProjectResources - user formatted integer
'  ProjectStart     - user formatted date "5-Jul-2018"
'  ProjectFinish    - user formatted date "5-Jul-2018"
'  CaptureStart     - user formatted date "5-Jul-2018"
'  CatpureFinish    - user formatted date "5-Jul-2018"

'--------------------------------------------------------------------------
'Private Variables and Objects
'--------------------------------------------------------------------------
'Private variables for storing the Main Properties of the Object
'Assigned to Caption Property of the Form.
Private cFormTitle As String
'Set to True if you'd like the Excel Status bar to also Show status.
Private cExcelStatusBar As Boolean
'Used to set to total number of actions the user wishes to perform.
Private cTotalActions As Long
'Remembers the number of completed actions. This variable can be overriddin
'by setting a number to the ActionNumber Property.
Private cActionNumber As Long
'Sets a "skipping" value to allow updating the progress bar less often
'if the number of total actions is huge. **MUST** be changed by the caller
'in a property
Private cSkipUpdates As Long
'Remembers the message that the user wishes to display in the Progressbar's
'Status bar.
Private cStatusMessage As String
'Remembers the Initial Width of the Progressbar Label.
Private cBarWidth As Double
'Remembers the Percentage of Actions Completed. Assigned to the Percent
'indicator in the ProgressBar form.
Private cPercentComplete As String

'Private Variables to Remember if certain properties were set
Private cFormShowStatus As Boolean
Private cTotalActionsSet As Boolean

'Private Variables to Handle the Colour Changes
Private cStartColourSet As Boolean
Private cEndColourSet As Boolean
Private cChangeColours As Boolean
Private cStartColour As XlRgbColor
Private cEndColour As XlRgbColor
Private cStartRed As Long, cEndRed As Long
Private cStartGreen As Long, cEndGreen As Long
Private cStartBlue As Long, cEndBlue As Long

'Private Variables for User Control
Private cUserRequestedStop As Boolean
Private cShowTimeRemaining As Boolean
Private cTimeStarted As Double
Private cLastStepTimer As Double
Private cLastStepNumber As Long
Private cTimeInterval As Double
Const MAX_INTERVALS As Long = 50
Private intervalTimes(1 To MAX_INTERVALS) As Double

'--------------------------------------------------------------------------

'--------------------------------------------------------------------------
'Error Numbers and Description:
'--------------------------------------------------------------------------
' 1 - Set this Property before Running the Show Method.
' 2 - Set TotalActions Property First.
' 3 - Current Action number is greater than Total Actions.
' 4 - TotalActions cannot be changed after it has been set.
' 5 - Run the Show Method First
' 6 - Run the Complete Method only after all the actions have been completed.
' 7 - Progress Bar has already been Loaded.
' 8 - Set StartColour First.
'--------------------------------------------------------------------------

'##########################################################################
'Class Events
'##########################################################################

'This Procedure is Run when the Form in Initiated
Private Sub UserForm_Initialize()
    'Set Default Values for all the Variables
    cActionNumber = 0
    cTotalActions = 0
    cStatusMessage = "Ready"
    cFormTitle = "Progress Bar"
    cExcelStatusBar = False
    cFormShowStatus = False
    cTotalActionsSet = False
    cPercentComplete = "0%"
    cStartColourSet = False
    cEndColourSet = False
    cStartColour = rgbDodgerBlue
    cChangeColours = False
    cUserRequestedStop = False
    cSkipUpdates = 1
    cShowTimeRemaining = True

    With Me
        .Title = cFormTitle
        .StatusMessageBox.Caption = " " & cStatusMessage
        .PercentIndicator.Caption = cPercentComplete
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '--- for this userform, same as clicking the Stop Button
    StopButton_Click
End Sub

'This Sub is run when the class is terminated
Private Sub UserForm_Terminate()
    Application.StatusBar = False
    If Not Me.PercentIndicator = "100%" Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        End
    End If
End Sub

'##########################################################################
'Properties
'##########################################################################

Public Property Get UserRequestedStop() As Boolean
    '--- only set to True if the user clicks the "Stop" button
    DoEvents
    UserRequestedStop = cUserRequestedStop
End Property

'This Procedure is Executed when the Title Propoerty is Set
Public Property Let Title(value As String)

    'Proceed only if the Form has not been loaded already
    If cFormShowStatus Then
        Err.Raise 1, cFormTitle, "Set this Property before Running the Show Method."
    Else
        'Proceed if the user did not send a blank string
        If Not value = vbNullString Then
            'Initialize the private class variable
            cFormTitle = value
            'Update the Form's title if it has already been loaded
            If Not Me Is Nothing Then
                'Do Events makes sure the rest of your macro keeps running
                DoEvents
                Me.Caption = cFormTitle
            End If
        End If
    End If

End Property

'This Procedure lets the User tries to access the Title Property.
Public Property Get Title() As String
    Title = cFormTitle
End Property

'This Procedure is Executed when the ExcelStatusBar Propoerty is Set.
Public Property Let ExcelStatusBar(value As Boolean)

    'Proceed only if the Form has not been loaded already
    If cFormShowStatus Then
        Err.Raise 1, cFormTitle, "Set this Property before Running the Show Method."
    Else
        'Initialize the private class variable
        cExcelStatusBar = value
        'If the user wants to see the Status messages in Excel's
        'Status bar also, make sure it is displayed.
        If value Then
            Application.DisplayStatusBar = True
        End If
    End If
End Property

'This Procedure lets the User tries to access the ExcelStatusBar Property.
Public Property Get ExcelStatusBar() As Boolean
    ExcelStatusBar = cExcelStatusBar
End Property

'This Procedure is Executed when the TotalActions Propoerty is Set.
Public Property Let TotalActions(value As Long)

    'Proceed only if the Form has not been loaded already
    If cFormShowStatus Then
        Err.Raise 1, cFormTitle, "Set this Property before Running the Show Method."
    Else
        'Proceed if the User has not set the TotalActions property
        'Else Display an error message
        If cTotalActionsSet Then
            Err.Raise 4, cFormTitle, "TotalActions cannot be changed after it has been set."
        Else
            'Initialize the private class variables
            cTotalActions = value
            'This is used to make sure the user does not change this later
            cTotalActionsSet = True
        End If
    End If
   
End Property

'This Procedure lets the User tries to access the ActionNumber Property.
Public Property Get TotalActions() As Long
    TotalActions = cTotalActions
End Property

'This Procedure is Executed when the StartColour Propoerty is Set.
Public Property Let StartColour(value As XlRgbColor)

    'Proceed only if the Form has not been loaded already
    If cFormShowStatus Then
        Err.Raise 1, cFormTitle, "Set this Property before Running the Show Method."
    Else
        cStartColourSet = True
        cStartColour = value
        cStartRed = GetPrimaryColour(cStartColour, "R")
        cStartGreen = GetPrimaryColour(cStartColour, "G")
        cStartBlue = GetPrimaryColour(cStartColour, "B")
    End If

End Property

'This Procedure is Executed when the EndColour Propoerty is Set.
Public Property Let EndColour(value As XlRgbColor)

    'Proceed only if the Form has not been loaded already
    If cFormShowStatus Then
        Err.Raise 1, cFormTitle, "Set this Property before Running the Show Method."
    Else
        If Not cStartColourSet Then
            Err.Raise 8, cFormTitle, "Set StartColour First."
        Else
            cEndColourSet = True
            cEndColour = value
            cEndRed = GetPrimaryColour(cEndColour, "R")
            cEndGreen = GetPrimaryColour(cEndColour, "G")
            cEndBlue = GetPrimaryColour(cEndColour, "B")
            cChangeColours = Not CBool(cStartColour = cEndColour)
        End If
    End If

End Property

'This Procedure is Executed when the ActionNumber Propoerty is Set.
Public Property Let ActionNumber(value As Long)

    'Update the private class variable
    cActionNumber = value

    'Call the Sub that Checks if the inputs are valid
    'and refreshes the Progressbar
    UpdateTheBar

End Property

'This Procedure lets the User tries to access the ActionNumber Property.
Public Property Get ActionNumber() As Long
    ActionNumber = cActionNumber
End Property

'This Procedure is Executed when the ProgressStatusMessage Propoerty is Set.
Public Property Let StatusMessage(value As String)

    'Update the private class variable
    cStatusMessage = value

    'Call the Sub that Checks if the inputs are valid
    'and refreshes the Progressbar
    UpdateTheBar

End Property

'This Procedure lets the User tries to access the ProgressStatusMessage Property.
Public Property Get StatusMessage() As String
    StatusMessage = cStatusMessage
End Property

Public Property Let SkipUpdatesEvery(ByVal newSkipValue As Long)
    cSkipUpdates = newSkipValue
End Property

Public Property Get SkipUpdatesEvery() As Long
    SkipUpdatesEvery = cSkipUpdates
End Property

Public Property Let ShowTimeRemaining(ByVal newState As Boolean)
    With Me
        .TimeRemainingLabel.Visible = newState
        .TimeRemainingTitle.Visible = newState
        cShowTimeRemaining = newState
    End With
End Property

Public Property Get ShowTimeRemaining() As Boolean
    ShowTimeRemaining = cShowTimeRemaining
End Property

'##########################################################################
'Public Methods
'##########################################################################

Public Sub ShowBar()
    If cFormShowStatus Then
        Err.Raise 7, cFormTitle, "Progress Bar has already been Loaded."
    Else
        'Do Events makes sure the rest of your macro keeps running
        DoEvents
        'Remember the Initial Width of the ProgressBar
        cBarWidth = Me.ProgressBar.Width
        'Set the Width of the Progressbar to Zero
        Me.ProgressBar.Width = 0
        'Update the Title of the Form
        Me.Caption = cFormTitle
        'Initialize the Private Class Variable
        cFormShowStatus = True
        'Change the Colour of the Progressbar to StartColour
        Me.ProgressBar.BackColor = cStartColour
        Me.ProgressBox.BorderColor = cStartColour
        'Show the Form
        Me.Show
        'Repaint the Form
        Me.Repaint
        '--- capture the start time for the time remaining label
        cTimeStarted = CDbl(Timer)
        cLastStepTimer = cTimeStarted
    End If
End Sub

'NextAction:Let the Progressbar know an action has been completed. I recommend using this
'method over manually overriding the ProgressStatusMessage and CurrentAction Properties
Public Sub NextAction(Optional ByVal ProgressStatusMessage As String, _
                      Optional ByVal ShowActionCount As Boolean = True)

    cActionNumber = cActionNumber + 1
    If ShowActionCount Then
        cStatusMessage = "Action " & cActionNumber & " of " _
                       & cTotalActions & " | " & ProgressStatusMessage
    Else
        cStatusMessage = ProgressStatusMessage
    End If

    UpdateTheBar
    
End Sub

'Complete: This Method Can be used to Let the User Know that the Run Has
'been Completed. It Changes the statues message to the message specified
'and releases the control this object has over the excel status bar.
Public Sub Complete(Optional ByVal WaitForSeconds As Long = 3, _
                    Optional ByVal Prompt As String = "Complete")
    
    Dim Counter As Long

    'Proceed if the ProgressBar has already been loaded
    'Display an error message otherwise
    If cFormShowStatus Then
        'Display an error message if the CurrentAction numeber is lesser than the Number of
        'TotalActions. Otherwise, Change the Status Message to the desired message.
        If cActionNumber < cTotalActions Then
            Err.Raise 6, cFormTitle, _
                      "Run the Complete Method only after all the actions have been completed."
        Else
            'Release control over Excel's status bar
            If cExcelStatusBar Then
                Application.StatusBar = False
            End If

            If WaitForSeconds > 0 Then
                For Counter = WaitForSeconds To 1 Step -1
                    DoEvents
                    Me.StatusMessageBox.Caption = " " & Prompt & _
                                                  " | This Window will close in " & Counter & " " & _
                                                  IIf(Counter = 1, "second.", "seconds.")
                    Application.Wait (Now() + TimeValue("00:00:01"))
                Next Counter
                Terminate
            Else
                'Do Events makes sure the rest of your macro keeps running
                DoEvents
                'Update the Status Message
                Me.StatusMessageBox.Caption = " " & Prompt
            End If
            
        End If
    Else
        Err.Raise 5, cFormTitle, "Run the Show Method First"
    End If
End Sub

'Terminate: Let the user Terminate Manually if they prefer
Public Sub Terminate()
    'Terminate the Form if it is already loaded. Display an Error Message
    'Otherwise
    If cFormShowStatus Then
        'If the Form is Loaded, Unload it
        If cFormShowStatus Then
            Me.Hide
            cFormShowStatus = False
            cTotalActionsSet = False
            cActionNumber = 0
            cTotalActions = 0
        End If
        'Return the Appliation StatusBar control to Excel
        If cExcelStatusBar Then
            Application.StatusBar = False
        End If
    Else
        Err.Raise 5, cFormTitle, "Run the Show Method First"
    End If
End Sub

'##########################################################################
Private Sub StopButton_Click()
    cUserRequestedStop = True
End Sub

'##########################################################################
'Private Subs needed by this Class
'##########################################################################
Private Sub UpdateTheBar()

    'Proceed only if the user has already set the TotalActions property
    'Else Display an error message
    If cTotalActionsSet Then
        'Proceed only if the CurrentAction number is lesser than or equal to
        'the Total Actions
        If cActionNumber > cTotalActions Then
            Err.Raise 3, cFormTitle, _
                      "Current Action number is greater than Total Actions."
        Else
            'Proceed only if the Form has already been Showed. Display an
            'error message otherwise
            If cFormShowStatus Then
                'Call the Procedure that Updates the Bar
                UpdateProgress
            Else
                Err.Raise 5, cFormTitle, "Run the Show Method First"
            End If
        End If
    Else
        Err.Raise 2, cFormTitle, "Set TotalActions Property First."
    End If

End Sub

Private Sub UpdateProgress()

    '--- check to see if the modulo works to update the bar this time
    If (cActionNumber Mod cSkipUpdates) = 0 Then
        'Declare Sub Level Variables
        Dim FractionComplete As Double
        Dim ProgressPercent As String
        Dim BarWidth As Double
        Dim BarColour As XlRgbColor
    
        'Initialize Variables
        FractionComplete = cActionNumber / cTotalActions
        BarWidth = cBarWidth * FractionComplete
        cPercentComplete = Format$(FractionComplete * 100, "0") & "%"
    
        'Do Events makes sure the rest of your macro keeps running
        DoEvents
        'Change the Width of the Label
        Me.ProgressBar.Width = BarWidth
        'Update the Percent Indicator
        Me.PercentIndicator.Caption = cPercentComplete
    
        'Change the Colour of the Progressbar if needed
        If cChangeColours Then
            BarColour = RGB( _
                        cStartRed + (cEndRed - cStartRed) * FractionComplete, _
                        cStartGreen + (cEndGreen - cStartGreen) * FractionComplete, _
                        cStartBlue + (cEndBlue - cStartBlue) * FractionComplete)
            Me.ProgressBar.BackColor = BarColour
        End If
    
        'Set the Status Bar Message
        Me.StatusMessageBox.Caption = " " & cStatusMessage
        'Update Excel's Status Bar if needed
        If cExcelStatusBar Then
            Application.StatusBar = ProgressText(cActionNumber, cTotalActions) & _
                                    " | " & cPercentComplete & " | " & cStatusMessage
        End If
        
        '--- always update the time remaining, but only update
        '    the label if it's enabled and visible
        Dim stepElapsedTime As Double
        Dim stepsElapsed As Long
        Dim secondsPerStep As Double
        stepElapsedTime = CDbl(Timer) - cLastStepTimer
        stepsElapsed = cActionNumber - cLastStepNumber
        secondsPerStep = stepElapsedTime / stepsElapsed

        Const SMOOTHING_FACTOR As Double = 0.005
        Dim stepsRemaining As Long
        Dim timeRemaining As Double
        Dim averageSpeed As Double
        Dim minsLeft As Long
        Dim secsLeft As Long
        secondsPerStep = Round(AverageSecondsPerStep, 4)
        averageSpeed = SMOOTHING_FACTOR * secondsPerStep + _
                       (1 - SMOOTHING_FACTOR) * averageSpeed
        stepsRemaining = cTotalActions - cActionNumber
        timeRemaining = stepsRemaining * secondsPerStep
        minsLeft = Fix(timeRemaining / 60#)
        secsLeft = Fix(timeRemaining - Fix(timeRemaining / 60#) * 60#)
        If cShowTimeRemaining Then
            Me.TimeRemainingLabel = Format$(minsLeft, "#0") & " minutes " & _
                                    Format$(secsLeft, "#0") & " seconds"
        End If
        cLastStepNumber = cActionNumber
        cLastStepTimer = CDbl(Timer)
        
        'Repaint the Form
        Me.Repaint
    End If
End Sub
'##########################################################################

'##########################################################################
'Private Functions Needed by this Class Module
'##########################################################################
Private Function AverageSecondsPerStep() As Double
    Static index As Long
    
    If cTimeStarted = cLastStepTimer Then
        '--- condition to initialize the interval array
        For index = 1 To MAX_INTERVALS
            intervalTimes(index) = 0#
        Next index
        index = 1
    End If
    
    Dim stepElapsedTime As Double
    Dim stepsElapsed As Long
    Dim secondsPerStep As Double
    stepElapsedTime = CDbl(Timer) - cLastStepTimer
    stepsElapsed = cActionNumber - cLastStepNumber
    secondsPerStep = stepElapsedTime / stepsElapsed
    
    '--- store it in the next slot
    '    the array is used to store the last N interval time
    '    set up as a circular buffer
    intervalTimes(index) = secondsPerStep
    If index = MAX_INTERVALS Then
        index = 1
    Else
        index = index + 1
    End If
        
    '--- capture for next time
    cLastStepNumber = cActionNumber
    cLastStepTimer = CDbl(Timer)
    
    AverageSecondsPerStep = Application.WorksheetFunction.Average(intervalTimes)
End Function

'##########################################################################
'GetPrimaryColour: Function used to extract the numerical value of each
    'of the promary colours in a colour. Every colour is a mixture of three
    'primary colours: Red, Blue and Green. In Excel, colours are represented
    'as Hexadecimal Numbers of the Format "00BBGGRR", but they are stored as
    'Decimal numbers. This function converts the Long Variable into Hexadicamal
    'first, and then extracts the two characters that represent the primary
    'colour and then converts it back to long.
    '##########################################################################
    'Author     : Ejaz Ahmed
    'Email      : StrugglingToExcel@outlook.com
    'Website    : http://strugglingtoexcel.wordpress.com/
    'Date       : 26 August 2014
    '##########################################################################
Private Function GetPrimaryColour(ByVal WhichColour As XlRgbColor, _
                                  ByVal RedBlueGreen As String) As Long
   
    'Declate Function Level Variables
    Dim HexString As String

    'Convert Decimal to HexaDecimal
    HexString = CStr(Hex$(WhichColour))
    'Prefix 0's so the string is always 8 Characters in length
    HexString = String(8 - Len(HexString), "0") & HexString

    'Extract the Red Blue or Green part of the Hexadecimal Number
    'stored in HexString. Remember, that we need to prefix "&H"
    'to tell excel that the number is Hexadecimal, so we can use
    'the cLng function to convert it into the Decimal System Later.
    Select Case StrConv(RedBlueGreen, vbUpperCase)
    Case "R"
        HexString = "&H" & Mid$(HexString, 7, 2)
    Case "G"
        HexString = "&H" & Mid$(HexString, 5, 2)
    Case "B"
        HexString = "&H" & Mid$(HexString, 3, 2)
    Case Else
        HexString = "-100"
    End Select

    'Return the colour value in the 0 to 225 format.
    GetPrimaryColour = CLng(HexString)
    
End Function

'##########################################################################
'ProgressText: Function to generate the Text, using unicode characters, to
    'show a progress bar on Excel's Status Bar
    '##########################################################################
    'Author     : Ejaz Ahmed
    'Email      : StrugglingToExcel@outlook.com
    'Website    : http://strugglingtoexcel.wordpress.com/
    'Date       : 22 September 2014
    '##########################################################################
Private Function ProgressText(ByVal ActionNumber As Long, _
                      ByVal TotalActions As Long, _
                      Optional ByVal BarLength As Long = 15) As String
     
    Dim BarComplete As Long
    Dim BarInComplete As Long
    Dim BarChar As String
    Dim SpaceChar As String
    Dim TempString As String

    BarChar = ChrW$(&H2589)
    SpaceChar = ChrW$(&H2000)

    BarLength = Round(BarLength / 2, 0) * 2
    BarComplete = Fix((ActionNumber * BarLength) / TotalActions)
    BarInComplete = BarLength - BarComplete
    ProgressText = String(BarComplete, BarChar) & String(BarInComplete, SpaceChar)
  
End Function



