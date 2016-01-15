VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNewMsg 
   Caption         =   "New TOMP Message"
   ClientHeight    =   1365
   ClientLeft      =   17040
   ClientTop       =   10380
   ClientWidth     =   6315
   OleObjectBlob   =   "frmNewMsg.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmNewMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------
'-----------------------------Windows API Calls-------------------------
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
'Gets the handle of the desktop to make the userform independant of outlook
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
'This finds the window in memory
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long

'This resets the parent window
Private Declare Function SetWindowLongA Lib "user32" _
        (ByVal hwnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long

'This defines the x, x and y position of a window
Private Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long
      
Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
      
'This defines a parent window as the desktop
Private Const GWL_HWNDPARENT As Long = (-8&)
Private Const BRDR As Long = (-16)
'**********************************
'These constants can be set into
'the flag for your preference
'**********************************
'do not change the size of the window
Private Const SWP_NOSIZE = &H1
'do not change the position of the window
Private Const SWP_NOMOVE = &H2
'if included this will prevent the activate event from occuring
Private Const SWP_NOACTIVATE = &H10
'if included this will show the form w/o using the activate event
Private Const SWP_SHOWWINDOW = &H40

'set as many of the above into this with the Or keyword
Private Const FLAGS As Long = SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW

'This will pull the window to the top of the "Top Most Windows" z order
'(windows keeps 2 z orders NOTTOPMOST and TOPMOST
Private Const HWND_TOPMOST = -1

'This holds the memory handle for the window
Private hwnd As Long
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
Public item As MailItem


'This method acts as an override for the userform.show method
'it will avoid the form activate event
Public Sub showForm(newItem As MailItem, pos As Variant)
On Error GoTo ErrHandler
logProcedure "|- enter showForm"
    logProcedure "set the global variable for the click events"
    'set the global variable for the click events
    Set item = newItem
    DoEvents
    logProcedure "update the captions of our subject label"
    'update the captions of our lables
    Me.lblSubject.Caption = item.subject
    DoEvents
    logProcedure "update the captions of our sender lable"
    Me.lblSender.Caption = item.sender
    DoEvents
    'set the window handle to this window's handle
    logProcedure "set the window handle to this window's handle"
    hwnd = FindWindow(vbNullString, Me.Caption)
    DoEvents
    logProcedure "Change the window's partent to the desktop"
    'Change the window's partent to the desktop
    SetWindowLongA hwnd, GWL_HWNDPARENT, GetDesktopWindow
    DoEvents
    'variable for the window stype section of memory
    Dim lStyle As Long
    DoEvents
    logProcedure "get that from our current window handle"
    'get that from our current window handle
    lStyle = GetWindowLong(hwnd, -16)
    DoEvents
    logProcedure "reset the variable to include no border"
    'reset the variable to include no border
    lStyle = lStyle And Not &HC800000
    DoEvents
    logProcedure "set the variable in the correct place in memory"
    'set the variable in the correct place in memory
    SetWindowLongA hwnd, -16, lStyle
    DoEvents
    logProcedure "remove any whitespace Height"
    'this will remove any whitespace that the titlebar or border was occupying
    Me.Height = Me.InsideHeight
    DoEvents
    logProcedure "remove any whitespace Width"
    Me.Width = Me.InsideWidth
    DoEvents
    logProcedure "call SetWindowPos"
    'set the z position to the top of the z order
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
    DoEvents
    logProcedure "Call Wait(4)"
    wait (4)
    DoEvents
    logProcedure "Call Unload"
    Unload Me
    DoEvents
logProcedure "exit showForm -|"
Exit Sub
ErrHandler:
    Dim errorNumber As Integer
    errorNumber = Err.number
    Dim errorDescription As String
    errorDescription = Err.description
    logDebug "frmNewMsg.showForm", errorDescription, errorNumber
    logTxt "frmNewMsg.showForm", errorDescription, errorNumber
Resume Next
End Sub

Public Sub setPosition(left As Integer, top As Integer)
On Error GoTo ErrHandler
    Me.left = left
    Me.top = top
Exit Sub
ErrHandler:
    Dim errorNumber As Integer
    errorNumber = Err.number
    Dim errorDescription As String
    errorDescription = Err.description
    logDebug "frmNewMsg.setPosition", errorDescription, errorNumber
    logTxt " frmNewMsg.setPosition", errorDescription, errorNumber
Resume Next
End Sub

'this will display the inspector, intended to be placed in
'click or double click events
Private Sub ShowMsg()
On Error GoTo ErrHandler
    Dim inspector As inspector
    Set inspector = item.GetInspector
    inspector.Display
Exit Sub
ErrHandler:
    Dim errorNumber As Integer
    errorNumber = Err.number
    Dim errorDescription As String
    errorDescription = Err.description
    logDebug "frmNewMsg.setPosition", errorDescription, errorNumber
    logTxt "frmNewMsg.setPosition", errorDescription, errorNumber
Resume Next
End Sub

Public Sub wait(seconds As Integer)
On Error GoTo ErrHandler
    Sleep (seconds * 1000)
Exit Sub
ErrHandler:
    Dim errorNumber As Integer
    errorNumber = Err.number
    Dim errorDescription As String
    errorDescription = Err.description
    logDebug "frmNewMsg.wait", errorDescription, errorNumber
    logTxt "frmNewMsg.wait", errorDescription, errorNumber
Resume Next
End Sub

Private Sub btnComplete_Click()
On Error GoTo ErrHandler
    Dim pos(1) As Variant
    pos(0) = Me.left
    pos(1) = Me.top
    ThisOutlookSession.position = pos
    saveSettings
    Unload Me
Exit Sub
ErrHandler:
    Dim errorNumber As Integer
    errorNumber = Err.number
    Dim errorDescription As String
    errorDescription = Err.description
    logDebug "frmNewMsg.btnComplete_Click", errorDescription, errorNumber
    logTxt "frmNewMsg.btnComplete_Click", errorDescription, errorNumber
Resume Next
End Sub

Private Sub UserForm_Click()
On Error GoTo ErrHandler
    ShowMsg
Exit Sub
ErrHandler:
    logTxt " frmNewMsg.UserForm_Click", Err.description, Err.number
Resume Next
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    ShowMsg
Exit Sub
ErrHandler:
    Dim errorNumber As Integer
    errorNumber = Err.number
    Dim errorDescription As String
    errorDescription = Err.description
    logDebug "frmNewMsg.UserForm_DblClick", errorDescription, errorNumber
    logTxt " frmNewMsg.UserForm_DblClick", errorDescription, errorNumber
Resume Next
End Sub

Public Sub setLabelColor(color As String)
On Error GoTo ErrHandler
    If color = "dark" Then
        lblFro.ForeColor = RGB(255, 255, 255)
        lblSubj.ForeColor = RGB(255, 255, 255)
        lblSender.ForeColor = RGB(255, 255, 255)
        lblSubject.ForeColor = RGB(255, 255, 255)
        lblTitle.ForeColor = RGB(255, 255, 255)
    ElseIf color = "light" Then
        lblFro.ForeColor = RGB(224, 224, 224)
        lblSubj.ForeColor = RGB(224, 224, 224)
        lblSender.ForeColor = RGB(224, 224, 224)
        lblSubject.ForeColor = RGB(224, 224, 224)
        lblTitle.ForeColor = RGB(224, 224, 224)
    End If
Exit Sub
ErrHandler:
    Dim errorNumber As Integer
    errorNumber = Err.number
    Dim errorDescription As String
    errorDescription = Err.description
    logDebug "frmNewMsg.setLabelColor", errorDescription, errorNumber
    logTxt " frmNewMsg.setLabelColor", errorDescription, errorNumber
Resume Next
End Sub
