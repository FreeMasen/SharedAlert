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
'Gets the handle of the desktop to make the userform independant of outlook
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
'This finds the window in memory
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" _
        (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long

'This resets the parent window
Private Declare Function SetWindowLongA Lib "User32" _
        (ByVal hwnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long

'This defines the x, x and y position of a window
Private Declare Function SetWindowPos Lib "User32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long
      
Private Declare Function GetWindowLong Lib "User32" _
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
'this holds the item that created the userform
Public Item As mailItem

'This method acts as an override for the userform.show method
'it will avoid the form activate event
Public Sub ShowForm(newItem As mailItem)
    
    'set the global variable for the click events
    Set Item = newItem
    'update the captions of our lables
    Me.lblSubject.Caption = Item.subject
    Me.lblSender.Caption = Item.sender
    'set the window handle to this window's handle
    hwnd = FindWindow(vbNullString, Me.Caption)
    Dim Style As Long
    hwnd = FindWindow(vbNullString, Me.Caption)
    'Change the window's partent to the desktop
    SetWindowLongA hwnd, GWL_HWNDPARENT, GetDesktopWindow
    'variable for the windows type section of memory
    Dim style As Long
    'get that from our current window handle
    style = GetWindowLong(hwnd, -16)
    'reset the variable to include no border
    style = style And Not &HC800000
    'set the variable in the correct place in memory
    SetWindowLongA hwnd, -16, lStyle
    'this will remove any whitespace that the titlebar or border was occupying
    Me.Height = Me.InsideHeight
    Me.Width = Me.InsideWidth
    'set the z position to the top of the z order
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
    wait (4)
    Unload Me
End Sub

'this will display the inspector, intended to be placed in
'click or double click events
Private Sub ShowMsg()
    Dim inspector As inspector
    Set inspector = Item.GetInspector
    inspector.Display
End Sub

Public Sub wait(seconds As Integer)
    'use a do loop with doEvents in it to wait the passed number of seconds
    Dim waitTil As Date
    waitTil = Now + VBA.TimeSerial(0, 0, seconds)
    Do While Now() < waitTil
    DoEvents
    Loop
End Sub

Private Sub UserForm_Click()
    ShowMsg
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ShowMsg
End Sub
