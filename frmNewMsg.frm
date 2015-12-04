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
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long
      
'This defines a parent window as the desktop
Private Const GWL_HWNDPARENT As Long = (-8&)

'The first 3 const here are just saying that the
'window will not move ore be resized by SetWindowPos
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const FLAGS As Long = SWP_NOMOVE Or SWP_NOSIZE

'This will pull the window to the top of the "Top Most Windows" z order
'(windows keeps 2 z orders NOTTOPMOST and TOPMOST
Private Const HWND_TOPMOST = -1

'This holds the memory handle for the window
Private hwnd As Long
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------

'This holds our mailitem for the double click event
Public item As Outlook.mailItem

Private Sub UserForm_Initialize()
    'set the window handle to this window's handle
    hwnd = FindWindow(vbNullString, Me.Caption)
    
    'setup our window with a desktop parent
    SetWindowLongA hwnd, GWL_HWNDPARENT, 0&
End Sub

Private Sub UserForm_Activate()
    'set the z position to the top of the z order
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
    'this just stops the window from holing anything up
    DoEvents
    'open, wait 4 second then close
    wait (4)
    Unload Me
End Sub

Private Sub wait(seconds As Integer)
    'use a do loop with doEvents in it to wait the passed number of seconds
    Dim waitTil As Date
    waitTil = Now + VBA.TimeSerial(0, 0, seconds)
    Do While Now() < waitTil
    DoEvents
    Loop
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'open and inspector with the current item
Dim insp As Inspector
Set insp = item.GetInspector
insp.Display
End Sub

