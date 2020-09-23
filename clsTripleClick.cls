VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTripleClick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Two API's only!
Private Declare Function GetDoubleClickTime Lib "user32" () As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

'An API Mouse event which triggers button presses and releases
'Just added to prevent selecting after triple click
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cbuttons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTUP = &H4

Private bDoubleClicked   As Boolean 'Looks if there is a Double Click
Private lFirstTime       As Long    'Fires after Double Click
Private lSecondTime      As Long    'Fires after Triple Click
Private lTotalTime       As Long    'Calculate how long it took for the triple click
Private lDoubleClickTime As Long    'Gets time (in miliseconds) between two clicks, from windows
Private m_MouseUpAfter   As Boolean 'Prevents selecting, by sending Mouse Up Event

'Event Declarations, added a couple for example!
Event TripleClick()
Event DblClick()
Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'You can add all kids off Objects. I used a TextBox in this example.
'It is allso possible to use Hwnd's instead of this method, but anyway...
Private WithEvents vObj As TextBox
Attribute vObj.VB_VarHelpID = -1


'Set which Object to give a Triple Click Event.
Public Property Get vObject() As Variant
   Set vObject = vObj
End Property

Public Property Set vObject(ByVal NewObject As Variant)
   Set vObj = NewObject
End Property

Public Property Get MouseUpAfter() As Boolean
    MouseUpAfter = m_MouseUpAfter
End Property

Public Property Let MouseUpAfter(ByVal New_MouseUpAfter As Boolean)
   m_MouseUpAfter = New_MouseUpAfter
End Property

Private Sub Class_Initialize()
    m_MouseUpAfter = True   'True means send event Mouse Up, so stop selecting when mouse is down :-)
End Sub

Private Sub Class_Terminate()
    Set vObj = Nothing
End Sub



'Just a couple Events
Private Sub vObj_Click()
    RaiseEvent Click
End Sub

Private Sub vObj_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub vObj_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub vObj_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub vObj_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub vObj_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub



'*****************************************
'*  Ok, this part is doing the stuff!!!  *
'*****************************************

Private Sub vObj_DblClick()
    RaiseEvent DblClick                    'Just Raised Double Click Event.
    bDoubleClicked = True                  'Ok. there is a Double Click now.
    lDoubleClickTime = GetDoubleClickTime  'Get the standard double click time from windows.
    lFirstTime = GetTickCount              'Get the time when double click is fired first.
End Sub

Private Sub vObj_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)              'Just Raised Mouse Down Event.
    
    If Button = vbLeftButton Then                          'Do only when Mouse Left Button is pushed
        Select Case bDoubleClicked                         'Select a case boolean
            Case True                                      'When this is true...
                lSecondTime = GetTickCount                 '...then get second time again
                lTotalTime = (lSecondTime - lFirstTime)    'Calculate to see how long it took to click third mouse click
                
                If lTotalTime <= lDoubleClickTime Then     'Only RaiseEvent if total time is smaller then DoubleClickTime
                    RaiseEvent TripleClick                 'Yes! Triple Click have been raised
                End If
                
                If m_MouseUpAfter Then
                    mouse_event MOUSEEVENTF_LEFTUP, _
                                0, 0, 0, 0                 'Prevents selecting while mouse button is down after Triple Click
                End If
                
                bDoubleClicked = False                     'Do false for next Double Click
        End Select
    End If
End Sub

