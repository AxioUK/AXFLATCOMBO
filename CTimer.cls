VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *********************************************************************
'  Copyright ©2005 Karl E. Peterson, All Rights Reserved
'  http://vb.mvps.org/samples/TimerObj
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private m_TmrID As Long
Private m_hWnd As Long
Private m_Enabled As Boolean
Private m_Interval As Long

Private Const defEnabled As Boolean = False
Private Const defInterval As Long = 0

Public Event Timer()

' ****************************************************
'  Initialization and Termination
' ****************************************************
Private Sub Class_Initialize()
   ' Set default values for members
   m_Enabled = defEnabled
   m_Interval = defInterval
   ' Get handle for main application window.
   m_hWnd = hWndMain()
End Sub

Private Sub Class_Terminate()
   ' Make sure timer is wiped out.
   Call KillTimer(m_hWnd, m_TmrID)
End Sub

' ****************************************************
'  Public Properties
' ****************************************************
Public Property Get Enabled() As Boolean
   ' Return cached setting.
   Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)
   ' Stop current timer, if already running.
   If m_Enabled Then
      Call KillTimer(m_hWnd, m_TmrID)
      m_TmrID = 0
   End If
   
   ' Cache passed value.
   m_Enabled = NewEnabled
   
   ' An interval of zero implicitly disables timer.
   If m_Interval > 0 Then
      ' Set new timer, if requested.
      If m_Enabled Then
         ' Pass pointer to Me so we can return event to this instance.
         m_TmrID = SetTimer(m_hWnd, ObjPtr(Me), m_Interval, AddressOf TimerProc)
      End If
   End If
End Property

Public Property Get Interval() As Long
   ' Return cached value.
   Interval = m_Interval
End Property

Public Property Let Interval(ByVal NewInterval As Long)
   ' Cache passed value, and reset timer.
   If NewInterval >= 0 Then
      m_Interval = NewInterval
      Me.Enabled = m_Enabled
   End If
End Property

' ****************************************************
'  Public Methods
' ****************************************************
Public Sub RaiseTimer()
   RaiseEvent Timer
End Sub

