Attribute VB_Name = "MSharedTimer"
' *********************************************************************
'  Copyright ©1995-2005 Karl E. Peterson, All Rights Reserved
'  http://vb.mvps.org/samples/TimerObj
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

' Win32 API Declarations
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const GWL_HWNDPARENT As Long = (-8)

#Const VBA = False

Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal oTimer As CTimer, ByVal dwTime As Long)
   ' Alert appropriate timer object instance.
   oTimer.RaiseTimer
End Sub

Public Function hWndMain() As Long
   ' This function returns the toplevel application hWnd.
   ' If running within VBA, this would be the main Word,
   ' Excel, PowerPoint (etc) window. If running in ClassicVB
   ' this would be the hidden top-level window.
   Call EnumThreadWindows(GetCurrentThreadId(), AddressOf EnumThreadWndProc, VarPtr(hWndMain))
End Function

Private Function EnumThreadWndProc(ByVal hWnd As Long, ByVal lpResult As Long) As Long
   ' This function depends on the conditional constant value, VBA
   #If VBA = False Then
      
      ' Test to see if this window is parented.
      ' If not, it's the one we're looking for!
      If GetWindowLong(hWnd, GWL_HWNDPARENT) Then
         ' Continue enumeration.
         EnumThreadWndProc = True
      Else
         ' Copy hWnd to result variable pointer,
         ' and stop enumeration.
         Call CopyMemory(ByVal lpResult, hWnd, 4)
         EnumThreadWndProc = False
      End If
   
   #ElseIf VBA = True Then
      Dim WindowText As String
      
      ' Continue enumeration, by default.
      EnumThreadWndProc = True
      
      ' Make sure this window isn't parented.
      ' Quick way to eliminate most windows.
      If GetWindowLong(hWnd, GWL_HWNDPARENT) = 0 Then
         
         ' Grab title/caption of this window.
         WindowText = Space$(512)
         If GetWindowText(hWnd, WindowText, Len(WindowText)) Then
         
            ' Look for application name in window title?
            If InStr(WindowText, Application.Name) Then
         
               ' Copy hWnd to result variable pointer,
               ' and stop enumeration.
               Call CopyMemory(ByVal lpResult, hWnd, 4&)
               EnumThreadWndProc = False
            End If
         End If
      End If
   #End If
End Function

