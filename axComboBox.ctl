VERSION 5.00
Begin VB.UserControl axComboBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   KeyPreview      =   -1  'True
   ScaleHeight     =   145
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   253
   ToolboxBitmap   =   "axComboBox.ctx":0000
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   2085
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   2085
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   260
         Left            =   195
         TabIndex        =   4
         Top             =   150
         Width           =   1230
      End
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   1560
         ScaleHeight     =   360
         ScaleWidth      =   390
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   75
         Width           =   390
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2805
      Left            =   0
      ScaleHeight     =   187
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   139
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   675
      Width           =   2085
      Begin VB.ListBox Lst 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   225
         Sorted          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   135
         Width           =   1635
      End
   End
End
Attribute VB_Name = "axComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'-----DECLARACIONES-API---------------------------------------------------------
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'--------------------------------------------------------
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'--------------------------------------------------------
Private Declare Function CreateBitmap Lib "GDI32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
'Private Declare Function GetDIBits Lib "GDI32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, iPic As StdPicture) As Long
' recupera el estilo del Listbox
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
' cambia el estilo
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
' refresca y vuelve a redibujar el control
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'--------------------------------------------------------
Private Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String)
Private Declare Function LockWindowUpdate& Lib "user32" (ByVal hwndLock As Long)
'--------------------------------------------------------
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'Private Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
 
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
'--------------------------------------------------------

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X           As Long
    Y           As Long
End Type

Private Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type
    
Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type
    
Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'-------ENUMS---------------
Private Enum Blends
    RGBBlend = 0
    HSLBlend = 1
End Enum

Public Enum iSelectMode
    SingleClick = 0
    DoubleClick = 1
End Enum

Public Enum eStyleCombo
    [Dropdown Combo] = 0
    [Dropdown List] = 1
End Enum

Public Enum eEnterKeyBehav
    eNone = 0
    eKeyTab = 1
    eAddItem = 2
End Enum

'#############################################################################################################################
'Subclassing Code (all credits to Paul Caton!)
Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
  cbSize                             As Long
  dwFlags                            As TRACKMOUSEEVENT_FLAGS
  hwndTrack                          As Long
  dwHoverTime                        As Long
End Type

Private bTrack                       As Boolean
Private bTrackUser32                 As Boolean
Private mInCtrl                      As Boolean

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Enum eMsgWhen
    MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Type tSubData                                                                   'Subclass data type
    hWnd                               As Long                                            'Handle of the window being subclassed
    nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
    nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
    nMsgCntA                           As Long                                            'Msg after table entry count
    nMsgCntB                           As Long                                            'Msg before table entry count
    aMsgTblA()                         As Long                                            'Msg after table array
    aMsgTblB()                         As Long                                            'Msg Before table array
End Type

Private sc_aSubData()                As tSubData                                        'Subclass data array
Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WM_SETFOCUS            As Long = &H7
Private Const WM_KILLFOCUS           As Long = &H8
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSEHOVER As Long = &H2A1
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_VSCROLL As Long = &H115
Private Const WM_HSCROLL As Long = &H114
Private Const WM_LBUTTONDOWN         As Long = &H201
Private Const WM_RBUTTONDOWN         As Long = &H204
Private Const WM_GETMINMAXINFO       As Long = &H24
Private Const WM_SIZE                As Long = &H5
Private Const WM_WINDOWPOSCHANGED    As Long = &H47
Private Const WM_WINDOWPOSCHANGING   As Long = &H46
Private Const EVENT_TIMEOUT         As Long = 500
Private Const AUTOSCROLL_TIMEOUT    As Long = 50
'Private Const GWL_EXSTYLE = -20
'Private Const WS_EX_TOOLWINDOW = &H80
Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
'#############################################################################################################################
'User Control Declarations
'------CONSTANTES-----------------
Private Const DIB_RGB_COLORS = 0&
Private Const BI_RGB = 0&
Private Const m_def_BackColor = &HFFFFFF
Private Const m_def_BorderColor = &H808080
Private Const m_def_FocusColor = &HFFFFC0
Private Const m_def_ButtonColor = &H404040
Private Const m_def_Items = 8
Private Const m_def_Gradient = False
Private Const m_def_StyleCombo = 0
' constantes para SetWindowPos
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
' para GetWindowLong - SetWindowLong
Private Const GWL_STYLE = (-16)
Private Const WS_BORDER = &H800000
' para LoadRST
Private Const LB_ADDSTRING As Long = &H180
'################################################################
'------CONTROL--------------------------
Private WithEvents tmrMouseMove As CTimer
Attribute tmrMouseMove.VB_VarHelpID = -1
Private WithEvents tmrRelease As CTimer
Attribute tmrRelease.VB_VarHelpID = -1

'------VARIABLES-----------------
Dim OnLoad          As Boolean
Dim m_BackColor     As OLE_COLOR
Dim m_BorderColor   As OLE_COLOR
Dim m_FocusColor    As OLE_COLOR
Dim m_SelTextFocus  As Boolean
Dim m_ButtonColor   As OLE_COLOR
Dim Expanded        As Boolean
Dim m_Elements      As Integer
Dim m_SelectMode    As iSelectMode
Dim m_KeyBehavior   As eEnterKeyBehav
Dim m_Gradient      As Boolean
Dim m_StyleCombo    As eStyleCombo
Dim m_Enabled       As Boolean
Dim isDrawed        As Boolean
'You have to have MSScripting Runtime referenced : WshShell.SendKeys "{Tab}"
Dim WshShell        As Object
'------Private Variables---------
Private lBottomR    As Long
Private lBottomG    As Long
Private lBottomB    As Long
Private lTopR       As Long
Private lTopG       As Long
Private lTopB       As Long
Private Col1        As Long
Private Col2        As Long
Private LstH        As Integer
'Misc
Private mMDIChild   As Boolean
Private mInFocus    As Boolean
Private mMouseDown  As Boolean
Private mButtonRect As RECT
Private mScrollTick As Long
Private mIgnoreKeyPress As Boolean
Private mWindowsNT  As Boolean
Private mWindowsXP  As Boolean
Private m_hWnd      As Long
'################################################################
'------EVENTOS-----------------
Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event EnterKeyPress()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


'Subclass handler
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
'THIS MUST BE THE FIRST PUBLIC ROUTINE IN THIS FILE.
'That includes public properties also
'Parameters:
  'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
  'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
  'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
  'hWnd     - The window handle
  'uMsg     - The message number
  'wParam   - Message related data
  'lParam   - Message related data
        
    Select Case uMsg
    
    Case WM_KILLFOCUS
        'Another Control has got the focus
        DoKillFocus
        
    Case WM_MOUSEMOVE
        SetTimer False
        
        If m_Enabled Then
            If Not mInCtrl Then
                mInCtrl = True
                'RaiseEvent MouseEnter
                Call TrackMouseLeave(lng_hWnd)
                Call TrackMouseHover(lng_hWnd, 0)
            End If
        End If
        
        If IsMouseInScrollArea() Then
            DoAutoScroll
        End If
    
    Case WM_MOUSELEAVE
        mInCtrl = False
                    
        If picList.Visible Then
            Call GetAsyncKeyState(VK_LBUTTON)
            Call GetAsyncKeyState(VK_RBUTTON)
            SetTimer True
        End If
        
        'RaiseEvent MouseLeave
    
    Case WM_MOUSEHOVER
        If m_Enabled Then
            mInCtrl = False
        End If
              
    Case WM_WINDOWPOSCHANGING, WM_WINDOWPOSCHANGED, WM_GETMINMAXINFO, WM_SIZE, WM_LBUTTONDOWN, WM_RBUTTONDOWN
        'If Parent form is changing we want to close!
        DoKillFocus
        
    End Select
End Sub

Private Sub DoAutoScroll()
    Const MAX_COUNT As Long = 2147483647
    Static bActive As Boolean
    Dim uPoint  As POINTAPI
    Dim uRect As RECT
    Dim lCount As Long
    'This scrolls the list up/down when the mouse moves outside the DropDown
    'and the left button is pressed. It will terminate as soon as the mouse
    'moves back into the DropDown or the control loses focus
    'Prevent recursion
    If Not bActive Then
        bActive = True
        'Debug.Print "DoAutoScroll >"
        Call GetWindowRect(picList.hWnd, uRect)
        
        Do While mInFocus
            If (GetTickCount() - mScrollTick) > AUTOSCROLL_TIMEOUT Then
                mScrollTick = GetTickCount()
                
                Call GetCursorPos(uPoint)
                
            End If
            
            lCount = lCount + 1
            If (lCount Mod 10) = 0 Then
                DoEvents
            ElseIf lCount = MAX_COUNT Then
                lCount = 0
            End If
        Loop
        
        bActive = False
        'Debug.Print "DoAutoScroll <"
    End If
End Sub

Private Sub DoKillFocus()
    If picList.Visible Then
        SetDropDown
    End If

    If mInFocus Then
        mInFocus = False
    End If
End Sub

'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
  Dim hMod        As Long
  Dim bLibLoaded  As Boolean

  hMod = GetModuleHandleA(sModule)

  If hMod = 0 Then
    hMod = LoadLibraryA(sModule)
    If hMod Then
      bLibLoaded = True
    End If
  End If

  If hMod Then
    If GetProcAddress(hMod, sFunction) Then
      IsFunctionExported = True
    End If
  End If

  If bLibLoaded Then
    Call FreeLibrary(hMod)
  End If
End Function
'END Subclassing Code===================================================================================

Private Function IsMouseInScrollArea() As Boolean
    Dim uPoint  As POINTAPI
    Dim uRect As RECT
    
    Call GetWindowRect(picList.hWnd, uRect)
    Call GetCursorPos(uPoint)
    
    If (uPoint.Y < uRect.Top) Or (uPoint.Y > uRect.Bottom) Then
        IsMouseInScrollArea = True
    End If
End Function

Private Sub SetDropDown()
    'Dim R As RECT
    'Dim lHeight As Long
    'Dim dLeft As Single
    'Dim dTop As Single
    'Dim nCount As Integer
    
    With picList
        'If List is open then Close...
        If .Visible And Not IsMouseInScrollArea Then
            SetTimer False
            UserControl.Height = 350
            picList.Visible = False
            Expanded = False
            DrawPin (False)
            
        ElseIf ListCount() > 0 Then
            SetTimer True
        End If
    End With
End Sub

Private Sub SetTimer(bEnabled As Boolean)
    'If tmrRelease.Enabled <> bEnabled Then
        If bEnabled Then
            tmrRelease.Enabled = True
            'Debug.Print "Timer ON"
        Else
            tmrRelease.Enabled = False
            'Debug.Print "Timer OFF"
        End If
    'End If
End Sub

Private Sub tmrRelease_Timer()
    Dim uPoint  As POINTAPI
    Dim uRect As RECT
    Dim nLB As Integer
    Dim nRB As Integer
    
    '#############################################################################################################################
    'This is soley for detecting if we have clicked on a container which does not generate
    'WM_KILLFOCUS message for us to detect. i.e. the parent Form or a Frame
    
    'I don't like Timers in UserControls but wanted to make the Control behave as a normal Combo which
    'closes DropDown when the above situation occurs. I may still remove this "feature"!
    
    'NOTE: This Timer is only Enabled when we detect a WM_MOUSELEAVE so it does not fire unneccessarily
    'while the DropDown is displayed. It is Disabled as soon as the mouse re-enters the DropDown.
    '#############################################################################################################################
    
    Call GetCursorPos(uPoint)
    Call GetWindowRect(picList.hWnd, uRect)
        
    nLB = GetAsyncKeyState(VK_LBUTTON)
    nRB = GetAsyncKeyState(VK_RBUTTON)
    
    If (uPoint.X >= uRect.Left) And (uPoint.X <= uRect.Right) And (uPoint.Y >= uRect.Top) And (uPoint.Y <= uRect.Bottom) Then
        'The mouse pointer is within the Dropdown list
    ElseIf nLB Or nRB Then
        Select Case WindowFromPointXY(uPoint.X, uPoint.Y)
        Case UserControl.hWnd
            'The mouse pointer is within the Control
        Case Else
            If (GetTickCount() - mScrollTick) > EVENT_TIMEOUT Then
                If picList.Visible Then
                    SetDropDown
                Else
                    SetTimer False
                End If
            End If
        
        End Select
    End If
End Sub

'Track the mouse hovering the indicated window
Private Sub TrackMouseHover(ByVal lng_hWnd As Long, lHoverTime As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
  If bTrack Then
    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_HOVER
      .hwndTrack = lng_hWnd
      .dwHoverTime = lHoverTime
    End With

    If bTrackUser32 Then
      Call TrackMouseEvent(tme)
    Else
      Call TrackMouseEventComCtl(tme)
    End If
  End If
End Sub

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
  If bTrack Then
    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_LEAVE
      .hwndTrack = lng_hWnd
    End With

    If bTrackUser32 Then
      Call TrackMouseEvent(tme)
    Else
      Call TrackMouseEventComCtl(tme)
    End If
  End If
End Sub

'=======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.
'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
On Error GoTo Errs
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
Errs:
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
On Error GoTo Errs
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
Errs:
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
On Error GoTo Errs
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hWnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hWnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
'  If Not bAdd Then
'    Debug.Assert False                                                                  'hWnd not found, programmer error
'  End If
Errs:

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function

'========================================================================================
'Subclass routines below here - The programmer may call any of the following Subclass_??? routines
'======================================================================================================================================================
'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
On Error GoTo Errs

'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Const CODE_LEN              As Long = 204                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim i                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
'If it's the first time through here..
  If aBuf(1) = 0 Then
  
'The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'Convert the string from hex pairs to bytes and store in the static machine code buffer
    i = 1
    Do While j < CODE_LEN
      j = j + 1
      aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      i = i + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hWnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
  End With
Errs:
End Function

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hWnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
Errs:
End Sub

'Stop all subclassing
Private Sub Subclass_StopAll()
On Error GoTo Errs
  Dim i As Long
  
  i = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While i >= 0                                                                       'Iterate through each element
    With sc_aSubData(i)
      If .hWnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hWnd)                                                       'Subclass_Stop
      End If
    End With
    i = i - 1                                                                           'Next element
  Loop
Errs:
End Sub

Private Sub Txt_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Txt_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_ExitFocus()
    DoKillFocus
    'MsgBox "ExitFocus"
End Sub

Private Sub UserControl_InitProperties()
'Default Value Properties
m_BackColor = m_def_BackColor
m_BorderColor = m_def_BorderColor
m_FocusColor = m_def_FocusColor
m_ButtonColor = m_def_ButtonColor
m_SelectMode = SingleClick
m_Gradient = m_def_Gradient
m_Elements = m_def_Items
m_StyleCombo = m_def_StyleCombo
Expanded = False
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled And (Button = vbLeftButton) Then
        If (X > mButtonRect.Left) Then
            SetDropDown
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        Call FindPointer
        If X > UserControl.ScaleWidth Or X < 0 Or Y > UserControl.ScaleHeight Or Y < 0 Then
            ReleaseCapture
            mInCtrl = False
        ElseIf mInCtrl Then
            RaiseEvent MouseMove(Button, Shift, X, Y)
        Else
            mInCtrl = True
            Call TrackMouseLeave(UserControl.hWnd)
             
            RaiseEvent MouseMove(Button, Shift, X, Y)
        End If
    End If
End Sub

Private Sub UserControl_Paint()
UserControl_Resize
DrawPin (False)
DrawBorders
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'################################################################################
Dim Index As Integer
With PropBag
  Set Txt.Font = .ReadProperty("Font", Ambient.Font)
  Set Lst.Font = .ReadProperty("Font", Ambient.Font)
  Txt.ForeColor = .ReadProperty("ForeColor", vbButtonText)
  Txt.Text = .ReadProperty("Text", "")
  m_Enabled = .ReadProperty("Enabled", True)
  m_BackColor = .ReadProperty("BackColor", m_def_BackColor)
  m_BorderColor = .ReadProperty("BorderColor", m_def_BorderColor)
  m_ButtonColor = .ReadProperty("ButtonColor", m_def_ButtonColor)
  m_FocusColor = .ReadProperty("BackColorOnFocus", m_def_FocusColor)
  m_SelTextFocus = .ReadProperty("SelTextOnFocus", False)
  m_Elements = .ReadProperty("ItemsInList", m_def_Items)
  Lst.List(Index) = .ReadProperty("List" & Index, "")
  Lst.ListIndex = .ReadProperty("ListIndex", 0)
  m_KeyBehavior = .ReadProperty("EnterKeyBehavior", eNone)
  m_SelectMode = .ReadProperty("ItemSelectMode", SingleClick)
  m_Gradient = .ReadProperty("Gradient", m_def_Gradient)
  m_StyleCombo = .ReadProperty("StyleCombo", m_def_StyleCombo)
End With

UserControl_Resize

  Txt.SelLength = PropBag.ReadProperty("SelLength", 0)
  Txt.SelStart = PropBag.ReadProperty("SelStart", 0)
  Txt.SelText = PropBag.ReadProperty("SelText", "")
'################################################################################
    'Subclassing
    If Ambient.UserMode Then
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
        
        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If
        End If
        
        If TypeOf UserControl.Parent Is MDIForm Then
            mMDIChild = False
        Else
            mMDIChild = UserControl.Parent.MDIChild
        End If
        
        With UserControl.Parent
            Call Subclass_Start(.hWnd)
            Call Subclass_AddMsg(.hWnd, WM_WINDOWPOSCHANGING, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_WINDOWPOSCHANGED, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_GETMINMAXINFO, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_LBUTTONDOWN, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_SIZE, MSG_AFTER)
        End With

        With UserControl
            Call Subclass_Start(.hWnd)
            Call Subclass_AddMsg(.hWnd, WM_KILLFOCUS, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_SETFOCUS, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_MOUSEWHEEL, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_MOUSELEAVE, MSG_AFTER)
        End With

        With picList
            Call Subclass_Start(.hWnd)
            Call Subclass_AddMsg(.hWnd, WM_MOUSEWHEEL, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_MOUSEMOVE, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_MOUSELEAVE, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_MOUSEHOVER, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_HSCROLL, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_VSCROLL, MSG_AFTER)
        End With
        
    End If
    
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
With UserControl
    If Expanded = False Then
        .Height = 350
        Pic.Move 1, 1, .ScaleWidth - 2, .ScaleHeight - 2
        picButton.Move Pic.ScaleWidth - 350, 0, 350, 350
        Txt.Move 20, ((Pic.ScaleHeight / 2) - (Txt.Height / 2)), Pic.ScaleWidth - 360 ', Pic.ScaleHeight - 2
        Txt.Locked = False
    End If
    If m_StyleCombo = [Dropdown List] Then
      Txt.Locked = True
    Else
      Txt.Locked = False
    End If
End With

Call FindPointer
End Sub

Private Sub UserControl_Show()
'    'This modifies the PictureBox control so that it is not bound by
'    'its Container
'    'Dropdown can render over any Container the control is in
'    '(such as a Frame) and is not restricted by the Forms Boundaries

'    Dim lResult As Long
'    lResult = GetWindowLong(picList.hwnd, GWL_EXSTYLE)
'    Call SetWindowLong(picList.hwnd, GWL_EXSTYLE, lResult Or WS_EX_TOOLWINDOW)
'    Call SetWindowPos(picList.hwnd, picList.hwnd, 0, 0, 0, 0, 39)
'    Call SetWindowLong(picList.hwnd, -8, Parent.hwnd)
'    Call SetParent(picList.hwnd, 0)
End Sub

Private Sub UserControl_Terminate()
    Dim lHWnd As Long

    On Local Error GoTo UserControl_TerminateError
    
    
    Call Subclass_Stop(picList.hWnd)
    Call Subclass_Stop(UserControl.hWnd)
    
    If mMDIChild Then
        Call Subclass_StopAll
    Else
        lHWnd = UserControl.Parent.hWnd
        If lHWnd <> 0 Then
            Call Subclass_Stop(UserControl.Parent.hWnd)
        End If
    End If
    
UserControl_TerminateError:
    Exit Sub
End Sub

Public Sub Refresh()
  DrawPin (False)
  DrawBorders
  UserControl.Refresh
  UserControl_Resize
End Sub

Public Property Get hdc() As Long
   hdc = UserControl.hdc
End Property

Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

' Función que carga el campo en el combobox o listbox
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function LoadfromRST(Rst As ADODB.Recordset, Columna As String) As Boolean
'  Dim ret                 As Long
'  Dim Mensaje_SendMessage As Long
'
' ' On Error GoTo Error_Function:
'  ' verifica que el recordset contenga un conjunto de registros
'  If Rst.BOF And Rst.EOF Then
'    MsgBox " No hay registros para agregar", vbInformation
'    Call LockWindowUpdate(0&)
'    ' sale
'    Exit Function
'  End If
'
'  Mensaje_SendMessage = LB_ADDSTRING ' mensaje para SendMessage
'
'  ' deshabilita el repintado del control para que cargue los datos mas rapidamente
'  Call LockWindowUpdate(Lst.hWnd)
'  DoEvents
'  ' Posiciona el recordset en el primer registro
'  Rst.MoveFirst
'  ' elimina todo el contenido del combo o listbox( opcional )
'  Lst.Clear
'  ' recorre las filas del recordset
'  Do Until Rst.EOF
'    ' chequea que el valor no sea un nulo
'    If Not IsNull(Rst(Columna).Value) Then
'      'Agrega el dato en el control con el mensaje CB_ADDSTRING o LB_ADDSTRING dependiendo del tipo de control
'      ret = SendMessageByString(Lst.hWnd, Mensaje_SendMessage, 0, Rst(Columna).Value)
'    End If
'    ' siguiente registro
'    Rst.MoveNext
'  Loop
'
'  ' selecciona el primer elemento del listado
'  If Lst.ListCount > 0 Then
'    Lst.ListIndex = 0
'  End If
'
'  ' vuelve a habilitar el repintado
'  Call LockWindowUpdate(0&)
'  ' retorno
'  LoadfromRST = True
'
'  Exit Function
'  ' rutina de error
'Error_Function:
'  MsgBox Err.Description, vbCritical
'  ' En caso de error vuelve a activar el repintado
'  Call LockWindowUpdate(0&)
'  Lst.Refresh
'End Function

Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Integer)
If IsMissing(Index) Then
    Lst.AddItem Item
Else
    Lst.AddItem Item, Index
End If
End Sub

Public Sub RemoveItem(ByVal Index As Integer)
On Error Resume Next
    Lst.RemoveItem Index
End Sub

Public Sub Clear()
    Lst.Clear

End Sub


'------Helper Functions----------------
Private Function CreateGradient(Width As Long, Height As Long, LeftToRight As Boolean, LeftTopColor As Long, RightBottomColor As Long, BlendType As Blends) As StdPicture
    Dim hBmp As Long, Bits() As Byte
    Dim RS As Byte, GS As Byte, BS As Byte
    Dim RE As Byte, GE As Byte, BE As Byte
    Dim HS As Single, SS As Single, lS As Single
    Dim HE As Single, SE As Single, LE As Single
    Dim rc As Byte, GC As Byte, BC As Byte
    Dim X As Long, Y As Long
    ReDim Bits(0 To 3, 0 To Width - 1, 0 To Height - 1)
    
    RgbCol LeftTopColor, RS, GS, BS
    RgbCol RightBottomColor, RE, GE, BE
    
    If BlendType = RGBBlend Then
        If LeftToRight Then
            For X = 0 To Width - 1
                rc = (1& * RS - RE) * ((Width - 1 - X) / (Width - 1)) + RE
                GC = (1& * GS - GE) * ((Width - 1 - X) / (Width - 1)) + GE
                BC = (1& * BS - BE) * ((Width - 1 - X) / (Width - 1)) + BE
                For Y = 0 To Height - 1
                    Bits(2, X, Y) = rc
                    Bits(1, X, Y) = GC
                    Bits(0, X, Y) = BC
                Next
            Next
        Else
            For Y = 0 To Height - 1
                rc = (1& * RS - RE) * ((Height - 1 - Y) / (Height - 1)) + RE
                GC = (1& * GS - GE) * ((Height - 1 - Y) / (Height - 1)) + GE
                BC = (1& * BS - BE) * ((Height - 1 - Y) / (Height - 1)) + BE
                For X = 0 To Width - 1
                    Bits(2, X, Y) = rc
                    Bits(1, X, Y) = GC
                    Bits(0, X, Y) = BC
                Next
            Next
        End If
    ElseIf BlendType = HSLBlend Then
        RGBToHSL RS, GS, BS, HS, SS, lS
        RGBToHSL RE, GE, BE, HE, SE, LE
        If LeftToRight Then
            For X = 0 To Width - 1
                HSLToRGB (1& * HS - HE) * ((Width - 1 - X) / (Width - 1)) + HE, _
                        (1& * SS - SE) * ((Width - 1 - X) / (Width - 1)) + SE, _
                        (1& * lS - LE) * ((Width - 1 - X) / (Width - 1)) + LE, _
                        rc, GC, BC
                For Y = 0 To Height - 1
                    Bits(2, X, Y) = rc
                    Bits(1, X, Y) = GC
                    Bits(0, X, Y) = BC
                Next
            Next
        Else
            For Y = 0 To Height - 1
                HSLToRGB (1& * HS - HE) * ((Height - 1 - Y) / (Height - 1)) + HE, _
                        (1& * SS - SE) * ((Height - 1 - Y) / (Height - 1)) + SE, _
                        (1& * lS - LE) * ((Height - 1 - Y) / (Height - 1)) + LE, _
                        rc, GC, BC
                For X = 0 To Width - 1
                    Bits(2, X, Y) = rc
                    Bits(1, X, Y) = GC
                    Bits(0, X, Y) = BC
                Next
            Next
        End If
    End If

    Dim BI As BITMAPINFO
    With BI.bmiHeader
        .biSize = Len(BI.bmiHeader)
        .biWidth = Width
        .biHeight = -Height
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        .biSizeImage = ((((.biWidth * .biBitCount) + 31) \ 32) * 4) * Abs(.biHeight)
    End With
    hBmp = CreateBitmap(Width, Height, 1&, 32&, ByVal 0)
    SetDIBits 0&, hBmp, 0, Abs(BI.bmiHeader.biHeight), Bits(0, 0, 0), BI, DIB_RGB_COLORS

    Dim IGuid As Guid, PicDst As PictDesc
    
    With IGuid
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    With PicDst
        .cbSizeofStruct = Len(PicDst)
        .hImage = hBmp
        .picType = vbPicTypeBitmap
    End With
    OleCreatePictureIndirect PicDst, IGuid, True, CreateGradient
End Function

Private Sub DrawBorders()
UserControl.Cls
Dim Rgn As Long

With UserControl
    Rgn = CreateRoundRectRgn(0, 0, .Width, .Height, 0, 0)
    SetWindowRgn .hWnd, Rgn, True
    DeleteObject Rgn
    .DrawWidth = 1
    .ForeColor = m_BorderColor
    RoundRect .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0
End With

With picList
    Rgn = CreateRoundRectRgn(0, 0, .Width, .Height, 0, 0)
    SetWindowRgn .hWnd, Rgn, True
    DeleteObject Rgn
    .DrawWidth = 1
    .ForeColor = m_BorderColor
    RoundRect .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0
End With

Dim lng_Estilo As Long

    With Lst
        '.Appearance = 0 ' flat
        lng_Estilo = GetWindowLong(.hWnd, GWL_STYLE)
        lng_Estilo = lng_Estilo And Not WS_BORDER ' sin borde
        ' aplica
        SetWindowLong .hWnd, GWL_STYLE, lng_Estilo
        ' refresh
        SetWindowPos .hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or _
                                           SWP_NOACTIVATE Or _
                                           SWP_NOMOVE Or _
                                           SWP_NOOWNERZORDER Or _
                                           SWP_NOSIZE Or _
                                           SWP_NOZORDER
    End With
    
End Sub

Private Sub DrawPin(isMouseOver As Boolean)
  Dim i As Integer
  Dim iHorizontal1 As Integer
  Dim iHorizontal2 As Integer
  Dim iVertical As Integer
  Dim cColor As OLE_COLOR
  
  picButton.Cls
  
  If Expanded = False Then
        Col1 = m_ButtonColor
        Col2 = &HFFFFFF
  Else
        Col1 = &HFFFFFF
        Col2 = m_ButtonColor
  End If
  
  lBottomR = (Col1 And &HFF&)
  lBottomG = (Col1 And &HFF00&) / &H100
  lBottomB = (Col1 And &HFF0000) / &H10000
  
  lTopR = (Col2 And &HFF&)
  lTopG = (Col2 And &HFF00&) / &H100
  lTopB = (Col2 And &HFF0000) / &H10000

If m_Gradient Then
  Set picButton.Picture = CreateGradient(picButton.Width / Screen.TwipsPerPixelX, picButton.Height / Screen.TwipsPerPixelY, False, RGB(lTopR, lTopG, lTopB), RGB(lBottomR, lBottomG, lBottomB), RGBBlend)
Else
  picButton.BackColor = m_ButtonColor
End If

If isMouseOver = False Then
    cColor = m_BackColor
Else
    cColor = m_FocusColor
End If
  

If Expanded = False Then
      iHorizontal1 = 175 '210
      iHorizontal2 = 160 '195
      iVertical = 105
      For i = 1 To 2
          ' 1st Line of 1st Arrow
          picButton.Line (picButton.Width - (iHorizontal1 + 45), iVertical)-(picButton.Width - (iHorizontal1 + 15), iVertical), cColor
          picButton.Line (picButton.Width - (iHorizontal2 - 15), iVertical)-(picButton.Width - (iHorizontal2 - 45), iVertical), cColor
          iVertical = iVertical + 15
      
          ' 2nd Line of 1st Arrow
          picButton.Line (picButton.Width - (iHorizontal1 + 30), iVertical)-(picButton.Width - iHorizontal1, iVertical), cColor
          picButton.Line (picButton.Width - iHorizontal2, iVertical)-(picButton.Width - (iHorizontal2 - 30), iVertical), cColor
          iVertical = iVertical + 15
          
          ' 1st Line of 2nd Arrow
          picButton.Line (picButton.Width - (iHorizontal1 + 15), iVertical)-(picButton.Width - (iHorizontal1 - 30), iVertical), cColor
          iVertical = iVertical + 15
          
          ' 2nd Line of 2nd Arrow
          picButton.Line (picButton.Width - iHorizontal1, iVertical)-(picButton.Width - (iHorizontal1 - 15), iVertical), cColor
          iVertical = iVertical + 15
      Next

Else
      iHorizontal1 = 175
      iHorizontal2 = 160
      iVertical = 105
      For i = 1 To 2
          ' 1st Line of 1st Arrow
          picButton.Line (picButton.Width - iHorizontal1, iVertical)-(picButton.Width - (iHorizontal1 - 15), iVertical), cColor
          iVertical = iVertical + 15
          
          ' 2nd Line of 1st Arrow
          picButton.Line (picButton.Width - (iHorizontal1 + 15), iVertical)-(picButton.Width - (iHorizontal1 - 30), iVertical), cColor
          iVertical = iVertical + 15
          
          ' 1st Line of 2nd Arrow
          picButton.Line (picButton.Width - (iHorizontal1 + 30), iVertical)-(picButton.Width - iHorizontal1, iVertical), cColor
          picButton.Line (picButton.Width - iHorizontal2, iVertical)-(picButton.Width - (iHorizontal2 - 30), iVertical), cColor
          iVertical = iVertical + 15
          
          ' 2nd Line of 2nd Arrow
          picButton.Line (picButton.Width - (iHorizontal1 + 45), iVertical)-(picButton.Width - (iHorizontal1 + 15), iVertical), cColor
          picButton.Line (picButton.Width - (iHorizontal2 - 15), iVertical)-(picButton.Width - (iHorizontal2 - 45), iVertical), cColor
          iVertical = iVertical + 15
      Next
End If

Debug.Print "DrawPin:Expanded " & Expanded
End Sub

Private Sub FindPointer()
    Dim pt As POINTAPI

    GetCursorPos pt
    
'If Expanded Then Exit Sub

    If WindowFromPointXY(pt.X, pt.Y) = picButton.hWnd Then
       Call DrawPin(True)
       Call DrawBorders
    Else
        Call DrawPin(False)
        tmrMouseMove.Enabled = False
    End If
    
    isDrawed = True
End Sub

Private Sub HSLToRGB(ByVal H As Single, ByVal s As Single, ByVal l As Single, R As Byte, g As Byte, b As Byte)
    Dim rR As Single, rG As Single, rB As Single
    Dim Min As Single, Max As Single
    
    If s = 0 Then
        rR = l: rG = l: rB = l
    Else
        If l <= 0.5 Then
            Min = l * (1 - s)
        Else
            Min = l - s * (1 - l)
        End If
        Max = 2 * l - Min
       
        If (H < 1) Then
            rR = Max
            If (H < 0) Then
                rG = Min
                rB = rG - H * (Max - Min)
            Else
                rB = Min
                rG = H * (Max - Min) + rB
            End If
        ElseIf (H < 3) Then
            rG = Max
            If (H < 2) Then
                rB = Min
                rR = rB - (H - 2) * (Max - Min)
            Else
                rR = Min
                rB = (H - 2) * (Max - Min) + rR
            End If
        Else
            rB = Max
            If (H < 4) Then
                rR = Min
                rG = rR - (H - 4) * (Max - Min)
            Else
                rG = Min
                rR = (H - 4) * (Max - Min) + rG
            End If
        End If
    End If
    R = rR * 255: g = rG * 255: b = rB * 255
End Sub

Private Sub RgbCol(col As Long, ByRef R As Byte, ByRef g As Byte, ByRef b As Byte)
    R = col And &HFF&
    g = (col And &HFF00&) \ &H100&
    b = (col And &HFF0000) \ &H10000
End Sub

Private Sub RGBToHSL(ByVal R As Byte, ByVal g As Byte, ByVal b As Byte, H As Single, s As Single, l As Single)
    Dim Max As Single
    Dim Min As Single
    Dim Delta As Single
    Dim rR As Single, rG As Single, rB As Single

    rR = R / 255: rG = g / 255: rB = b / 255

    Max = Maximum(rR, rG, rB)
    Min = Minimum(rR, rG, rB)
    l = (Max + Min) / 2
    If Max = Min Then
        s = 0
        H = 0
    Else
        If l <= 0.5 Then
            s = (Max - Min) / (Max + Min)
        Else
            s = (Max - Min) / (2 - Max - Min)
        End If
        
        Delta = Max - Min
        If rR = Max Then
            H = (rG - rB) / Delta
        ElseIf rG = Max Then
            H = 2 + (rB - rR) / Delta
        ElseIf rB = Max Then
            H = 4 + (rR - rG) / Delta
        End If
    End If
End Sub

Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
     'http://www.vbAccelerator.com
    If (rR > rG) Then
        If (rR > rB) Then
            Maximum = rR
        Else
            Maximum = rB
        End If
    Else
        If (rB > rG) Then
            Maximum = rB
        Else
            Maximum = rG
        End If
    End If
End Function

Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
     'http://www.vbAccelerator.com
    If (rR < rG) Then
        If (rR < rB) Then
            Minimum = rR
        Else
            Minimum = rB
        End If
    Else
        If (rB < rG) Then
            Minimum = rB
        Else
            Minimum = rG
        End If
    End If
End Function

Private Sub UserControlsCreate()
Set tmrMouseMove = New CTimer
Set tmrRelease = New CTimer

tmrMouseMove.Interval = 100
tmrMouseMove.Enabled = False
tmrRelease.Enabled = False
tmrRelease.Interval = 10

End Sub

Private Sub Lst_DblClick()
If OnLoad = False Then
  If m_SelectMode = DoubleClick Then
      Txt.Text = Lst.Text
      UserControl.Height = 350
      picList.Visible = False

  Else
    Exit Sub
  End If
End If

RaiseEvent DblClick

End Sub

Private Sub Lst_Click()
If OnLoad = False Then
  If m_SelectMode = SingleClick Then
      Txt.Text = Lst.Text
      UserControl.Height = 350
      picList.Visible = False
      Expanded = False
      DrawPin (False)
      
  Else
    Exit Sub
  End If
End If

RaiseEvent Click
End Sub


Private Sub picButton_Click()
LstH = 225 + (195 * (m_Elements - 1))

OnLoad = False
Expanded = Not Expanded

With UserControl
    If Expanded = True Then
        picList.Visible = True
        picList.Move 0, Pic.Height + 1, .ScaleWidth - 2, (LstH + 10) / 15
        Lst.Move 1, 1, picList.ScaleWidth, (LstH + 10) / 15
        .Height = LstH + 350
    Else
        .Height = 350
        picList.Visible = False
    End If
End With

Call DrawBorders

Debug.Print "----------------------------------------------------------"
Debug.Print "picButton:Expanded " & Expanded
End Sub

Private Sub picButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tmrMouseMove.Enabled = True
End Sub

Private Sub tmrMouseMove_Timer()
If isDrawed = True Then Exit Sub

Call FindPointer

End Sub

Private Sub Txt_Change()
RaiseEvent Change
End Sub

Private Sub Txt_GotFocus()
With Txt
  If m_SelTextFocus = True Then
      .SelStart = 0
      .SelLength = Len(.Text)
  End If
  .BackColor = m_FocusColor
End With

Pic.BackColor = m_FocusColor
End Sub

Private Sub Txt_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  Select Case m_KeyBehavior
      Case Is = eNone
        'Do Nothing
      Case Is = eKeyTab
        WshShell.SendKeys "{Tab}"
        
      Case Is = eAddItem
        Dim iL As Integer
        With Lst
          iL = .ListCount
          .AddItem Txt.Text, iL
          'Abro List
          'picButton_Click
        End With
  End Select

  RaiseEvent EnterKeyPress
End If
'---------------------------------
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Txt_LostFocus()
Txt.BackColor = m_BackColor
Pic.BackColor = m_BackColor

End Sub

Private Sub UserControl_Initialize()
Set WshShell = CreateObject("WScript.Shell")
Call UserControlsCreate
Call FindPointer

isDrawed = False
Expanded = False
OnLoad = True
Txt.BackColor = vbWhite
Lst.Clear
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer
With PropBag
   .WriteProperty "Font", Txt.Font, Ambient.Font
   .WriteProperty "ForeColor", Txt.ForeColor, vbButtonText
   .WriteProperty "Enabled", m_Enabled, True
   .WriteProperty "Text", Txt.Text, ""
   .WriteProperty "BackColor", m_BackColor, m_def_BackColor
   .WriteProperty "BorderColor", m_BorderColor, m_def_BorderColor
   .WriteProperty "ButtonColor", m_ButtonColor, m_def_ButtonColor
   .WriteProperty "BackColorOnFocus", m_FocusColor, m_def_FocusColor
   .WriteProperty "SelTextOnFocus", m_SelTextFocus, False
   .WriteProperty "ItemsInList", m_Elements, m_def_Items
   .WriteProperty "ItemSelectMode", m_SelectMode, SingleClick
   .WriteProperty "List" & Index, Lst.List(Index), ""
   .WriteProperty "ListIndex", Lst.ListIndex, 0
   .WriteProperty "EnterKeyBehavior", m_KeyBehavior, eNone
   .WriteProperty "ItemSelectMode", m_SelectMode, SingleClick
   .WriteProperty "Gradient", m_Gradient, m_def_Gradient
   .WriteProperty "StyleCombo", m_StyleCombo, m_def_StyleCombo
End With

UserControl_Resize

  Call PropBag.WriteProperty("SelLength", Txt.SelLength, 0)
  Call PropBag.WriteProperty("SelStart", Txt.SelStart, 0)
  Call PropBag.WriteProperty("SelText", Txt.SelText, "")
  
End Sub

''------------------ PROPERTIES --------------------------
Public Property Get BackColor() As OLE_COLOR
BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
m_BackColor = NewBackColor
Txt.BackColor = NewBackColor
UserControl.BackColor = NewBackColor
Pic.BackColor = NewBackColor
PropertyChanged "BackColor"
Refresh
End Property

Public Property Get BackColorOnFocus() As OLE_COLOR
BackColorOnFocus = m_FocusColor
End Property

Public Property Let BackColorOnFocus(ByVal NewColor As OLE_COLOR)
m_FocusColor = NewColor
PropertyChanged "BackColorOnFocus"
Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)
m_BorderColor = NewBorderColor
PropertyChanged "BorderColor"
Refresh
End Property

Public Property Get ButtonColor() As OLE_COLOR
ButtonColor = m_ButtonColor
End Property

Public Property Let ButtonColor(ByVal NewButtonColor As OLE_COLOR)
m_ButtonColor = NewButtonColor
PropertyChanged "ButtonColor"
Refresh
End Property

Public Property Get Enabled() As Boolean
Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
m_Enabled = New_Enabled
PropertyChanged "Enabled"
End Property

Public Property Get Gradient() As Boolean
Gradient = m_Gradient
End Property

Public Property Let Gradient(ByVal New_Gradient As Boolean)
m_Gradient = New_Gradient
PropertyChanged "Gradient"
Refresh
End Property

Public Property Get EnterKeyBehavior() As eEnterKeyBehav
EnterKeyBehavior = m_KeyBehavior
End Property

Public Property Let EnterKeyBehavior(ByVal NewBehavior As eEnterKeyBehav)
m_KeyBehavior = NewBehavior
PropertyChanged "EnterKeyBehavior"
End Property

Public Property Get Font() As Font
Set Font = Txt.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
Set Txt.Font = New_Font
Set Lst.Font = New_Font
PropertyChanged "Font"
Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = Txt.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
Txt.ForeColor = New_ForeColor
Lst.ForeColor = New_ForeColor
PropertyChanged "ForeColor"
Refresh
End Property

Public Property Get ItemSelectMode() As iSelectMode
    ItemSelectMode = m_SelectMode
End Property

Public Property Let ItemSelectMode(ByVal new_Select As iSelectMode)
    m_SelectMode = new_Select
    PropertyChanged "ItemSelectMode"
End Property

Public Property Get ItemsInList() As Integer
ItemsInList = m_Elements
End Property

Public Property Let ItemsInList(ByVal vNewValue As Integer)
m_Elements = vNewValue
PropertyChanged "ItemsInList"
Refresh
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Lst,Lst,-1,ListCount
Public Property Get ListCount() As Integer
    ListCount = Lst.ListCount
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Lst,Lst,-1,List
Public Property Get List(ByVal Index As Integer) As String
    List = Lst.List(Index)
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Lst,Lst,-1,ListIndex
Public Property Get ListIndex() As Integer
    ListIndex = Lst.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    Lst.ListIndex() = New_ListIndex
    PropertyChanged "ListIndex"
    Txt.Text = Lst.Text
End Property

Public Property Let List(ByVal Index As Integer, ByVal New_List As String)
    Lst.List(Index) = New_List
    PropertyChanged "List"
End Property

Public Property Get SelTextOnFocus() As Boolean
SelTextOnFocus = m_SelTextFocus
End Property

Public Property Let SelTextOnFocus(ByVal bSel As Boolean)
m_SelTextFocus = bSel
PropertyChanged "SelTextOnFocus"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Lst,Lst,-1,Sorted
Public Property Get Sorted() As Boolean
    Sorted = Lst.Sorted
End Property

Public Property Get Text() As String
Text = Txt.Text
End Property

Public Property Let Text(ByVal New_Text As String)
Txt.Text = New_Text
PropertyChanged "Text"
Refresh
End Property


'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Txt,Txt,-1,SelLength
Public Property Get SelLength() As Long
  SelLength = Txt.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
  Txt.SelLength() = New_SelLength
  PropertyChanged "SelLength"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Txt,Txt,-1,SelStart
Public Property Get SelStart() As Long
  SelStart = Txt.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
  Txt.SelStart() = New_SelStart
  PropertyChanged "SelStart"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Txt,Txt,-1,SelText
Public Property Get SelText() As String
  SelText = Txt.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
  Txt.SelText() = New_SelText
  PropertyChanged "SelText"
End Property
'm_StyleCombo
Public Property Get StyleCombo() As eStyleCombo
  StyleCombo = m_StyleCombo
End Property

Public Property Let StyleCombo(ByVal NewStyleCombo As eStyleCombo)
  m_StyleCombo = NewStyleCombo
  PropertyChanged "StyleCombo"
  Refresh
End Property
