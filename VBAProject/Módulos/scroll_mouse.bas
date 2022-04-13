Attribute VB_Name = "scroll_mouse"


Option Explicit

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type MOUSEHOOKSTRUCT
        pt As POINTAPI
        hWnd As Long
        wHitTestCode As Long
        dwExtraInfo As Long
End Type

Private Declare Function FindWindow Lib "User32" _
                                        Alias "FindWindowA" ( _
                                                        ByVal lpClassName As String, _
                                                        ByVal lpWindowName As String) As Long

Private Declare Function GetWindowLong Lib "user32.dll" _
                                        Alias "GetWindowLongA" ( _
                                                        ByVal hWnd As Long, _
                                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowsHookEx Lib "User32" _
                                        Alias "SetWindowsHookExA" ( _
                                                        ByVal idHook As Long, _
                                                        ByVal lpfn As Long, _
                                                        ByVal hmod As Long, _
                                                        ByVal dwThreadId As Long) As Long

Private Declare Function CallNextHookEx Lib "User32" ( _
                                                        ByVal hHook As Long, _
                                                        ByVal nCode As Long, _
                                                        ByVal wParam As Long, _
                                                        lParam As Any) As Long

Private Declare Function UnhookWindowsHookEx Lib "User32" ( _
                                                        ByVal hHook As Long) As Long

Private Declare Function WindowFromPoint Lib "User32" ( _
                                                        ByVal xPoint As Long, _
                                                        ByVal yPoint As Long) As Long

Private Declare Function GetCursorPos Lib "user32.dll" ( _
                                                        ByRef lpPoint As POINTAPI) As Long

Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const HC_ACTION As Long = 0
Private Const GWL_HINSTANCE As Long = (-6)


Private mLngMouseHook As Long
Private mListBoxHwnd As Long
Private mbHook As Boolean
Private mCtl As MSForms.Control
Dim n As Long

Sub HookListScroll(frm As Object, Ctl As MSForms.Control)
Dim lngAppInst As Long
Dim hwndUnderCursor As Long
Dim tPT As POINTAPI
     GetCursorPos tPT
     hwndUnderCursor = WindowFromPoint(tPT.X, tPT.Y)
     If Not frm.ActiveControl Is Ctl Then
             Ctl.SetFocus
     End If
     If mListBoxHwnd <> hwndUnderCursor Then
             UnhookListScroll
             Set mCtl = Ctl
             mListBoxHwnd = hwndUnderCursor
             lngAppInst = GetWindowLong(mListBoxHwnd, GWL_HINSTANCE)
             
             If Not mbHook Then
                     mLngMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, lngAppInst, 0)
                     mbHook = mLngMouseHook <> 0
             End If
     End If
End Sub

Sub UnhookListScroll()
     If mbHook Then
                Set mCtl = Nothing
             UnhookWindowsHookEx mLngMouseHook
             mLngMouseHook = 0
             mListBoxHwnd = 0
             mbHook = False
        End If
End Sub

Private Function MouseProc( _
             ByVal nCode As Long, ByVal wParam As Long, _
             ByRef lParam As MOUSEHOOKSTRUCT) As Long
Dim idx As Long
        On Error GoTo errH
     If (nCode = HC_ACTION) Then
             If WindowFromPoint(lParam.pt.X, lParam.pt.Y) = mListBoxHwnd Then
                     If wParam = WM_MOUSEWHEEL Then
                                MouseProc = True
                                
                                If lParam.hWnd > 0 Then idx = -1 Else idx = 1
                             
                             idx = idx + mCtl.TopIndex
                             If idx >= 0 Then mCtl.TopIndex = idx
                                Exit Function
                     End If
             Else
                     UnhookListScroll
             End If
     End If
     MouseProc = CallNextHookEx(mLngMouseHook, nCode, wParam, ByVal lParam)
     Exit Function
errH:
     UnhookListScroll
End Function




