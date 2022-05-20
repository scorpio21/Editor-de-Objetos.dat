Attribute VB_Name = "modSubClass"
Option Explicit

'===============================================================
'ListView LabelEdit
'© 2004 by Michiel Meulendijk
'
'This module handles the subclassing and belongs to the
'LabelEdit class module.
'===============================================================

Private Declare Function GetWindowLong _
                Lib "user32" _
                Alias "GetWindowLongA" (ByVal Hwnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong& _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal Hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long)

Private Declare Function CallWindowProc _
                Lib "user32" _
                Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                         ByVal Hwnd As Long, _
                                         ByVal msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long

Private Const GWL_WNDPROC = (-4)

Private Const WM_VSCROLL = &H115

Private Const WM_HSCROLL = &H114

Dim WndProcOld As Long

Dim colWnd     As Collection

Dim colClass   As Collection

'SubClass Code
Public Function WindProc(ByVal Hwnd As Long, _
                         ByVal wMsg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long

    If wMsg = WM_VSCROLL Or wMsg = WM_HSCROLL Then colClass.item("H" & Hwnd).SetText
    WindProc = CallWindowProc(WndProcOld&, Hwnd&, wMsg&, wParam&, lParam&)

End Function

Public Sub InitSubClass()
    Set colClass = New Collection

End Sub

Public Sub CloseSubClass()
    Set colClass = Nothing

End Sub

Public Sub SubClassWnd(Hwnd As Long, Class As Object)
    colClass.Add Class, "H" & Hwnd
    WndProcOld& = SetWindowLong(Hwnd, GWL_WNDPROC, AddressOf WindProc)

End Sub

Public Sub UnSubClassWnd(Hwnd As Long)
    SetWindowLong Hwnd, GWL_WNDPROC, WndProcOld&
    WndProcOld& = 0

End Sub

