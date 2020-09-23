Attribute VB_Name = "Module1"
Option Explicit
'***********************
'  Written by GioRock  *
'***********************
'***********************
'  Created by GioRock  *
'***********************

' I want to try explain how load menu resource in VB.
' The sample resource is Addobe Photoshop menu
' I get this with my program named <<Clone Menu>>
' I can't publish this program now, because it's very complex
' strucured algorythm and hard to descript
' Capture detailed menu window and transform them in
' precise resource file '*.rc' and accelerator table too
' Maybe later.....

' Memory manage function
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

' Menu function
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Const MF_POPUP = &H10&
Private Const MF_BYPOSITION = &H400
Private Const MF_BYCOMMAND = &H0
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As Any) As Long
Private Const TPM_RIGHTBUTTON = &H2&
Private Const TPM_LEFTALIGN = &H0&

' Cursor function
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

' Message function
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Public Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    Pt As POINTAPI
End Type
Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long

' SubClassing Function
Private Declare Function CallWindowProc Lib "User32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "User32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC As Long = (-4)

' Window message
Private Const WM_DESTROY = &H2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_COMMAND = &H111
Private Const WM_SETFOCUS = &H7
Private Const WM_MENUSELECT As Long = &H11F
Private Const WM_INITMENU = &H116
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_SYSCOMMAND = &H112
Private Const WM_CHAR = &H102
Private Const WM_KEYDOWN = &H100
Private Const WM_RBUTTONDOWN = &H204

' To get Caption Menu
Private sCaption As String
' Handle procedure
Private OldWndProc As Long
Private SubclassedhWnd As Long


Public Function MAKELONG(ByVal wLow As Integer, ByVal wHi As Integer) As Long
Dim lTemp As Long
    ' Return a Long value from two Integer value
    ' Length Integer = 2, lTemp = wLow store first two byte
    ' MAKELONG base pointer + 2 = wHi
    ' store second two byte
    lTemp = wLow
    MoveMemory ByVal VarPtr(lTemp) + 2, wHi, 2 ' Length Integer = 2
    MAKELONG = lTemp
End Function
Public Function HIWORD(ByVal l As Long) As Integer
    MoveMemory HIWORD, ByVal VarPtr(l) + 2, 2
End Function

Public Function LOWORD(ByVal dwValue As Long) As Integer
    MoveMemory LOWORD, dwValue, 2
End Function



Public Function StripTerminator(ByVal strString As String, ByVal cChar As Byte) As String
    'Rimuove gli spazi finali (nulli) di una variabile.
    'In genere è utilizzato con i valori recuperati dal registro di configurazione o da una funzione di Windows.
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(cChar))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
Public Sub HookWindow(SubClassForm As Form)
    ' If already subclass exit
    If OldWndProc <> 0 Then Exit Sub
    SubclassedhWnd = SubClassForm.hwnd
    ' Get previous hanldle procedure
    OldWndProc = GetWindowLong(SubClassForm.hwnd, GWL_WNDPROC)
    ' SubClassing Window
    SetWindowLong SubClassForm.hwnd, GWL_WNDPROC, AddressOf WindowProc
End Sub
Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim MenuItemStr As String * 128
Dim MenuHandle As Integer
    
    Select Case uMsg
        Case WM_RBUTTONDOWN
            ' Display Popup menu Edit
            Dim Pt As POINTAPI
            GetCursorPos Pt
            TrackPopupMenu GetSubMenu(Form1.hMenu, 1), TPM_RIGHTBUTTON Or TPM_LEFTALIGN, Pt.X, Pt.Y, 0, Form1.hwnd, Null
            WindowProc = 0
            Exit Function
        Case WM_MENUSELECT
            If lParam = 0 Then
                ' If menu is closed, clear StatusBar content.
                Form1.Label1.Caption = ""
                WindowProc = 0
                Exit Function
            End If
            ' LOWORD(wParam) = ID command or Position Menu
            MenuHandle = LOWORD(wParam)
            ' If selected menu is a popup, send as parameter the command position
            If (HIWORD(wParam) And MF_POPUP) = MF_POPUP Then
                ' Get string menu
                If GetMenuString(lParam, MenuHandle, MenuItemStr, 127, MF_BYPOSITION) = 0 Then: Exit Function
            Else
                ' Otherwise, send as parameter the command position
                ' Get string menu
                If GetMenuString(lParam, MenuHandle, MenuItemStr, 127, MF_BYCOMMAND) = 0 Then: Exit Function
            End If
            ' Show command in StatusBar.
            sCaption = GetRightMenuName(MenuItemStr)
            Form1.Label1.Caption = "ID: " + CStr(LOWORD(wParam)) + ", Name: " + Replace(sCaption, vbTab, Space$(4))
            WindowProc = 0
            Exit Function
        Case WM_COMMAND
            ' LOWORD(wParam) = ID command
            MenuHandle = LOWORD(wParam)
            ' Get string menu
            If GetMenuString(Form1.hMenu, MenuHandle, MenuItemStr, 127, 0) = 0 Then: Exit Function
                sCaption = GetRightMenuName(MenuItemStr)
                MsgBox "ID: " + CStr(LOWORD(wParam)) + ", Name: " + Replace(sCaption, vbTab, Space$(4)), vbInformation, App.EXEName
                WindowProc = 0
            Exit Function
        ' Only in VBIDE to display menu bar
        Case WM_SETFOCUS
            If Form1.hMenu <> 0 Then
                SetMenu Form1.hwnd, Form1.hMenu
                WindowProc = 0
                Exit Function
            End If
        Case WM_DESTROY
            'Since DefWindowProc doesn't automatically call
            'PostQuitMessage (WM_QUIT). We need to do it ourselves.
            'You can use DestroyWindow to get rid of the window manually.
            PostQuitMessage 0
            WindowProc = 0
            Exit Function
        Case WM_NCDESTROY
            ' Befor Destroy UnHook
            UnHookWindow
            WindowProc = 0
            Exit Function
    End Select
    
    ' Call Default Window procedure
    WindowProc = CallWindowProc(OldWndProc, hwnd, uMsg, wParam, lParam)

End Function
Public Sub UnHookWindow()
    ' If not subclassing exit
    If OldWndProc = 0 Then Exit Sub
    ' Restore default value
    SetWindowLong SubclassedhWnd, GWL_WNDPROC, OldWndProc
    OldWndProc = 0
End Sub



Public Function GetRightMenuName(ByVal strString As String) As String
Dim sCapt As String

    sCapt = StripTerminator(strString, 0)
    sCapt = StripTerminator(strString, 8)
    sCapt = StripTerminator(strString, 9)
    sCapt = StripTerminator(strString, 10)
    
    GetRightMenuName = sCapt
        
End Function
