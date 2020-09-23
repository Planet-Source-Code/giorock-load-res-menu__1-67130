VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Load RES Menu & Accelerators"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   433
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Width           =   8205
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

' Menu Function
Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function LoadMenuIndirect Lib "user32" Alias "LoadMenuIndirectA" (lpMenuTemplate As Any) As Long

' Accelerator Function
Private Type ACCEL
    fVirt As Byte
    key As Integer
    cmd As Integer
End Type
Private Declare Function LoadAccelerators Lib "user32" Alias "LoadAcceleratorsA" (ByVal hInstance As Long, ByVal lpTableName As String) As Long
Private Declare Function CreateAcceleratorTable Lib "user32" Alias "CreateAcceleratorTableA" (lpaccl As ACCEL, ByVal cEntries As Long) As Long
Private Declare Function DestroyAcceleratorTable Lib "user32" (ByVal haccel As Long) As Long
Private Declare Function TranslateAccelerator Lib "user32" Alias "TranslateAcceleratorA" (ByVal hwnd As Long, ByVal hAccTable As Long, lpMsg As MSG) As Long

' Constant VB Resource
Private Const VB_RES_MENU = &H4
Private Const VB_RES_ACCEL = &H9

' My Data
Public hMenu As Long
Private hAcc As Long
Private bQuit As Long


Private Sub Form_Load()
Dim s As String
Dim Acc() As ACCEL
Dim i As Integer
Dim wMsg As MSG
    
    ' Ensure Form visibility
    Show
    DoEvents
    
    ' In previous version of program, i use LoadMenu function
    ' to create a menu from resource
    ' The inconvenient, is that you must compile a project to
    ' view result
    ' So, i've got an idea....
    ' Use VB function LoadResData to obtain string, then
    ' pass string pointer at LoadMenuIndirect API call
    hMenu = LoadMenuIndirect(ByVal StrPtr(LoadResData(101, VB_RES_MENU)))
    
    ' Set the Menu handle at specified Window
    SetMenu hwnd, hMenu
    
    ' Draw and display menu
    DrawMenuBar hwnd
    
    ' In previous version of program, i use LoadAccelerators function
    ' to create an accelerator from resource
    ' The inconvenient, is that you must compile a project to
    ' view result
    ' So, i've got an idea....
    ' Use VB function LoadResData to obtain string, then
    ' .....
    
    s = CStr(LoadResData(104, VB_RES_ACCEL))
    
    ' Accelerators is stored in resource in 4 bytes length packets
    ' so, ((Len(s) / 4) - 1) obtain maximum array to create
    ReDim Acc((Len(s) / 4) - 1) As ACCEL

    ' I use ACCEL structure to copy bytes by string
    ' (Byte)ACCEL->fVirt = Specifies the accelerator flags
    ' (Integer)ACCEL->key = Specifies the accelerator key
    ' (Integer)ACCEL->cmd = Specifies the accelerator identifier
    For i = 0 To UBound(Acc)
        MoveMemory Acc(i).fVirt, ByVal StrPtr(Mid$(s, (i * 4) + 1, 1)), 1 ' copy 1 byte
        MoveMemory Acc(i).key, ByVal StrPtr(Mid$(s, (i * 4) + 2, 1)), 1   ' copy 1 byte
        MoveMemory Acc(i).cmd, ByVal StrPtr(Mid$(s, (i * 4) + 3, 2)), 2   ' copy 2 bytes
'        Debug.Print Acc(i).fVirt, Chr$(Acc(i).key), Acc(i).cmd
    Next i

    ' Finally, i can create Accelerator Table with this function
    hAcc = CreateAcceleratorTable(Acc(0), UBound(Acc) + 1)
    
    ' Clean up aray
    Erase Acc
    
    ' SubClass Window
    HookWindow Me
    
    Do While GetMessage(wMsg, 0, 0, 0) And Not bQuit
        ' The TranslateAccelerator function processes accelerator
        ' keys for menu commands
        TranslateAccelerator hwnd, hAcc, wMsg
        ' TranslateMessage takes keyboard messages and converts
        ' them to WM_CHAR for easier processing.
        TranslateMessage wMsg
        ' Dispatchmessage calls the default window procedure
        ' to process the window message. (WndProc)
        DispatchMessage wMsg
    Loop
    
End Sub


Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        Label1.Move 0, Me.ScaleHeight - Label1.Height, Me.ScaleWidth
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    ' Force 'Exit Do' to Quit
    ' otherwise executable don't shut down
    bQuit = True
    
    ' UnSubClass Window
    UnHookWindow
    
    ' Clean up menu and Accelerator
    DestroyAcceleratorTable hAcc
    DestroyMenu hMenu
    
    ' Destroy Form and free memory
    End
    Set Form1 = Nothing
    
End Sub


