Attribute VB_Name = "modApiMsgBox"
Option Explicit

'private const and variables for subclassing
Private Const HCBT_ACTIVATE = 5
Private Const WH_CBT = 5
Private CurrenthHook As Long

'private SetWindowPos flags
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOREPOSITION = &H200
Private Const SWP_NOSIZE = &H1

'private GetWindow flags
Private Const GWL_HINSTANCE = (-6)

'API functions from USER32

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
  (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowPos Lib "user32" _
  (ByVal hWnd As Long, _
   ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
   ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    
Private Declare Function MessageBoxEx Lib "user32" Alias "MessageBoxExA" _
  (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, _
   ByVal uType As Long, ByVal wLanguageId As Long) As Long

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
  (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
  
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long


'workaround function for hook routine
Private Function FARPROC(ByVal pfn As Long) As Long
  'Procedure that receives and returns
  'the passed value of the AddressOf operator.
  'This workaround is needed as you can't assign
  'AddressOf directly to a member of a user-
  'defined type, but you can assign it to another
  'long and use that (as returned here)
   FARPROC = pfn
End Function


'first par : ByVal hWnd As Long,
Public Function MsgBoxNew(ByVal Prompt As String, _
                          Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                          Optional ByVal Title As String = "", _
                          Optional ByVal HelpFile As String, _
                          Optional ByVal Context, _
                          Optional ByVal CenterForm As Boolean = True) As VbMsgBoxResult
  Dim ret As Long, hInst As Long, Thread As Long
  Dim hWnd As Long
  
  'Set up the CBT hook
  hWnd = GetForegroundWindow
  'ParenthWnd = hwnd

  hInst = GetWindowLong(hWnd, GWL_HINSTANCE)
  Thread = GetCurrentThreadId()

  If CenterForm Then
    CurrenthHook = SetWindowsHookEx(WH_CBT, FARPROC(AddressOf WinProcCenterForm), hInst, Thread)
  Else
    CurrenthHook = SetWindowsHookEx(WH_CBT, FARPROC(AddressOf WinProcCenterScreen), hInst, Thread)
  End If
  ret = MessageBoxEx(hWnd, Prompt, Title, Buttons, 0)
  MsgBoxNew = ret
End Function

Private Function WinProcCenterScreen(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim PosX As Long, PosY As Long
  Dim dlgWidth As Long, dlgHeight As Long

  If lMsg = HCBT_ACTIVATE Then
    
    'position the msgbox
    Call ComputeWindowPos(wParam, 0, MM_SCREEN_CENTER, 0, PosX, PosY, dlgWidth, dlgHeight)
    SetWindowPos wParam, 0, PosX, PosY, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
    
    'release the CBT hook
    UnhookWindowsHookEx CurrenthHook
  
  End If
  WinProcCenterScreen = False
End Function

Private Function WinProcCenterForm(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim PosX As Long, PosY As Long
  Dim dlgWidth As Long, dlgHeight As Long
  
  If lMsg = HCBT_ACTIVATE Then
    
    'position the msgbox
    Call ComputeWindowPos(wParam, GetParent(wParam), MM_PARENT_CENTER, 0, PosX, PosY, dlgWidth, dlgHeight)
    SetWindowPos wParam, 0, PosX, PosY, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
    
    'release the CBT hook
    UnhookWindowsHookEx CurrenthHook
  
  End If
  WinProcCenterForm = False
End Function

