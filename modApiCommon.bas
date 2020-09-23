Attribute VB_Name = "modApiCommon"
'---------------------------------------------------------------------
'  Module  m o d A p i C o m m o n
'
'  Author           : J.-C. Stritt
'  Last update      : 25-SEP-2002
'  First release    : 25-AUG-2002
'  Environment      : Visual Basic 6.0 SP5
'  Operating system : Windows XP
'
'  Goal             : define some often used API structures
'                     and functions
'
'  API declarations for comdlg32
'    Sub InitCommonControls
'
'  API declarations for ole32
'    Sub CoTaskMemFree
'
'  API declarations for user32
'    Function GetForegroundWindow
'    Function GetParent
'    Function FindWindow
'    Function GetWindowRect
'    Function GetClientRect
'    Function MoveWindow
'    Sub GetCursorPos
'    Function SendMessage
'
'  API declarations for kernel32
'    Function GetCurrentThreadId
'    Sub CopyMemory
'    Function GlobalLock
'    Function GlobalUnlock
'    Function LocalAlloc
'    Function LocalFree
'    Function GlobalAlloc
'    Function GlobalFree
'    Function GetVersionEx
'
'  Special routines :
'    Function AddFilterItem
'
'    Function TrimNull
'    Function TrimBackSlash
'    Function AddBackSlash
'
'    Function ExtractPathName
'    Function ExtractFileName
'
'    Function IsWin2000Plus
'
'    Function GetScreenUsableArea
'    Sub ComputeWindowPosXY
'    Sub CorrectWindowPosY
'
'    Sub StdFontToLogFont
'---------------------------------------------------------------------
Option Explicit

'some most used structures
Public Type TPOINTAPI
  x As Long
  y As Long
End Type

Public Type TRECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type TLOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * 31
End Type


'move constants (dialog box position)
Public Enum MoveEnum
  MM_NONE = 0          'no move
  MM_SCREEN_CENTER = 1 'move to screen centre
  MM_PARENT_CENTER = 2 'move to parent form centre
  MM_PARENT_SHIFT = 3  'move to parent with a top and left shift
  MM_MOUSE_SHIFT = 4   'move to mouse position (vertical centered, left shift)
End Enum


'some private type and const for OS version
Private Type TOSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Private Const VER_PLATFORM_WIN32s = 0        'Windows 3.x is running, using Win32s
Private Const VER_PLATFORM_WIN32_WINDOWS = 1 'Windows 95 or 98 is running.
Private Const VER_PLATFORM_WIN32_NT = 2      'Windows NT is running.


'memory allocation constants
Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_ZEROINIT = &H40


'API declarations for comdlg32
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

'API declarations for ole32
Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pidl As Long)

'API declarations for user32
Public Declare Function GetForegroundWindow Lib "user32" () As Long

Public Declare Function GetParent Lib "user32" _
  (ByVal hWnd As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
  (ByVal lpClassName As String, _
   ByVal lpWindowName As String) As Long
   
Public Declare Function GetWindowRect Lib "user32" _
  (ByVal hWnd As Long, _
   lpRect As TRECT) As Long

Public Declare Function GetClientRect Lib "user32" _
  (ByVal hWnd As Long, lpRect As TRECT) As Long

Public Declare Function MoveWindow Lib "user32" _
  (ByVal hWnd As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal bRepaint As Long) As Long
   
Public Declare Sub GetCursorPos Lib "user32" _
  (lpPoint As TPOINTAPI)
  
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private Const SPI_GETWORKAREA& = 48
Private Declare Function SystemParametersInfo Lib "user32" _
   Alias "SystemParametersInfoA" _
  (ByVal uAction As Long, _
   ByVal uParam As Long, _
   lpvParam As Any, _
   ByVal fuWinIni As Long) As Long

   
'API declarations for kernel32
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
  (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
  
Public Declare Function GlobalLock Lib "kernel32" _
  (ByVal hMem As Long) As Long

Public Declare Function GlobalUnlock Lib "kernel32" _
  (ByVal hMem As Long) As Long

Public Declare Function LocalAlloc Lib "kernel32" _
   (ByVal uFlags As Long, _
    ByVal uBytes As Long) As Long
    
Public Declare Function LocalFree Lib "kernel32" _
   (ByVal hMem As Long) As Long

Public Declare Function GlobalAlloc Lib "kernel32" _
  (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
  
Public Declare Function GlobalFree Lib "kernel32" _
  (ByVal hMem As Long) As Long

Public Declare Function MulDiv Lib "kernel32" _
  (ByVal nNumber As Long, _
   ByVal nNumerator As Long, _
   ByVal nDenominator As Long) As Long

Public Declare Function GetVersionEx Lib "kernel32" _
  Alias "GetVersionExA" (lpVersionInformation As TOSVERSIONINFO) As Long


'API declarations for gid32
Public Declare Function GetDeviceCaps Lib "gdi32" _
  (ByVal hdc As Long, _
   ByVal nIndex As Long) As Long



'add a new entry onto a file filter string
Public Function AddFilterItem(ByVal strFilter As String, _
                              ByVal strDescription As String, _
                              Optional ByVal varItem As Variant) As String
    If IsMissing(varItem) Then varItem = "*.*"
    AddFilterItem = strFilter & strDescription & vbNullChar _
                              & varItem & vbNullChar
End Function



'suppress the ending null char in a string
Public Function TrimNull(ByVal strItem As String) As String
  Dim intPos As Integer
  intPos = InStr(strItem, vbNullChar)
  If intPos > 0 Then
    TrimNull = Left(strItem, intPos - 1)
  Else
    TrimNull = strItem
  End If
End Function

'suppress the last backslash in a path name
Public Function TrimBackSlash(ByVal strPath As String) As String
  TrimBackSlash = strPath
  If Len(strPath) > 0 Then
    If Right(strPath, 1) = "\" Then
      TrimBackSlash = Left(strPath, Len(strPath) - 1)
    End If
  End If
End Function

'add a backslash if necessary at end of a path name
Function AddBackSlash(ByVal strPath As String) As String
  If (Len(strPath) > 0) And (Right(strPath, 1) <> "\") Then
    AddBackSlash = strPath & "\"
  Else
    AddBackSlash = strPath
  End If
End Function

'extract all parts from a file name (DriveLetter, DirPath, FName and Extension)
Private Sub FileNameExtractAllParts(ByVal AccessName As String, _
                                    ByRef DriveLetter As String, _
                                    ByRef DirPath As String, _
                                    ByRef FName As String, _
                                    ByRef Extension As String)
  Dim PathLength As Integer
  Dim ThisLength As Integer
  Dim Offset As Integer
  Dim FileNameFound As Boolean, DotFound As Boolean

  DriveLetter = ""
  DirPath = ""
  FName = ""
  Extension = ""

  If Mid(AccessName, 2, 1) = ":" Then 'find the drive letter.
    DriveLetter = Left(AccessName, 2)
    AccessName = Mid(AccessName, 3)
  End If

  PathLength = Len(AccessName)

  DotFound = False
  For Offset = PathLength To 1 Step -1 'find the next delimiter.
    Select Case Mid(AccessName, Offset, 1)
      Case "."
        'this indicates either an extension or a . or a ..
        If Not DotFound Then
          ThisLength = Len(AccessName) - Offset
          If ThisLength >= 1 Then ' Extension
            Extension = Mid(AccessName, Offset, ThisLength + 1)
          End If
          AccessName = Left(AccessName, Offset - 1)
          DotFound = True
        End If
      Case "\"
        'this indicates a path delimiter.
        ThisLength = Len(AccessName) - Offset
        If ThisLength >= 1 Then ' Filename
          FName = Mid(AccessName, Offset + 1, ThisLength)
          AccessName = Left(AccessName, Offset)
          FileNameFound = True
          Exit For
        End If
    End Select
  Next Offset
  If FileNameFound = False Then
    FName = AccessName
  Else
    DirPath = AccessName
  End If
  If (Len(Extension) = 0) And (Len(FName) > 0) And (InStr(FName, ".") = 0) Then
    DirPath = DirPath & FName
    FName = ""
  End If
End Sub

'extract path name only from a full file name with path
Public Function ExtractPathName(ByVal FullFileName As String) As String
  Dim DriveLetter As String, DirPath As String, FName As String, Extension As String
  Call FileNameExtractAllParts(FullFileName, DriveLetter, DirPath, FName, Extension)
  ExtractPathName = AddBackSlash(DriveLetter & DirPath)
End Function

'extract file name only from a full file name with path
Public Function ExtractFileName(ByVal FullFileName As String) As String
  Dim DriveLetter As String, DirPath As String, FName As String, Extension As String
  Call FileNameExtractAllParts(FullFileName, DriveLetter, DirPath, FName, Extension)
  ExtractFileName = FName & Extension
End Function



'returns True if running Win2000 or WinXP
Public Function IsWin2000Plus() As Boolean
  Dim OSV As TOSVERSIONINFO
  IsWin2000Plus = False
  OSV.dwOSVersionInfoSize = Len(OSV)
  If GetVersionEx(OSV) = 1 Then
    'PlatformId contains a value representing the OS.
     IsWin2000Plus = (OSV.dwPlatformId = VER_PLATFORM_WIN32_NT) And _
                     (OSV.dwMajorVersion >= 5)
  End If
End Function



'return usable screen area (without taskbar size)
Public Sub GetScreenUsableArea(ByRef ScrRect As TRECT)
  Call SystemParametersInfo(SPI_GETWORKAREA, 0&, ScrRect, 0&)
End Sub



'compute window X-Y position
Public Sub ComputeWindowPos(ByVal hWnd As Long, _
                            ByVal ParhWnd As Variant, _
                            ByVal MoveMode As MoveEnum, _
                            ByVal PixShift As Long, _
                            ByRef PosX As Long, _
                            ByRef PosY As Long, _
                            ByRef DlgWidth As Long, _
                            ByRef DlgHeight As Long)
  
  Dim ScrRect As TRECT, ScrWidth As Long, ScrHeight As Long
  Dim ParRect As TRECT, ParWidth As Long, ParHeight As Long
  Dim DlgRect As TRECT, P As TPOINTAPI
  
  'get usable screen size
  Call GetScreenUsableArea(ScrRect)
  ScrWidth = ScrRect.Right - ScrRect.Left
  ScrHeight = ScrRect.Bottom - ScrRect.Top
  
  'get parent form size
  Call GetWindowRect(ParhWnd, ParRect)
  ParWidth = ParRect.Right - ParRect.Left
  ParHeight = ParRect.Bottom - ParRect.Top
  
  'get common dialog box size
  Call GetWindowRect(hWnd, DlgRect)
  DlgWidth = DlgRect.Right - DlgRect.Left
  DlgHeight = DlgRect.Bottom - DlgRect.Top
  
  'compute position
  Select Case MoveMode
    Case MM_SCREEN_CENTER
      PosX = ScrRect.Left + (ScrWidth - DlgWidth) \ 2
      PosY = ScrRect.Top + (ScrHeight - DlgHeight) \ 2
    Case MM_PARENT_CENTER
      PosX = ParRect.Left + ParWidth \ 2 - DlgWidth \ 2
      PosY = ParRect.Top + ParHeight \ 2 - DlgHeight \ 2
    Case MM_PARENT_SHIFT
      PosX = ParRect.Left + PixShift
      PosY = ParRect.Top - PixShift
    Case MM_MOUSE_SHIFT
      Call GetCursorPos(P)
      PosX = P.x + PixShift
      PosY = P.y - DlgHeight \ 2
      If (PosX + DlgWidth) > ScrRect.Right Then
        PosX = P.x - PixShift - DlgWidth
      End If
  End Select
   
  'correct X-Y position
  If PosX < ScrRect.Left Then PosX = ScrRect.Left
  If (PosX + DlgWidth) > ScrRect.Right Then
    PosX = ParRect.Left + ParWidth - DlgWidth - PixShift
    If (PosX + DlgWidth) > ScrRect.Right Then PosX = ScrRect.Right - DlgWidth - PixShift
    If PosX < ScrRect.Left Then PosX = ScrRect.Left
  End If
  If PosY < ScrRect.Top Then PosY = ScrRect.Top
  If (PosY + DlgHeight) > ScrRect.Bottom Then
    PosY = ScrRect.Bottom - DlgHeight
    If PosY < ScrRect.Top Then PosY = ScrRect.Top
  End If
End Sub



'correct top position of the active window
Public Sub CorrectWindowPosY()
  Dim hWnd As Long
  Dim ScrRect As TRECT, ScrWidth As Long, ScrHeight As Long
  Dim DlgRect As TRECT, DlgHeight As Long, DlgWidth As Long
  Dim PosX As Long, PosY As Long
  
  'get usable screen size
  Call GetScreenUsableArea(ScrRect)
  ScrWidth = ScrRect.Right - ScrRect.Left
  ScrHeight = ScrRect.Bottom - ScrRect.Top
    
  'get common dialog box size
  Call GetWindowRect(hWnd, DlgRect)
  DlgWidth = DlgRect.Right - DlgRect.Left
  DlgHeight = DlgRect.Bottom - DlgRect.Top
    
  'set to current pos
  PosX = DlgRect.Left
  PosY = DlgRect.Top
    
  'correct Y position
  If (PosY + DlgHeight) > ScrRect.Bottom Then
    PosY = ScrRect.Bottom - DlgHeight
    If PosY < ScrRect.Top Then PosY = ScrRect.Top
    Call MoveWindow(hWnd, PosX, PosY, DlgWidth, DlgHeight, True)
  End If
End Sub



'convert an OLE StdFont to a TLOGFONT structure
Public Sub StdFontToLogFont(ByVal hdc As Long, ByRef stFont As StdFont, ByRef lgFont As TLOGFONT)
  Const LOGPIXELSY = 90
  Const FW_NORMAL = 400
  Const FW_BOLD = 700
  Const DEFAULT_QUALITY = 0
  Const DEFAULT_PITCH = 0
  Const OUT_DEFAULT_PRECIS = 0
  Const CLIP_DEFAULT_PRECIS = 0
  With lgFont
    'from StdFont
    .lfFaceName = stFont.Name & vbNullChar 'string must be null-terminated
    .lfHeight = -MulDiv((stFont.Size), (GetDeviceCaps(hdc, LOGPIXELSY)), 72)
    .lfItalic = stFont.Italic
    .lfWeight = IIf(stFont.Bold, FW_BOLD, FW_NORMAL) 'aFont.Weight not used
    .lfUnderline = stFont.Underline
    .lfStrikeOut = stFont.Strikethrough
    .lfCharSet = stFont.Charset
    'default properties
    .lfWidth = 0                            'determine default width
    .lfEscapement = 0                       'angle between baseline and escapement vector
    .lfOrientation = 0                      'angle between baseline and orientation vector
    .lfQuality = DEFAULT_QUALITY            'default quality setting
    .lfPitchAndFamily = DEFAULT_PITCH       'default pitch
    .lfOutPrecision = OUT_DEFAULT_PRECIS    'default precision mapping
    .lfClipPrecision = CLIP_DEFAULT_PRECIS  'default clipping precision
  End With
End Sub

