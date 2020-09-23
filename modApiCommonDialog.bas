Attribute VB_Name = "modApiCommonDialog"
'---------------------------------------------------------------------
'  Module  m o d A p i C o m m o n D i a l o g
'
'  Author           : J.-C. Stritt
'  Last update      : 23-SEP-2002
'  First release    : 25-AUG-2002
'  Environment      : Visual Basic 6.0 SP5
'  Operating system : Windows XP
'
'  Goal             : define all often used API structures
'                     and functions to common dialog boxes
'                     include :
'                     - show color
'                     - show font
'                     - show page setup
'                     - show printer
'                     - show file open or save
'                     - show folder
'
'  Remarks          : - ok to XP like design
'
'                     - works on Win95/98/ME/NT4/2000/XP
'                       (only checked on XP and ME)
'
'                     - ok to set position of dialog boxes
'                       with 5 choices :
'                       a) no move (last default position)
'                       b) move to the screen centre
'                       c) move to the parent form centre
'                       d) move to parent shiftted with a given value
'                       e) move to last mouse position
'                          shiftted with a given value (default)
'---------------------------------------------------------------------
Option Explicit

'private constant and var
Private Const DEFAULT_SHIFT As Long = 30 'default shift in pixels
Private CurMoveShiftValue As Long        'memorize the current given shift value
Private CurMoveMode As Long              'memorize the current move mode
Private CurOwner As Long                 'memorize the current owner (only for ShowFolder)

'some windows constants
Private Const WM_INITDIALOG As Long = &H110
Private Const WM_USER = &H400
Private Const WM_SIZE = &H5
Private Const SIZE_RESTORED = 0

'ShowColor constants and types structures
Public Enum ColorFlag
  CC_RGBINIT = &H1
  CC_FULLOPEN = &H2
  CC_PREVENTFULLOPEN = &H4
  CC_SHOWHELP = &H8
  CC_ENABLEHOOK = &H10
  CC_ENABLETEMPLATE = &H20
  CC_ENABLETEMPLATEHANDLE = &H40
  CC_SOLIDCOLOR = &H80
  CC_ANYCOLOR = &H100
End Enum

Private Type TCOLORDLG
  lStructSize     As Long
  hwndOwner       As Long
  hInstance       As Long
  rgbResult       As Long
  lpCustColors    As Long
  Flags           As Long
  lCustData       As Long
  lpfnHook        As Long
  lpTemplateName  As String
End Type


'ShowFont constants and types structures
Private Const REGULAR_FONTTYPE = &H400
Private Const FW_BOLD = 700

Public Enum FontFlag
  CF_ANSIONLY = &H400             'show only windows or Unicode fonts
  CF_APPLY = &H200                'show the "Apply" Button
  CF_BOTH = &H3                   'show printer and screen fonts
  CF_EFFECTS = &H100              'show effets (underline and strikethru)
  CF_ENABLEHOOK = &H8             'set the hook (callback) routine
  CF_ENABLETEMPLATE = &H10        'use template
  CF_ENABLETEMPLATEHANDLE = &H20  'tamplate handle (hInstance)
  CF_FIXEDPITCHONLY = &H4000      'show only fixed pitch fonts
  CF_FORCEFONTEXIST = &H10000     'font must exist flag
  CF_INITTOLOGFONTSTRUCT = &H40   'initialize with logfont structure
  CF_LIMITSIZE = &H2000           'limit size between nSizeMin and nSizeMax
  CF_NOOEMFONTS = &H800           'not show OEM fonts
  CF_NOFACESEL = &H80000          'no face selection
  CF_NOSCRIPTSEL = &H800000       'no script font selection
  CF_NOSIZESEL = &H200000         'no size setting
  CF_NOSIMULATIONS = &H1000       'not show an example
  CF_NOSTYLESEL = &H100000        'not set standard style
  CF_NOVECTORFONTS = &H800        'not show vector fonts
  CF_NOVERTFONTS = &H1000000      'not show vertical fonts
  CF_PRINTERFONTS = &H2           'show printer fonts
  CF_SCALABLEONLY = &H20000       'show only scalable fonts
  CF_SCREENFONTS = &H1            'show only screen fonts
  CF_SCRIPTSONLY = &H400          'show only script fonts
  CF_SELECTSCRIPT = &H400000      'select script fonts
  CF_SHOWHELP = &H4               'show the help buttton
  CF_TTONLY = &H40000             'show only truetype fonts
  CF_USESTYLE = &H80              'use information in lpStyle variable on dialog loading
  CF_WYSIWYG = &H8000             'use only fonts for screen and printer
End Enum

Public Const CF_DEFAULT = CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT
Public Const CF_STANDARD = CF_DEFAULT Or CF_BOTH
Public Const CF_SCREEN = CF_DEFAULT Or CF_SCREENFONTS Or CF_ANSIONLY
Public Const CF_PRINTER = CF_DEFAULT Or CF_PRINTERFONTS

Private Type TFONTDLG
  lStructSize As Long
  hwndOwner As Long          'caller's window handle
  hdc As Long                'printer DC/IC or NULL
  lpLogFont As Long          'ptr. to a TLOGFONT struct
  iPointSize As Long         '10 * size in points of selected font
  Flags As Long              'enum. type flags
  rgbColors As Long          'returned text color
  lCustData As Long          'data passed to hook fn.
  lpfnHook As Long           'ptr. to hook function
  lpTemplateName As String   'custom template name
  hInstance As Long          'instance handle of.EXE that contains cust. dlg. template
  lpszStyle As String        'return the style field here must be LF_FACESIZE or bigger
  nFontType As Integer       'same value reported to the EnumFonts
                             'call back with the extra FONTTYPE_bits added
  MissingAlignment As Integer
  nSizeMin As Long           'minimum pt size allowed &
  nSizeMax As Long           'max pt size allowed if CF_LIMITSIZE is used
End Type


'ShowPageSetup type structure
Public Type TPAGESETUPDLG
  lStructSize As Long
  hwndOwner As Long
  hDevMode As Long
  hDevNames As Long
  Flags As Long
  ptPaperSize As TPOINTAPI
  rtMinMargin As TRECT
  rtMargin As TRECT
  hInstance As Long
  lCustData As Long
  lpfnPageSetupHook As Long
  lpfnPagePaintHook As Long
  lpPageSetupTemplateName As String
  hPageSetupTemplate As Long
End Type


'ShowPrinter constants and types structures
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Public Enum PrintDlgFlag
  PD_ALLPAGES = &H0
  PD_SELECTION = &H1
  PD_PAGENUMS = &H2
  PD_NOSELECTION = &H4
  PD_NOPAGENUMS = &H8
  PD_COLLATE = &H10
  PD_PRINTTOFILE = &H20
  PD_PRINTSETUP = &H40
  PD_NOWARNING = &H80
  PD_RETURNDC = &H100
  PD_RETURNIC = &H200
  PD_RETURNDEFAULT = &H400
  PD_SHOWHELP = &H800
  PD_ENABLEPRINTHOOK = &H1000
  PD_ENABLESETUPHOOK = &H2000
  PD_ENABLEPRINTTEMPLATE = &H4000
  PD_ENABLESETUPTEMPLATE = &H8000
  PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
  PD_ENABLESETUPTEMPLATEHANDLE = &H20000
  PD_USEDEVMODECOPIES = &H40000
  PD_USEDEVMODECOPIESANDCOLLATE = &H40000
  PD_DISABLEPRINTTOFILE = &H80000
  PD_HIDEPRINTTOFILE = &H100000
  PD_NONETWORKBUTTON = &H200000
End Enum

Private Type TPRINTDLG
  lStructSize As Long
  hwndOwner As Long
  hDevMode As Long
  hDevNames As Long
  hdc As Long
  Flags As Long
  nFromPage As Integer
  nToPage As Integer
  nMinPage As Integer
  nMaxPage As Integer
  nCopies As Integer
  hInstanceLow As Integer
  hInstanceHigh As Integer
  lCustDataLow As Integer
  lCustDataHigh As Integer
  lpfnPrintHookLow As Integer  'low word of lpfnPrintHook
  lpfnPrintHookHigh As Integer 'high word of lpfnPrintHook
  lpfnSetupHookLow As Integer
  lpfnSetupHookHigh As Integer
  lpPrintTemplateNameLow As Integer
  lpPrintTemplateNameHigh As Integer
  lpSetupTemplateNameLow As Integer
  lpSetupTemplateNameHigh As Integer
  hPrintTemplateLow As Integer
  hPrintTemplateHigh As Integer
  hSetupTemplateLow As Integer
  hSetupTemplateHigh As Integer
End Type

'In that UDT, I declared no long data type member after the five integers,
'so no alignment problem will occur. But when you set the lpfnPrintHook, you
'need to set the high and low word separately as below,
'
'pd.lpfnPrintHookHigh = callbackaddress(AddressOf centerwindow) / 65536
'pd.lpfnPrintHookLow = callbackaddress(AddressOf centerwindow) Mod 65536

Private Type DEVNAMES_TYPE
  wDriverOffset As Integer
  wDeviceOffset As Integer
  wOutputOffset As Integer
  wDefault As Integer
  extra As String * 100
End Type

Private Type DEVMODE_TYPE
  dmDeviceName As String * CCHDEVICENAME
  dmSpecVersion As Integer
  dmDriverVersion As Integer
  dmSize As Integer
  dmDriverExtra As Integer
  dmFields As Long
  dmOrientation As Integer
  dmPaperSize As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  dmFormName As String * CCHFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long
End Type


'ShowOpen and ShowSave type
Public Enum OpenSaveFlag
  OFN_ALLOWMULTISELECT = &H200
  OFN_CREATEPROMPT = &H2000
  OFN_ENABLEHOOK = &H20
  OFN_ENABLETEMPLATE = &H40
  OFN_ENABLETEMPLATEHANDLE = &H80
  OFN_EXPLORER = &H80000
  OFN_EXTENSIONDIFFERENT = &H400
  OFN_FILEMUSTEXIST = &H1000
  OFN_HIDEREADONLY = &H4
  OFN_LONGNAMES = &H200000
  OFN_NOCHANGEDIR = &H8
  OFN_NODEREFERENCELINKS = &H100000
  OFN_NOLONGNAMES = &H40000
  OFN_NONETWORKBUTTON = &H20000
  OFN_NOREADONLYRETURN = &H8000& 'correct value is 32768
  OFN_NOTESTFILECREATE = &H10000
  OFN_NOVALIDATE = &H100
  OFN_OVERWRITEPROMPT = &H2
  OFN_PATHMUSTEXIST = &H800
  OFN_READONLY = &H1
  OFN_SHAREAWARE = &H4000
  OFN_SHAREFALLTHROUGH = 2
  OFN_SHAREWARN = 0
  OFN_SHARENOWARN = 1
  OFN_SHOWHELP = &H10
End Enum

Public Const OFN_FILE_OPEN_FLAGS = _
             OFN_EXPLORER _
          Or OFN_LONGNAMES _
          Or OFN_CREATEPROMPT _
          Or OFN_NODEREFERENCELINKS _
          Or OFN_HIDEREADONLY _
          Or OFN_NOCHANGEDIR
          

Public Const OFN_FILE_SAVE_FLAGS = _
             OFN_EXPLORER _
          Or OFN_LONGNAMES _
          Or OFN_OVERWRITEPROMPT _
          Or OFN_HIDEREADONLY _
          Or OFN_NOCHANGEDIR

Public Const OFN_OPEN = True
Public Const OFN_SAVE = False

Private Type TFILENAMEDLG
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  strFilter As String
  strCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  strFile As String
  nMaxFile As Long
  strFileTitle As String
  nMaxFileTitle As Long
  strInitialDir As String
  strTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  strDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
  pvReserved As Long 'new Win2000 / WinXP members
  dwReserved As Long 'new Win2000 / WinXP members
  FlagsEx    As Long 'new Win2000 / WinXP members
End Type


'ShowFolder constants and types structures
Public Const MAX_PATH = 260
Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100) 'for win95
Public Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104) 'NT and plus
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)  'for win95
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)  'NT and plus

Public Enum BrowseFlag
  BIF_RETURNONLYFSDIRS = &H1      'Only file system directories
  BIF_DONTGOBELOWDOMAIN = &H2     'No network folders below domain level
  BIF_STATUSTEXT = &H4            'Includes status area in the dialog (for callback)
  BIF_RETURNFSANCESTORS = &H8     'Only returns file system ancestors
  BIF_EDITBOX = &H10              'Allows user to rename selection
  BIF_VALIDATE = &H20             'Insist on valid edit box result (or CANCEL)
  BIF_USENEWUI = &H40             'Version 5.0. Use the new user-interface.
                                  'Setting this flag provides the user with
                                  'a larger dialog box that can be resized.
                                  'It has several new capabilities including:
                                  'dialog box, reordering, context menus, new
                                  'folders, drag and drop capability within
                                  'the delete, and other context menu commands.
                                  'To use you must call OleInitialize or
                                  'CoInitialize before calling SHBrowseForFolder.
  BIF_BROWSEFORCOMPUTER = &H1000  'Only returns computers.
  BIF_BROWSEFORPRINTER = &H2000   'Only returns printers.
  BIF_BROWSEINCLUDEFILES = &H4000 'Browse for everything
End Enum

Private Type TBROWSEINFO
  hwndOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfnHook As Long
  lParam As Long
  iImage As Long
End Type


'API declarations for comdlg32
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" _
  (pChoosecolor As TCOLORDLG) As Long

Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" _
  (pChoosefont As TFONTDLG) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
  (pOpenfilename As TFILENAMEDLG) As Long
  
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" _
  (pOpenfilename As TFILENAMEDLG) As Long
  
Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" _
  (pPrintdlg As TPRINTDLG) As Long
  
Private Declare Function PageSetupDlg Lib "comdlg32.dll" Alias "PageSetupDlgA" _
  (pPagesetupdlg As TPAGESETUPDLG) As Long
  

'API declarations for shell32
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
  (lpBrowseInfo As TBROWSEINFO) As Long



'you must initialize the dialog move system once with this
Public Sub DlgInitMoveSystem(Optional ByVal MoveMode As MoveEnum = MM_MOUSE_SHIFT, _
                             Optional ByVal MoveShiftValue As Long = DEFAULT_SHIFT)
  CurMoveMode = MoveMode
  CurMoveShiftValue = MoveShiftValue
End Sub



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



'generic move routine for dialog boxes
Private Function DlgMoveProc(ByVal hWnd As Long, _
                             ByVal uMsg As Long, _
                             ByVal wParam As Long, _
                             ByVal lParam As Long) As Long
                           
  Dim DlgWidth As Long, DlgHeight As Long
  Dim PosX As Long, PosY As Long
  
  If (uMsg = WM_INITDIALOG) Or (uMsg = BFFM_INITIALIZED) Then
    If (wParam = 0) And (uMsg = WM_INITDIALOG) Then hWnd = GetParent(hWnd)
    If uMsg = BFFM_INITIALIZED Then
      Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal lParam)
    End If
    If CurMoveMode > 0 Then
      Call ComputeWindowPos(hWnd, CurOwner, CurMoveMode, CurMoveShiftValue, _
                            PosX, PosY, DlgWidth, DlgHeight)
      Call MoveWindow(hWnd, PosX, PosY, DlgWidth, DlgHeight, False)
      DlgMoveProc = 1
    End If
  End If
End Function



'ShowColor
Public Function ShowColor(ByRef frmOwner As Form, _
                          ByVal InitColor As Long, _
                          ByRef CustomColors() As Long, _
                          Optional ByVal ShowMode As Integer = 0) As Long
                          'ShowMode = 0 standard display
                          '         = 1 full dialog box with custom color editor
                          '         = 2 disable custom color edition
  Dim cc As TCOLORDLG
  Dim lReturn As Long
  Dim S As String

  CurOwner = frmOwner.hWnd
  
  'set some generic values
  cc.lStructSize = Len(cc)     'set the structure size
  cc.hwndOwner = frmOwner.hWnd 'set the owner
  cc.hInstance = App.hInstance 'set the application's instance
   
  'set flags
  cc.Flags = CC_ANYCOLOR
  Select Case ShowMode
    Case 1: cc.Flags = cc.Flags Or CC_FULLOPEN 'show custom colours
    Case 2: cc.Flags = cc.Flags Or CC_PREVENTFULLOPEN 'prevent display of custom colours?
  End Select
  
  'initial colour specified
  cc.Flags = cc.Flags Or CC_RGBINIT
  cc.rgbResult = InitColor
   
  'hook the dialog ?
   If CurMoveMode <> MM_NONE Then
     cc.Flags = cc.Flags Or CC_ENABLEHOOK
     cc.lpfnHook = FARPROC(AddressOf DlgMoveProc)
   End If

  'set the custom colors
  cc.lpCustColors = VarPtr(CustomColors(0))

  'show the 'Select Color'-dialog
  If ChooseColor(cc) = 1 Then
    ShowColor = cc.rgbResult
  Else
    ShowColor = -1
  End If
End Function



'ShowFont
Public Function ShowFont(ByRef frmOwner As Form, _
                         ByRef aFont As StdFont, _
                         ByRef aFontColor As Long, _
                         Optional ByVal Flags As Variant) As Boolean
                         
                         
  Dim cf As TFONTDLG, lfont As TLOGFONT, hMem As Long, pMem As Long
  Dim retval As Long
  
  ShowFont = False
  
  'check parameters
  CurOwner = frmOwner.hWnd
  If IsMissing(Flags) Then Flags = CF_STANDARD
  
  'set some generic values
  cf.lStructSize = Len(cf)     'set the structure size
  cf.hwndOwner = frmOwner.hWnd 'set the owner
  cf.hInstance = App.hInstance 'set the application's instance
  cf.Flags = Flags
  
  'hook the dialog ?
  If CurMoveMode <> MM_NONE Then
    cf.Flags = cf.Flags Or CF_ENABLEHOOK
    cf.lpfnHook = FARPROC(AddressOf DlgMoveProc)
  End If
  
  'transfer stdfont to logfont
  Call StdFontToLogFont(frmOwner.hdc, aFont, lfont)
  hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lfont))
  pMem = GlobalLock(hMem)                  'lock and get pointer
  CopyMemory ByVal pMem, lfont, Len(lfont) 'copy structure's contents into block
  
  'initialize dialog box
  cf.hdc = frmOwner.hdc           'device context of owner form
  cf.lpLogFont = pMem             'pointer to TLOGFONT memory block buffer
  cf.iPointSize = aFont.Size * 10 'point font (in units of 1/10 point)
  cf.rgbColors = aFontColor       'font color
  cf.nFontType = REGULAR_FONTTYPE 'regular font type i.e. not bold or anything
  cf.nSizeMin = 6                 'minimum point size
  cf.nSizeMax = 72                'maximum point size
  
  'now, call the function. If successful, copy the TLOGFONT structure back into the structure
  'and then print out the attributes we mentioned earlier that the user selected.
  retval = ChooseFont(cf)  ' open the dialog box
  If retval <> 0 Then  ' success
    CopyMemory lfont, ByVal pMem, Len(lfont)  ' copy memory back
    ' Now make the fixed-length string holding the font name into a "normal" string.
    aFont.Name = Left(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
    aFont.Size = Round(cf.iPointSize / 10, 0)
    aFont.Bold = CBool(lfont.lfWeight >= FW_BOLD)
    aFont.Italic = CBool(lfont.lfItalic)
    aFont.Underline = CBool(lfont.lfUnderline)
    aFont.Strikethrough = CBool(lfont.lfStrikeOut)
    aFontColor = cf.rgbColors
    ShowFont = True
  End If
  'deallocate the memory block we created earlier. Note that this must
  'be done whether the function succeeded or not.
  retval = GlobalUnlock(hMem) 'destroy pointer, unlock block
  retval = GlobalFree(hMem)   'free the allocated memory
End Function



'ShowPageSetup
Public Function ShowPageSetup(ByRef frmOwner As Form, _
                              Optional ByVal Flags As Variant) As Long
  Dim ps As TPAGESETUPDLG
  
  'check parameter
  CurOwner = frmOwner.hWnd
  If IsMissing(Flags) Then Flags = 0
  
  'set some generic values
  With ps
    .lStructSize = Len(ps)      'set the structure size
    .hwndOwner = frmOwner.hWnd  'set the owner
    .hInstance = App.hInstance  'set the application's instance
    .Flags = Flags
    
    'hook the dialog ?
    If CurMoveMode <> MM_NONE Then
      .Flags = .Flags Or PD_ENABLESETUPHOOK
      .lpfnPageSetupHook = FARPROC(AddressOf DlgMoveProc)
    End If
  End With
  
  'show the page setup dialog
  If PageSetupDlg(ps) Then
    ShowPageSetup = 0
  Else
    ShowPageSetup = -1
  End If
End Function



'ShowPrinter
Public Sub ShowPrinter(ByRef frmOwner As Form, _
                       Optional ByVal Flags As Variant)
  'some code by Donald Grover
  Dim pd As TPRINTDLG
  Dim DevMode As DEVMODE_TYPE
  Dim DevName As DEVNAMES_TYPE

  Dim lpDevMode As Long, lpDevName As Long
  Dim bReturn As Integer
  Dim objPrinter As Printer, NewPrinterName As String

  'check parameter
  CurOwner = frmOwner.hWnd
  If IsMissing(Flags) Then Flags = 0
  
  With pd
    'set some generic values
    .lStructSize = Len(pd)     'set the structure size
    .hwndOwner = frmOwner.hWnd 'set the owner
    .hInstanceHigh = App.hInstance / 65536  'set the application's instance
    .hInstanceLow = App.hInstance Mod 65536 'set the application's instance
    .Flags = Flags
  
    'hook the dialog ?
    If CurMoveMode <> MM_NONE Then
      .Flags = .Flags Or PD_ENABLEPRINTHOOK
      .lpfnPrintHookHigh = FARPROC(AddressOf DlgMoveProc) / 65536
      .lpfnPrintHookLow = FARPROC(AddressOf DlgMoveProc) Mod 65536
    End If
  End With
  
  'use PrintDialog to get the handle to a memory
  'block with a DevMode and DevName structures

  On Error Resume Next
  'set the current orientation and duplex setting
  DevMode.dmDeviceName = Printer.DeviceName
  DevMode.dmSize = Len(DevMode)
  DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
  DevMode.dmPaperWidth = Printer.Width
  DevMode.dmOrientation = Printer.Orientation
  DevMode.dmPaperSize = Printer.PaperSize
  DevMode.dmDuplex = Printer.Duplex
  On Error GoTo 0

  'Allocate memory for the initialization hDevMode structure
  'and copy the settings gathered above into this memory
  pd.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
  lpDevMode = GlobalLock(pd.hDevMode)
  If lpDevMode > 0 Then
    CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
    bReturn = GlobalUnlock(pd.hDevMode)
  End If

  'Set the current driver, device, and port name strings
  With DevName
    .wDriverOffset = 8
    .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
    .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
    .wDefault = 0
  End With

  With Printer
    DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
  End With

  'Allocate memory for the initial hDevName structure
  'and copy the settings gathered above into this memory
  pd.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
  lpDevName = GlobalLock(pd.hDevNames)
  If lpDevName > 0 Then
    CopyMemory ByVal lpDevName, DevName, Len(DevName)
    bReturn = GlobalUnlock(lpDevName)
  End If

  'Call the print dialog up and let the user make changes
  If PrintDlg(pd) <> 0 Then

    'First get the DevName structure.
    lpDevName = GlobalLock(pd.hDevNames)
    CopyMemory DevName, ByVal lpDevName, 45
    bReturn = GlobalUnlock(lpDevName)
    GlobalFree pd.hDevNames

    'Next get the DevMode structure and set the printer
    'properties appropriately
    lpDevMode = GlobalLock(pd.hDevMode)
    CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
    bReturn = GlobalUnlock(pd.hDevMode)
    GlobalFree pd.hDevMode
    NewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
    If Printer.DeviceName <> NewPrinterName Then
      For Each objPrinter In Printers
        If UCase$(objPrinter.DeviceName) = NewPrinterName Then
          Set Printer = objPrinter
          'set printer toolbar name at this point
        End If
      Next
    End If

    On Error Resume Next
    'Set printer object properties according to selections made
    'by user
    Printer.Copies = DevMode.dmCopies
    Printer.Duplex = DevMode.dmDuplex
    Printer.Orientation = DevMode.dmOrientation
    Printer.PaperSize = DevMode.dmPaperSize
    Printer.PrintQuality = DevMode.dmPrintQuality
    Printer.ColorMode = DevMode.dmColor
    Printer.PaperBin = DevMode.dmDefaultSource
    On Error GoTo 0
  End If
End Sub
 
 
 
'ShowFileOpenSave
Public Function ShowFileOpenSave(ByRef frmOwner As Form, _
                                 Optional ByVal OpenFlag As Variant, _
                                 Optional ByVal DialogTitle As Variant, _
                                 Optional ByVal FullName As Variant, _
                                 Optional ByVal Filter As Variant, _
                                 Optional ByVal FilterIndex As Variant, _
                                 Optional ByVal DefaultExt As Variant, _
                                 Optional ByRef Flags As Variant) As String
  'In:
  'OpenFlag    ->> boolean (True=Open File / False=Save As)
  'DialogTitle ->> title for the dialog
  'FullName    ->> default file name with path (last file)
  'Filter      ->> a set of file filters, set up by calling AddFilterItem
  'FilterIndex ->> 1-based integer indicating which filter set to use
  'DefaultExt  ->> extension (for file saves)
  'Flags       ->> one or more of the OFN constants
  
  'Out:
  'Return Value ->> either Null or the selected filename
  
  Dim OFN As TFILENAMEDLG
  Dim fResult As Boolean
  Dim strFileName As String, strFileTitle As String
  Dim OnlyPath As String, OnlyFileName As String
    
  'check parameters
  CurOwner = frmOwner.hWnd
  If IsMissing(OpenFlag) Then OpenFlag = True
  If IsMissing(DialogTitle) Then DialogTitle = ""
  If IsMissing(FullName) Then FullName = ""
  If IsMissing(Filter) Then Filter = ""
  If IsMissing(FilterIndex) Then FilterIndex = 1
  If IsMissing(DefaultExt) Then DefaultExt = ""
  If IsMissing(Flags) Then Flags = IIf(OpenFlag, OFN_FILE_OPEN_FLAGS, OFN_FILE_SAVE_FLAGS)
  
  'extract path and file name
  OnlyPath = ExtractPathName(FullName)
  If OnlyPath = "" Then OnlyPath = CurDir
  OnlyFileName = ExtractFileName(FullName)

  'allocate string space for the returned strings.
  strFileName = Left(OnlyFileName & String(256, 0), 256)
  strFileTitle = String(256, 0)
  
  'set up the data structure before you call the function
  With OFN
    If IsWin2000Plus() Then
      .lStructSize = Len(OFN)
      .FlagsEx = 0 'any FlagsEx values desired
    Else
      .lStructSize = Len(OFN) - 12
    End If
    .hwndOwner = frmOwner.hWnd 'set the owner
    .hInstance = App.hInstance 'set the application's instance
    .strFilter = Filter
    .nFilterIndex = FilterIndex
    .strFile = strFileName
    .nMaxFile = Len(strFileName)
    .strFileTitle = strFileTitle
    .nMaxFileTitle = Len(strFileTitle)
    .strTitle = DialogTitle
    .strDefExt = DefaultExt
    .strInitialDir = OnlyPath
    .strCustomFilter = String(255, 0)
    .nMaxCustFilter = 255
    .lpfnHook = 0
    .Flags = Flags
    'hook the dialog ?
    If CurMoveMode <> MM_NONE Then
      .Flags = .Flags Or OFN_ENABLEHOOK
      .lpfnHook = FARPROC(AddressOf DlgMoveProc)
    End If
  End With
  
  'open the dialog box
  If OpenFlag Then
    fResult = GetOpenFileName(OFN)
  Else
    fResult = GetSaveFileName(OFN)
  End If
  
  'return file name
  If fResult Then
    ShowFileOpenSave = TrimNull(OFN.strFile)
  Else
    ShowFileOpenSave = FullName
  End If
End Function



'ShowFolder
Public Function ShowFolder(ByRef frmOwner As Form, _
                           Optional ByVal DialogTitle As Variant, _
                           Optional ByVal InitialDir As Variant, _
                           Optional ByRef Flags As Variant) As String
  Dim bi As TBROWSEINFO
  Dim strPath As String * MAX_PATH, sFolder As String
  Dim pidl As Long, lpSelPath As Long
  
  'check parameters
  CurOwner = frmOwner.hWnd
  If IsMissing(DialogTitle) Then DialogTitle = ""
  If IsMissing(InitialDir) Then InitialDir = CurDir
  If IsMissing(Flags) Then Flags = BIF_RETURNONLYFSDIRS
  sFolder = InitialDir
  
  'allocate string space for the returned path
  strPath = Left(InitialDir & String(MAX_PATH, 0), MAX_PATH)
  
  With bi
    .hwndOwner = frmOwner.hWnd 'set the owner
    .pidlRoot = 0&             'root folder = Desktop
    .lpszTitle = DialogTitle
    .ulFlags = Flags
    .lpfnHook = FARPROC(AddressOf DlgMoveProc)
    lpSelPath = LocalAlloc(LPTR, Len(strPath) + 1)
    CopyMemory ByVal lpSelPath, ByVal strPath, Len(strPath) + 1
    .lParam = lpSelPath
  End With
  pidl = SHBrowseForFolder(bi) 'display the dialog
   
  'parse the result
  If pidl Then
    If SHGetPathFromIDList(ByVal pidl, ByVal strPath) Then
      sFolder = TrimNull(strPath)
    End If
    Call CoTaskMemFree(pidl)
  End If
  Call LocalFree(lpSelPath)
  ShowFolder = AddBackSlash(sFolder)
End Function
