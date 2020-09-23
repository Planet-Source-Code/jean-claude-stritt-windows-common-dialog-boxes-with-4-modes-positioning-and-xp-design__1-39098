VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test form to API functions"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      Height          =   645
      Left            =   210
      TabIndex        =   15
      Top             =   4410
      Width           =   645
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   435
      Left            =   4050
      TabIndex        =   7
      Top             =   1785
      Width           =   1380
   End
   Begin VB.TextBox txtShift 
      Height          =   330
      Left            =   4050
      TabIndex        =   14
      Top             =   2520
      Width           =   1380
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   225
      Index           =   4
      Left            =   210
      TabIndex        =   12
      Top             =   3885
      Width           =   2000
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   225
      Index           =   3
      Left            =   210
      TabIndex        =   11
      Top             =   3540
      Width           =   2000
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   8
      Top             =   2520
      Width           =   2000
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   225
      Index           =   1
      Left            =   210
      TabIndex        =   9
      Top             =   2865
      Width           =   2000
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   225
      Index           =   2
      Left            =   210
      TabIndex        =   10
      Top             =   3195
      Width           =   2000
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   435
      Left            =   4050
      TabIndex        =   6
      Top             =   1260
      Width           =   1380
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   435
      Left            =   4050
      TabIndex        =   5
      Top             =   735
      Width           =   1380
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   435
      Left            =   2205
      TabIndex        =   4
      Top             =   1260
      Width           =   1380
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   435
      Left            =   2205
      TabIndex        =   3
      Top             =   735
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   435
      Left            =   315
      TabIndex        =   2
      Top             =   1260
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   315
      TabIndex        =   1
      Top             =   735
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Release 1.2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   315
      TabIndex        =   0
      Top             =   105
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Shift value in pixels"
      Height          =   225
      Left            =   2310
      TabIndex        =   13
      Top             =   2580
      Width           =   1590
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------
'  Application test form
'
'  Author           : J.-C. Stritt
'  Last update      : 18-SEP-2002
'  First release    : 25-AUG-2002
'  Environment      : Visual Basic 6.0 SP5
'  Operating system : Windows XP
'
'  Goal             : check of all functions that
'                     use standard dialog boxes with
'                     API call and positioning
'
'  Remark           : ok to XP like design with a call to
'                     InitCommonControls API in the
'                     Form_Initialize() code. You must have
'                     a manifest file in the current application
'                     directory (here Projet1.exe.manifest).
'                     See doc on Microsoft sites.
'---------------------------------------------------------------------
'
Option Explicit

Dim CustomColors(0 To 15) As Long
Dim LastColor As Long
Dim LastFont As StdFont
Dim LastFontColor As Long
Dim LastFileName As String
Dim LastDir As String


Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  Dim cnt As Long
  
  'populate the custom colours with a series of gray shades
  For cnt = 240 To 15 Step -15
    CustomColors((cnt \ 15) - 1) = RGB(cnt, cnt, cnt)
  Next cnt
  
  'Set the captions
  Command1.Caption = "Color"
  Command2.Caption = "Font"
  Command3.Caption = "Page setup"
  Command4.Caption = "Print setup"
  Command5.Caption = "File Open"
  Command6.Caption = "File Save"
  Command7.Caption = "Browse Folder"
  
  Option1(0).Caption = "No move"
  Option1(1).Caption = "Screen center"
  Option1(2).Caption = "Parent form center"
  Option1(3).Caption = "Parent form shift"
  Option1(4).Caption = "Mouse shift"
  Option1(4).Value = True
  
  txtShift.Text = 30
  
  'set last variables
  LastColor = Me.BackColor
  Set LastFont = New StdFont
  LastFont.Name = "Arial"
  LastFont.Size = 12
  LastFontColor = vbBlack
  LastFileName = ""
  LastDir = CurDir
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set LastFont = Nothing
End Sub



Function OkAsc1(KeyAscii As Integer, LLim As String, HLim As String, _
                Optional Others As String = "") As Integer
  Dim Okay As Boolean, l As Integer, i As Integer
  Okay = KeyAscii <= vbKeySpace
  Okay = Okay Or (KeyAscii >= Asc(LLim)) And (KeyAscii <= Asc(HLim))
  l = Len(Others)
  If (Not Okay) And (Len(Others) > 0) Then
    For i = 1 To l
      Okay = Okay Or (KeyAscii = Asc(Mid(Others, i, 1)))
    Next i
  End If
  If Okay Then
    OkAsc1 = KeyAscii
  Else
    OkAsc1 = 0
    Beep
  End If
End Function

Private Sub txtShift_KeyPress(KeyAscii As Integer)
  KeyAscii = OkAsc1(KeyAscii, "0", "9", "+-")
End Sub

Private Function GetMoveMode() As MoveEnum
  Dim i As Integer
  GetMoveMode = MM_NONE
  For i = 0 To 4
    If Option1(i).Value Then
      Select Case i
        Case 0: GetMoveMode = MM_NONE
        Case 1: GetMoveMode = MM_SCREEN_CENTER
        Case 2: GetMoveMode = MM_PARENT_CENTER
        Case 3: GetMoveMode = MM_PARENT_SHIFT
        Case 4: GetMoveMode = MM_MOUSE_SHIFT
      End Select
    End If
  Next i
End Function



'set color
Private Sub Command1_Click()
  Dim NewColor As Long, ctl As Control
  Call DlgInitMoveSystem(GetMoveMode(), txtShift.Text)
  NewColor = ShowColor(Me, LastColor, CustomColors, 1)
  If NewColor <> -1 Then
    Me.BackColor = NewColor
    For Each ctl In Me.Controls
      ctl.BackColor = NewColor
    Next ctl
    LastColor = NewColor
  Else
    MsgBoxNew "Operation canceled"
  End If
End Sub


'set font
Private Sub Command2_Click()
  Call DlgInitMoveSystem(GetMoveMode(), txtShift.Text)
  If ShowFont(Me, LastFont, LastFontColor, CF_PRINTER) Then
    Call MsgBoxNew("New font is : " & LastFont.Name & " in " & LastFont.Size & " pt")
  Else
    MsgBoxNew "Operation canceled"
  End If
End Sub


'set page setup
Private Sub Command3_Click()
  Call DlgInitMoveSystem(GetMoveMode(), txtShift.Text)
  Call ShowPageSetup(Me)
End Sub


'set printer
Private Sub Command4_Click()
  Call DlgInitMoveSystem(GetMoveMode(), txtShift.Text)
  Call ShowPrinter(Me)
End Sub


'open file
Private Sub Command5_Click()
  Dim sFile As String, sFilter As String
  sFilter = ""
  sFilter = AddFilterItem(sFilter, "VB projects (*.vbp)", "*.vbp")
  sFilter = AddFilterItem(sFilter, "VB files  (*.bas)", "*.bas")
  sFilter = AddFilterItem(sFilter, "All files (*.*)", "*.*")
  
  Call DlgInitMoveSystem(GetMoveMode(), txtShift.Text)
  sFile = ShowFileOpenSave(Me, True, "Open a VB file", LastFileName, sFilter, 1, ".txt")
  If sFile <> "" Then
    MsgBoxNew "You chose this file: " & sFile
    LastFileName = ExtractFileName(sFile)
    LastDir = ExtractPathName(sFile)
  Else
    MsgBoxNew "Operation canceled"
  End If
End Sub


'save as
Private Sub Command6_Click()
  Dim sFile As String, sFilter As String
  sFilter = ""
  sFilter = AddFilterItem(sFilter, "VB projects (*.vbp)", "*.vbp")
  sFilter = AddFilterItem(sFilter, "VB files  (*.bas)", "*.bas")
  sFilter = AddFilterItem(sFilter, "All files (*.*)", "*.*")
  
  Call DlgInitMoveSystem(GetMoveMode(), txtShift.Text)
  sFile = ShowFileOpenSave(Me, False, "Save as a VB file", LastFileName, sFilter, 1, ".txt")
  If sFile <> "" Then
    MsgBoxNew "You chose this file: " & sFile
    LastFileName = ExtractFileName(sFile)
    LastDir = ExtractPathName(sFile)
  Else
    MsgBoxNew "Operation canceled"
  End If
End Sub


'folder choice
Private Sub Command7_Click()
  Dim sPath As String
  Call DlgInitMoveSystem(GetMoveMode(), txtShift.Text)
  sPath = ShowFolder(Me, "Directories", LastDir)
  If sPath <> "" Then
    MsgBoxNew "You chose this folder: " & sPath
    LastDir = sPath
  Else
    MsgBoxNew "Operation canceled"
  End If
End Sub


'quit
Private Sub cmdQuit_Click()
  Unload Me
End Sub


