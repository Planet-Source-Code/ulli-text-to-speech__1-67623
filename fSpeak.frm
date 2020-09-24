VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fSpeak 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Speak"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7755
   Icon            =   "fSpeak.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDrop 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'Kein
      Height          =   645
      Left            =   6975
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   645
      ScaleWidth      =   720
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   90
      Width           =   720
      Begin VB.Shape shpDrop 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         FillColor       =   &H008080FF&
         FillStyle       =   7  'Diagonalkreuz
         Height          =   645
         Left            =   0
         Shape           =   3  'Kreis
         Top             =   0
         Width           =   645
      End
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   6975
      Top             =   4785
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open Textfile"
   End
   Begin VB.CommandButton btShutup 
      Caption         =   "&Shut Up"
      Height          =   435
      Left            =   4222
      TabIndex        =   3
      Top             =   4770
      Width           =   930
   End
   Begin VB.CommandButton btSpeak 
      Height          =   435
      Left            =   2602
      TabIndex        =   2
      Top             =   4770
      Width           =   930
   End
   Begin VB.TextBox txText 
      Height          =   3660
      Left            =   202
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   900
      Width           =   7350
   End
   Begin VB.Timer tmrToggle 
      Interval        =   500
      Left            =   6315
      Top             =   4785
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Drop .TXT file here:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5175
      TabIndex        =   0
      Top             =   315
      Width           =   1710
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSpeak 
         Caption         =   "Speak"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuShutup 
         Caption         =   "ShutUp"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "fSpeak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_TOPMOST   As Long = -1
Private Const SWP_NOSIZE    As Long = 1
Private Const SWP_NOMOVE    As Long = 2

Private MyWidth             As Long
Private HeightDiff          As Long
Private Speaking            As Boolean
Private Paused              As Boolean
Private WithEvents Voice    As SpVoice
Attribute Voice.VB_VarHelpID = -1

Private Sub btShutup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Speaking And Not Paused Then
        Paused = True
        Voice.Pause
        btSpeak.Caption = "&Resume"
    End If

End Sub

Private Sub btShutup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    txText.SetFocus

End Sub

Private Sub btSpeak_Click()

    Paused = False
    If Speaking Then
        Voice.Resume
        txText.SetFocus
      Else 'SPEAKING = FALSE/0
        Set Voice = New SpVoice
        Voice_EndStream 0, 0
        If txText = vbNullString Then
            shpDrop.Visible = False
            DoEvents
            Voice.Speak "nothing to read; you must open a text file, or drop one on the red landing pad.", SVSFPurgeBeforeSpeak
            SendKeys "%F", True
            SetWindowPos hWnd, SWP_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
          Else 'NOT TXTEXT...
            Speaking = True
            Voice.Speak txText, SVSFlagsAsync
            txText.SetFocus
        End If
    End If

End Sub

Private Sub Form_Load()

    Voice_EndStream 0, 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Speaking Then
        Voice.Pause
        Set Voice = Nothing
        Speaking = False
    End If

End Sub

Private Function GetText(hFile As Long) As String

    picDrop.Visible = False
    lb.Visible = False
    GetText = Input$(LOF(hFile), hFile)
    Close hFile
    Voice_EndStream 0, 0

End Function

Private Sub mnuClose_Click()

    btShutup_MouseDown 0, 0, 0, 0
    picDrop.Visible = True
    lb.Visible = True
    txText = vbNullString
    Form_Unload 0
    Voice_EndStream 0, 0

End Sub

Private Sub mnuExit_Click()

    Unload Me

End Sub

Private Sub mnuOpen_Click()

  Dim hFile     As Long

    mnuClose_Click
    With cDlg
        .Flags = cDlg.Flags Or _
                 cdlOFNExplorer Or _
                 cdlOFNHelpButton Or _
                 cdlOFNLongNames Or _
                 cdlOFNFileMustExist
        .Filter = "Text Files (*.txt)|*.txt"
        On Error Resume Next
            .ShowOpen
            hFile = Err
        On Error GoTo 0
        If hFile = 0 Then
            hFile = FreeFile
            Open .FileName For Binary As hFile
            txText = GetText(hFile)
        End If
    End With 'CDLG

End Sub

Private Sub mnuShutup_Click()

    btShutup_MouseDown 0, 0, 0, 0

End Sub

Private Sub mnuSpeak_Click()

    btSpeak_Click

End Sub

Private Sub picDrop_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim hFile     As Long

    hFile = FreeFile
    On Error Resume Next
        Open Data.Files(1) For Binary As hFile
        If Err Then
            MsgBox "Files only, please", vbExclamation, "Oops..."
          ElseIf LCase$(Right$(Data.Files(1), 4)) <> ".txt" Then 'ERR = FALSE/0
            MsgBox "Only .TXT files, please", vbExclamation, "Oops..."
          Else 'NOT LCASE$(RIGHT$(DATA.FILES(1),...
            txText = GetText(hFile)
            btSpeak_Click
        End If
    On Error GoTo 0

End Sub

Private Sub tmrToggle_Timer()

    shpDrop.Visible = Not shpDrop.Visible
    SetWindowPos hWnd, SWP_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

Private Sub txText_GotFocus()

    HideCaret txText.hWnd

End Sub

Private Sub txText_KeyPress(KeyAscii As Integer)

    KeyAscii = 0

End Sub

Private Sub Voice_EndStream(ByVal StreamNumber As Long, ByVal StreamPosition As Variant)

    btSpeak.Caption = "&Read this"
    Speaking = False
    On Error Resume Next
        btSpeak.SetFocus
        txText.SelLength = 0
    On Error GoTo 0

End Sub

Private Sub Voice_Word(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal CharacterPosition As Long, ByVal Length As Long)

    With txText
        .SetFocus
        .SelStart = CharacterPosition
        .SelLength = Length
    End With 'TXTEXT

End Sub

':) Ulli's VB Code Formatter V2.21.7 (2007-Jan-14 18:42)  Decl: 13  Code: 182  Total: 195 Lines
':) CommentOnly: 0 (0%)  Commented: 6 (3,1%)  Empty: 55 (28,2%)  Max Logic Depth: 3
