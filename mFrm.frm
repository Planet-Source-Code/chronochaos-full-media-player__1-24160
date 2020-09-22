VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form mFrm 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Chrono Chaos' Creations - Media Player"
   ClientHeight    =   2430
   ClientLeft      =   3030
   ClientTop       =   2760
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   162
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   279
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox files 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1410
      Left            =   120
      OLEDropMode     =   1  'Manual
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   3975
   End
   Begin VB.PictureBox pos 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   213
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   600
      Width           =   3255
   End
   Begin VB.Timer Scroller 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open File..."
      Filter          =   "Media Files|*.wav;*.mp3;*.avi;*.mid;*.midi;*.mpeg;*.mpg"
   End
   Begin VB.TextBox Status 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Buttons 
      BackStyle       =   0  'Transparent
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   27.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   6
      Left            =   3600
      TabIndex        =   9
      ToolTipText     =   "ID3 Info"
      Top             =   0
      Width           =   615
   End
   Begin VB.Label lTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Buttons 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   3120
      TabIndex        =   6
      ToolTipText     =   "Open File..."
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Buttons 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   2520
      TabIndex        =   5
      ToolTipText     =   "Forward"
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Buttons 
      BackStyle       =   0  'Transparent
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   4
      ToolTipText     =   "Stop"
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Buttons 
      BackStyle       =   0  'Transparent
      Caption         =   ";"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "Pause"
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Buttons 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   2
      ToolTipText     =   "Play"
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Buttons 
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Back"
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "mFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Got To Tell VB Everything's Declared                                                                                                                                                                                                                                                                                                                                                                                                          _
                                                                                                                                                                                                                                                                                         _
-----------------------------------------                                                                                                                                                                                                                                                         _
       Chrono Chaos' Creations                                                                                                                                                                                                                                         _
         www.chronochaos.com                                                                                                                                                                                                                                                                                                                                    _
         mail@chronochaos.com                                                                                                                                                                                                                                                                                                     _
-----------------------------------------

'This Is Just A Simple Media Player Example Feel Free To LEARN From My Code
'Any Questions Email Me. My Address Is Right Above
'This Code Is Free And "As-Is" Without Any Warranty Or Anything Like That
'If You Use This Code In Any Project Email Me With A Screen Shot
'Hope This Helps You Out
'Everythings Commented, If Its Not Then Use Your Common Sense (ex. 1 + 3 = 4 No Comment)
'Also Everything Has Be Broken Down Into Sections (For You People Who Copy And Paste All The Time)

'This Was Tested And Works Fine On My Windows 2000 System
'If This Doesn't Work Correctly On Your OS Then Please Email Me

Dim sScroll As String 'Holds What We Are Scrolling In The Status Box

Private Sub Buttons_Click(Index As Integer)
    'Instead Of Having 6 Different Subs We'll Just Use An Array
    On Error GoTo ErrHand:
    Select Case Index
        Case 0
            'Go Back
            If Val(files.Tag) = 0 Then
                files.ListIndex = files.ListCount - 1 'Set The Listbox's Index So That We Can Use The "Text" Property Correctly
            Else
                files.ListIndex = Val(files.Tag) - 1
            End If
            If files.ListCount = 0 Then GoTo ErrHand
            files.Tag = files.ListIndex 'Set The Tag To The Currently Playing Song
            OpenMedia files.Text
            MediaPlay
        Case 1
            MediaPlay
        Case 2
            MediaPause
        Case 3
            MediaStop
        Case 4
            'Go Forward
            If Val(files.Tag) = files.ListCount - 1 Then
                files.ListIndex = 0
            Else
                files.ListIndex = Val(files.Tag) + 1 'Set The Listbox's Index So That We Can Use The "Text" Property Correctly
            End If
            files.Tag = files.ListIndex 'Set The Tag To The Currently Playing Song
            OpenMedia files.Text
            MediaPlay
        Case 5
            pos.Cls 'Clear The Position Status Bar
            NewFile
        Case 6
            If FileExt(FileLoaded) = "mp3" Then id3Frm.Show
    End Select
ErrHand:
End Sub
Function NewFile() As Boolean
    On Error GoTo ErrHand
    cDialog.ShowOpen 'Show The Open Dialog
    Dim s As String
    s = FileExt(cDialog.FileName) 'Get The File Extension
    Select Case s 'We Select The File's Extension
        'I Prefer Selecting The Extension Because It Allows Me To
        'Display Different Error Messages
    Case "mp3", "wav", "midi", "mid", "avi", "mpg", "mpeg"
        files.AddItem cDialog.FileName
    Case "xxx"
        MsgBox "Custom Error Message Can Go Here"
    Case Else
        MsgBox "Unknown File Format. Please Make Sure The Format Is Valid And Try Again.", vbCritical + vbMsgBoxSetForeground, "Error"
        GoTo ErrHand
End Select
NewFile = True
Exit Function
ErrHand:
NewFile = False 'On Error Return False
End Function
Sub LoadStatus(Ext As String) 'Loads The Text That Will Be Scrolled
    If Ext = "mp3" Then 'Check To See What The Extension Is
    GetTag FileLoaded, Info 'Get The Id3 Tag
    sScroll = Trim(Replace(Info.Artist, Chr(0), "")) & " - " & Trim(Replace(Info.Name, Chr(0), "")) & " - "  'Set The sScroller's Text To The Artist And The Songname
    If Replace(Info.Artist, Chr(0), "") = "" And Replace(Info.Name, Chr(0), "") = "" Then GoTo noid3: 'Check To See If There's A Valid ID3
Else
noid3:
    sScroll = Trim(FileLoaded) & " - " 'If Its Not An Mp3 We Set sScroll To Be The Filename
End If
sScroll = sScroll & sScroll & sScroll & sScroll
'Why Four Times? It Just Makes It Scroll Smoother
'If You Got A Better/Simpler Way To Do This Then DO IT!
End Sub

Private Sub files_DblClick()
    pos.Cls 'Clear The Status Bar
    OpenMedia files.Text 'Load It
    MediaPlay 'Play It
    files.Tag = files.ListIndex 'Save The Listindex To Control The Front And Back Commands
End Sub

Private Sub files_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    On Error GoTo ErrH:
    Do
        i = i + 1 'Next File
        Select Case FileExt(Data.files(i)) 'Select The Current File's Ext
            Case "mp3", "wav", "midi", "mid", "avi", "mpg", "mpeg" 'If Its One Of Our File Types..
                files.AddItem Data.files(i) '..Add It To The List
        End Select
        Loop
ErrH:
End Sub

Private Sub Form_Load()
    Randomize
    Me.BackColor = RGB(Int((Rnd * 255) + 1), Int((Rnd * 255) + 1), Int((Rnd * 255) + 1))
    Status.BackColor = Me.BackColor
    lTime.BackColor = Me.BackColor
    pos.BackColor = Me.BackColor
    files.BackColor = Me.BackColor
    'Just Not To Get Old Too Fast Let's Make The Backcolor Change On Every Load
End Sub

Function Mci(sCommand As String) As String
    'Sends A Command To The Windows Media Dll
    Dim s As String * 255 'Create A String With A Pre-Defined Buffer Of 255 Char
    Call mciSendString(sCommand, s, 255, Me.hWnd) 'Call The Mci Send String Api Call
    Mci = Replace(s, Chr(0), "") 'Return Only What Is Needed
End Function

Function OpenMedia(fName As String) As Boolean
    'Opens The Media File
    'Returns True If We Successfully Opened The File
    MediaReset 'Close Everything
    OpenMedia = CBool(Val(Mci("open """ & fName & """ alias med")))
    Mci "set med time format milliseconds"
    FileLoaded = fName 'Set The File That Was Loaded
    LoadStatus FileExt(fName)
End Function

Private Sub Form_Unload(Cancel As Integer)
    MediaReset
End Sub
Function MediaReset()
    Mci "close all" 'Close Every File Thats Open
End Function
Function MediaPlay()
    Mci "play med"
End Function
Function MediaStop()
    Mci "stop med"
End Function
Function MediaPause()
    Mci "pause med"
End Function
Function MediaFullScreen()
    Mci "play med fullscreen"
End Function
Function MediaPlayRepeat()
    Mci "play med repeat"
End Function
Function MediaLength() As Long
    MediaLength = Val(Mci("status med length"))
End Function

Private Sub pos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Mci "seek med to " & Int(MediaLength / pos.Width * x) 'Seek To..
    MediaPlay
End Sub

Private Sub pos_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    pos.Cls 'We Clear It...
    '...So We Have To Draw The Position Line Again
    On Error Resume Next 'To Avoid The "Division By Zero Error"
    i = pos.Width / MediaLength * MediaPosition
    pos.Line (i, 0)-(i, pos.Height) 'We Make A Little Line
    FloodFill pos.hdc, 2, 2, vbWhite 'This Won't Actually Fill It With White..I Think Im Forgeting Something..
    pos.Line (x, 0)-(x, pos.Height), vbRed 'We Make Another Line For The Position To Jump To
    Status.Text = "Jump To " & SecsToMin(Int(MediaLength / pos.Width * x), True)
End Sub

Private Sub Scroller_Timer()
    Dim x As Long 'This Will Store The Current Scroll Position
    Dim i As Long 'This Will Store The To Be Drawn Line
    x = Val(Scroller.Tag) 'Get The Scroll Position
    x = x + 1 'Move The Scrolled Text Over 1 Char
    pos.Cls 'Clear The Position Picturebox
    Status = Mid(sScroll, x) & Left(sScroll, x) 'Set Status To The sScroll's Text
    If x >= Len(sScroll) Then x = 0 'If It Equals sScrolls Length Go Back To 0
    Scroller.Tag = x 'Set The Timers Tag To The Current Position
    lTime = SecsToMin(MediaPosition, True) 'Set The Current Time Of The Media File
    On Error Resume Next 'To Avoid The "Division By Zero Error"
    i = pos.Width / MediaLength * MediaPosition
    pos.Line (i, 0)-(i, pos.Height), vbWhite 'We Make A Little Line
    FloodFill pos.hdc, 2, 2, vbWhite
End Sub
Function MediaPosition() As Long
    MediaPosition = Val(Mci("status med position"))
End Function
Function MediaStatus() As Long
    Select Case Mci("status med mode")
        Case "playing"
            MediaStatus = 1
        Case "paused"
            MediaStatus = 2
    End Select 'If Its Stopped It Will Return Zero (0)
End Function
