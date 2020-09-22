VERSION 5.00
Begin VB.Form id3Frm 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ID3 Editor"
   ClientHeight    =   2430
   ClientLeft      =   4230
   ClientTop       =   3075
   ClientWidth     =   2550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmb 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   960
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   4
      Left            =   0
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   3
      Left            =   0
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   2
      Left            =   0
      MaxLength       =   30
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   0
      MaxLength       =   30
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   0
      MaxLength       =   30
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Okay"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "id3Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FLoaded As String 'The File That Was In Play When This Was Loaded
Private Sub Form_Load()
    Me.BackColor = mFrm.BackColor 'Set This Forms Backcolor To The Main Forms Color (Just To Make It Look Nice)
    mFrm.Enabled = False 'Disable The Main Form
    FLoaded = FileLoaded 'Set The Local File Name
    Dim ct As Control
    For Each ct In Me.Controls
        ct.BackColor = Me.BackColor 'Go Through Each Control And Set Its Color To The Backcolor
    Next
    Dim x As Long
    For x = 0 To 125
        cmb.AddItem Style(x) 'Add All The Styles To The Combobox
    Next
    cmb.AddItem "Unknown", 0
    GetTag FLoaded, Info 'Get The Tag Of The Current File
    txt(1).Text = Trim(Replace(Info.Album, Chr(0), "")) 'Set The Textboxes To The Info
    txt(2).Text = Trim(Replace(Info.Artist, Chr(0), ""))
    txt(4).Text = Trim(Replace(Info.Comments, Chr(0), ""))
    txt(3).Text = Trim(Replace(Info.Date, Chr(0), ""))
    txt(0).Text = Trim(Replace(Info.Name, Chr(0), ""))
    For x = 0 To 126 'Look For The Correct Style
        If Style(Val(Info.Style)) = cmb.List(x) Then
            cmb.ListIndex = x 'Set The Combobox To The Style
            Exit For
        End If
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mFrm.Enabled = True 'Enable The Main Form
End Sub

Private Sub Label1_Click()
    Info.Album = txt(1).Text
    Info.Artist = txt(2).Text
    Info.Comments = txt(4).Text
    Info.Date = txt(3).Text
    Info.Name = txt(0).Text
    Info.Style = StyleId(cmb.Text)
    'We Set All The Id3 Info To Be Saved
    If SaveTag(FLoaded, Info) = False Then MsgBox "Unable To Save ID3 Tag", vbCritical + vbMsgBoxSetForeground, "Error"
    'We Try To Save It, If We're Unable To We Say So
    Unload Me
End Sub

Private Sub Label2_Click()
    Unload Me 'Cancel Was Hit, Unload
End Sub
