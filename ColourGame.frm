VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   2700
   ClientTop       =   1650
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   563
   Begin VB.PictureBox menu 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   2280
      ScaleHeight     =   3105
      ScaleWidth      =   3585
      TabIndex        =   6
      Top             =   2160
      Width           =   3615
      Begin VB.Label HighScore 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Top Score"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label About 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label ExitG 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Label NewGame 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "New Game"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   3375
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3840
      Top             =   3120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Rounds: 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2280
      TabIndex        =   1
      Top             =   3240
      UseMnemonic     =   0   'False
      Width           =   3615
   End
   Begin VB.Label Command 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Width           =   4095
   End
   Begin VB.Label Command 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Index           =   3
      Left            =   4080
      TabIndex        =   4
      Top             =   3720
      Width           =   4095
   End
   Begin VB.Label Command 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Index           =   2
      Left            =   4080
      TabIndex        =   3
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Command 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3495
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Colour Game v1.0.0
'Author: Billy Gunnels
'July, 2003
'Thanks to Serrano, Optillusions.com, planetsourcecode.com
'
'This is just a little game I made to fuck with your brain
'If you want to improve the game (or even compleate its high score table)
'I would love to see what you do with it so please repost it
'Have Fun :)

Option Explicit

Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function CreateEllipticRgn& Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const RGN_AND = 1
Dim mMove As Boolean
Dim AnsN As Single
Dim Answered As Boolean
Dim FileName As String
Dim rounds As Single

Private Sub About_Click()
MsgBox "ColourGame v1.0.0" + vbCrLf + _
        "By Billy Gunnels" + vbCrLf + _
        "July, 2003" + vbCrLf + vbCrLf + _
        "Thanks To:" + vbCrLf + _
        """Planetsourcecode.com""" + vbCrLf + _
        """David Serrano"" - Form Shape" + vbCrLf + _
        """http://optillusions.com/dp/1-41.htm"" - Concept" + vbCrLf + vbCrLf + _
        "If you like the game please" + vbCrLf + _
        "improve upon it and repost! :)"
        
End Sub

Private Sub Command_Click(Index As Integer)
If Timer1.Enabled = False Then Exit Sub
If Index + 1 <> AnsN Then
    MsgBox "WRONG"
    menu.Visible = True
    Timer1.Enabled = False
    Exit Sub
End If
box.Caption = ""
rounds = rounds + 1
Label1.Caption = "Rounds: " + Str(rounds)
Answered = True
End Sub

Private Sub Command_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If Timer1.Enabled = False Then
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
End If
End Sub

Private Sub ExitG_Click()
Unload Me
End
End Sub

Private Sub Form_Load()
SetWindowRgn hWnd, CreateEllipticRgn(50, 50, 500, 500), True
Me.Show
End Sub

Private Sub NewGame_Click()
menu.Visible = False
Answered = True
rounds = 0
Label1.Caption = "Rounds: 0"
Timer1.Interval = 2000
box.Caption = ""
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim RndToken As Single
Dim x As Long

If Answered = False Then
    MsgBox "TOO SLOW"
    menu.Visible = True
    Timer1.Enabled = False
    Exit Sub
End If

'=======================================================
'Long ass code to randomize the position of the colors
'It can be done with a lot less code im just a little lazy
'to redo it all.
Randomize
RndToken = Int(4 * Rnd) + 1
If RndToken = 1 Then Command(0).Left = 0: Command(0).Top = 16: Command(0).Tag = "1"
If RndToken = 2 Then Command(0).Left = 272: Command(0).Top = 16: Command(0).Tag = "2"
If RndToken = 3 Then Command(0).Left = 0: Command(0).Top = 248: Command(0).Tag = "3"
If RndToken = 4 Then Command(0).Left = 272: Command(0).Top = 248: Command(0).Tag = "4"
'--
roll1:
RndToken = Int(4 * Rnd) + 1
If RndToken = 1 Then
    If Command(0).Tag <> "1" Then
        Command(1).Left = 0: Command(1).Top = 16: Command(1).Tag = "1"
    Else
        GoTo roll1
    End If
End If
If RndToken = 2 Then
    If Command(0).Tag <> "2" Then
        Command(1).Left = 272: Command(1).Top = 16: Command(1).Tag = "2"
    Else
        GoTo roll1
    End If
End If
If RndToken = 3 Then
    If Command(0).Tag <> "3" Then
        Command(1).Left = 0: Command(1).Top = 248: Command(1).Tag = "3"
    Else
        GoTo roll1
    End If
End If
If RndToken = 4 Then
    If Command(0).Tag <> "4" Then
        Command(1).Left = 272: Command(1).Top = 248: Command(1).Tag = "4"
    Else
        GoTo roll1
    End If
End If
'--
Roll2:
RndToken = Int(4 * Rnd) + 1
If RndToken = 1 Then
    If Command(0).Tag <> "1" And Command(1).Tag <> "1" Then
        Command(2).Left = 0: Command(2).Top = 16: Command(2).Tag = "1"
    Else
        GoTo Roll2
    End If
End If
If RndToken = 2 Then
    If Command(0).Tag <> "2" And Command(1).Tag <> "2" Then
        Command(2).Left = 272: Command(2).Top = 16: Command(2).Tag = "2"
    Else
        GoTo Roll2
    End If
End If
If RndToken = 3 Then
    If Command(0).Tag <> "3" And Command(1).Tag <> "3" Then
        Command(2).Left = 0: Command(2).Top = 248: Command(2).Tag = "3"
    Else
        GoTo Roll2
    End If
End If
If RndToken = 4 Then
    If Command(0).Tag <> "4" And Command(1).Tag <> "4" Then
        Command(2).Left = 272: Command(2).Top = 248: Command(2).Tag = "4"
    Else
        GoTo Roll2
    End If
End If
'--
Roll3:
RndToken = Int(4 * Rnd) + 1
If RndToken = 1 Then
    If Command(0).Tag <> "1" And Command(1).Tag <> "1" And Command(2).Tag <> "1" Then
        Command(3).Left = 0: Command(3).Top = 16: Command(3).Tag = "1"
    Else
        GoTo Roll3
    End If
End If
If RndToken = 2 Then
    If Command(0).Tag <> "2" And Command(1).Tag <> "2" And Command(2).Tag <> "2" Then
        Command(3).Left = 272: Command(3).Top = 16: Command(3).Tag = "2"
    Else
        GoTo Roll3
    End If
End If
If RndToken = 3 Then
    If Command(0).Tag <> "3" And Command(1).Tag <> "3" And Command(2).Tag <> "3" Then
        Command(3).Left = 0: Command(3).Top = 248: Command(3).Tag = "3"
    Else
        GoTo Roll3
    End If
End If
If RndToken = 4 Then
    If Command(0).Tag <> "4" And Command(1).Tag <> "4" And Command(2).Tag <> "4" Then
        Command(3).Left = 272: Command(3).Top = 248: Command(3).Tag = "4"
    Else
        GoTo Roll3
    End If
End If
'=========================================================

RndToken = Int(4 * Rnd)
If RndToken = 1 Then x = sndPlaySound(App.Path + "\red.wav", &H1)
If RndToken = 2 Then x = sndPlaySound(App.Path + "\blue.wav", &H1)
If RndToken = 3 Then x = sndPlaySound(App.Path + "\green.wav", &H1)
If RndToken = 4 Then x = sndPlaySound(App.Path + "\yellow.wav", &H1)
' /\Play random wave /\
' \/Random text color\/
RndToken = Int(4 * Rnd) + 1
If RndToken = 1 Then box.ForeColor = vbRed
If RndToken = 2 Then box.ForeColor = vbBlue
If RndToken = 3 Then box.ForeColor = vbGreen
If RndToken = 4 Then box.ForeColor = vbYellow

AnsN = Int(4 * Rnd) + 1
If AnsN = 1 Then box.Caption = "RED"
If AnsN = 2 Then box.Caption = "BLUE"
If AnsN = 3 Then box.Caption = "GREEN"
If AnsN = 4 Then box.Caption = "YELLOW"

Answered = False
If Timer1.Interval > 900 Then Timer1.Interval = Timer1.Interval - 100

End Sub

Public Sub SaveListBox(TheList As ListBox, Directory As String)
Dim SaveList As Long
    
On Error Resume Next
Open Directory$ For Output As #1


For SaveList& = 0 To TheList.ListCount - 1
    Print #1, TheList.List(SaveList&)
Next SaveList&
Close #1
End Sub

Public Sub LoadListBox(TheList As ListBox, Directory As String)
Dim MyString As String

On Error Resume Next
Open Directory$ For Input As #1

While Not EOF(1)
    Input #1, MyString$
    DoEvents
        TheList.AddItem MyString$
    Wend
    Close #1
        
End Sub
