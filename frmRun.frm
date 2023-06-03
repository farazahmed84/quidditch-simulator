VERSION 5.00
Begin VB.Form frmRun 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quidditch Simulator"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   2520
      TabIndex        =   22
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton btnLineup2 
      Caption         =   "View Lineup"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton btnLineup1 
      Caption         =   "View Lineup"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "Simulate"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox cboT2 
      Height          =   315
      ItemData        =   "frmRun.frx":0000
      Left            =   2760
      List            =   "frmRun.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox cboT1 
      Height          =   315
      ItemData        =   "frmRun.frx":0004
      Left            =   360
      List            =   "frmRun.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   2880
      TabIndex        =   21
      Top             =   4680
      Width           =   45
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   2880
      TabIndex        =   20
      Top             =   4320
      Width           =   45
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   2880
      TabIndex        =   19
      Top             =   3960
      Width           =   45
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   2880
      TabIndex        =   18
      Top             =   3600
      Width           =   45
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   2880
      TabIndex        =   17
      Top             =   3240
      Width           =   45
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   2880
      TabIndex        =   16
      Top             =   2880
      Width           =   45
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   2880
      TabIndex        =   15
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblSelTeam2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Select 2nd Team:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblSelTeam1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Select 1st Team:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblVs 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Vs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Width           =   375
   End
End
Attribute VB_Name = "frmRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnLineup1_Click()
Dim T As String, c1 As String, c1p As Integer, c1s As Integer, c1i As Integer, c1r As Integer
Dim c2 As String, c2p As Integer, c2s As Integer, c2i As Integer, c2r As Integer
Dim c3 As String, c3p As Integer, c3s As Integer, c3i As Integer, c3r As Integer
Dim k As String, kb As Integer
Dim s As String, ss As Integer
Dim b1 As String, bb1 As Integer, bs1 As Integer
Dim b2 As String, bb2 As Integer, bs2 As Integer
Open "teams.txt" For Input As #1
Do Until EOF(1)
Input #1, T
Input #1, c1, c1p, c1s, cli, clr
Input #1, c2, c2p, c2s, c2i, c2r
Input #1, c3, c3p, c3s, c3i, c3r
Input #1, k, kb
Input #1, s, ss
Input #1, b1, bb1, bs1
Input #1, b2, bb2, bs2
If cboT1.Text = T Then
Label1.Caption = "Chaser 1:"
Label2.Caption = "Chaser 2:"
Label3.Caption = "Chaser 3:"
Label4.Caption = "Keeper:"
Label5.Caption = "Seeker:"
Label6.Caption = "Beater 1:"
Label7.Caption = "Beater 2:"

Label8.Caption = c1
Label9.Caption = c2
Label10.Caption = c3
Label11.Caption = k
Label12.Caption = s
Label13.Caption = b1
Label14.Caption = b2
End If
Loop
Close 1

End Sub

Private Sub btnLineup2_Click()
Dim T As String, c1 As String, c1p As Integer, c1s As Integer, c1i As Integer, c1r As Integer
Dim c2 As String, c2p As Integer, c2s As Integer, c2i As Integer, c2r As Integer
Dim c3 As String, c3p As Integer, c3s As Integer, c3i As Integer, c3r As Integer
Dim k As String, kb As Integer
Dim s As String, ss As Integer
Dim b1 As String, bb1 As Integer, bs1 As Integer
Dim b2 As String, bb2 As Integer, bs2 As Integer
Open "teams.txt" For Input As #1
Do Until EOF(1)
Input #1, T
Input #1, c1, c1p, c1s, cli, clr
Input #1, c2, c2p, c2s, c2i, c2r
Input #1, c3, c3p, c3s, c3i, c3r
Input #1, k, kb
Input #1, s, ss
Input #1, b1, bb1, bs1
Input #1, b2, bb2, bs2
If cboT2.Text = T Then
Label1.Caption = "Chaser 1:"
Label2.Caption = "Chaser 2:"
Label3.Caption = "Chaser 3:"
Label4.Caption = "Keeper:"
Label5.Caption = "Seeker:"
Label6.Caption = "Beater 1:"
Label7.Caption = "Beater 2:"

Label8.Caption = c1
Label9.Caption = c2
Label10.Caption = c3
Label11.Caption = k
Label12.Caption = s
Label13.Caption = b1
Label14.Caption = b2
End If
Loop
Close 1
End Sub

Private Sub btnPlay_Click()
Dim T1 As String, T2 As String
T1Name = cboT1.Text
T2Name = cboT2.Text
If cboT1.Text = "" Or cboT2.Text = "" Then
MsgBox ("Please select both teams.")

ElseIf T1Name = T2Name Then
MsgBox ("Both teams should be different.")

Else
Dim T As String, c1 As String, c1p As Integer, c1s As Integer, c1i As Integer, c1r As Integer
Dim c2 As String, c2p As Integer, c2s As Integer, c2i As Integer, c2r As Integer
Dim c3 As String, c3p As Integer, c3s As Integer, c3i As Integer, c3r As Integer
Dim k As String, kb As Integer
Dim s As String, ss As Integer
Dim b1 As String, bb1 As Integer
Dim b2 As String, bb2 As Integer
Open "teams.txt" For Input As #1
Do Until EOF(1)
Input #1, T
Input #1, c1, c1p, c1s, cli, clr
Input #1, c2, c2p, c2s, c2i, c2r
Input #1, c3, c3p, c3s, c3i, c3r
Input #1, k, kb
Input #1, s, ss
Input #1, b1, bb1, bs1
Input #1, b2, bb2, bs2
If T1Name = T Then
T1CName(0) = c1
T1CName(1) = c2
T1CName(2) = c3
T1CPassing(0) = c1p
T1CPassing(1) = c2p
T1CPassing(2) = c3p
T1CShooting(0) = c1s
T1CShooting(1) = c2s
T1CShooting(2) = c3s
T1CIntercepting(0) = c1i
T1CIntercepting(1) = c2i
T1CIntercepting(2) = c3i
T1CReflexes(0) = c1r
T1CReflexes(1) = c2r
T1CReflexes(2) = c3r
T1KName = k
T1KBlocking = kb
T1SName = s
T1SSeeking = ss
T1BName(0) = b1
T1BBeating(0) = bb1
T1BSaving(0) = bs1
T1BName(1) = b2
T1BBeating(1) = bb2
T1BSaving(1) = bs2
ElseIf T2Name = T Then
T2CName(0) = c1
T2CName(1) = c2
T2CName(2) = c3
T2CPassing(0) = c1p
T2CPassing(1) = c2p
T2CPassing(2) = c3p
T2CShooting(0) = c1s
T2CShooting(1) = c2s
T2CShooting(2) = c3s
T2CIntercepting(0) = c1i
T2CIntercepting(1) = c2i
T2CIntercepting(2) = c3i
T2CReflexes(0) = c1r
T2CReflexes(1) = c2r
T2CReflexes(2) = c3r
T2KName = k
T2KBlocking = kb
T2SName = s
T2SSeeking = ss
T2BName(0) = b1
T2BBeating(0) = bb1
T2BSaving(0) = bs1
T2BName(1) = b2
T2BBeating(1) = bb2
T2BSaving(1) = bs2
End If
Loop
Close 1
Unload frmRun
frmSimulate.Show
End If
End Sub

Private Sub cmdBack_Click()
Unload frmRun
frmMain.Show
End Sub

Private Sub Form_Load()
Dim T As String, c1 As String, c1p As Integer, c1s As Integer, c1i As Integer, c1r As Integer
Dim c2 As String, c2p As Integer, c2s As Integer, c2i As Integer, c2r As Integer
Dim c3 As String, c3p As Integer, c3s As Integer, c3i As Integer, c3r As Integer
Dim k As String, kb As Integer
Dim s As String, ss As Integer
Dim b1 As String, bb1 As Integer, bs1 As Integer
Dim b2 As String, bb2 As Integer, bs2 As Integer
Open "teams.txt" For Input As #1
Do Until EOF(1)
Input #1, T
Input #1, c1, c1p, c1s, cli, clr
Input #1, c2, c2p, c2s, c2i, c2r
Input #1, c3, c3p, c3s, c3i, c3r
Input #1, k, kb
Input #1, s, ss
Input #1, b1, bb1, bs1
Input #1, b2, bb2, bs2
cboT1.AddItem T
cboT2.AddItem T
Loop
Close 1
End Sub


