VERSION 5.00
Begin VB.Form frmSimulate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simulation"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   4695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   8295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Score 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   5280
      Width           =   75
   End
End
Attribute VB_Name = "frmSimulate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
Unload frmSimulate
frmRun.Show
End Sub

Private Sub Command1_Click()
Dim Toss As Integer
Text1.Text = ""
T1Score = 0
T2Score = 0
SeekCnt = 0
Text1.Text = Text1.Text + "Lee Jordan: Ladies and Gentlemen, welcome to the game between " + T1Name + " and " + T2Name + "." + vbNewLine
Toss = Int(2 * Rnd) + 1
If Toss = 1 Then
Text1.Text = Text1.Text + "Lee Jordan: Madam Hooch threw the Quaffle in the air! And " + T1Name + " Got it!" + vbNewLine
Call T1C1Pass
ElseIf Toss = 2 Then
Text1.Text = Text1.Text + "Lee Jordan: Madam Hooch threw the Quaffle in the air! And " + T2Name + " Got it!" + vbNewLine
Call T2C1Pass
End If
End Sub
Sub T1C1Pass()
Dim Interception As Integer, Rno As Integer
If t1cch = -1 Then
t1cch = Int(Rnd * 3)
End If
t1nch = Int(Rnd * 2)
If t1cch = 0 Then
If t1nch = 0 Then
t1nch = 1
Else
t1nch = 2
End If
ElseIf t1cch = 1 Then
If t1nch = 1 Then
t1nch = 2
Else
t1nch = 0
End If
End If
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " has the quaffle." + vbNewLine
Rno = Int(3 * Rnd) + 1
If Rno = 1 Then
Rno = Int(2 * Rnd) + 1
If Rno = 1 Then
Rno = Int(2 * Rnd)
Call T2BeaterT1(T1BSaving(Rno), 0, Rno)
Else
Call T2Beater(T1CReflexes(t1cch), 0)
End If
Else
If SeekCnt = 50 Then
Text1.Text = Text1.Text + "Lee Jordan: Is it a snitch? Yes, it is! Both seekers are rushing toward the snitch!" + vbNewLine
Call SnitchChase
Else
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " passes the quaffle to " + T1CName(t1nch) + "." + vbNewLine
SeekCnt = SeekCnt + 1
Interception = Int(2 * Rnd) + 1
If Interception = 1 Then
Call T2Intercept(T1CPassing(t1cch), 0)
Else
t1cch = t1nch
Call T1C2Pass
End If
End If
End If
End Sub
Sub T1C2Pass()
Dim Interception As Integer, Rno As Integer
If t1cch = -1 Then
t1cch = Int(Rnd * 3)
End If
t1nch = Int(Rnd * 2)
If t1cch = 0 Then
If t1nch = 0 Then
t1nch = 1
Else
t1nch = 2
End If
ElseIf t1cch = 1 Then
If t1nch = 1 Then
t1nch = 2
Else
t1nch = 0
End If
End If
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " has the quaffle." + vbNewLine
Rno = Int(3 * Rnd) + 1
If Rno = 1 Then
Rno = Int(2 * Rnd) + 1
If Rno = 1 Then
Rno = Int(2 * Rnd)
Call T2BeaterT1(T1BSaving(Rno), 1, Rno)
Else
Call T2Beater(T1CReflexes(t1cch), 1)
End If
Else
If SeekCnt = 50 Then
Text1.Text = Text1.Text + "Lee Jordan: Is it a snitch? Yes, it is! Both seekers are rushing toward the snitch!" + vbNewLine
Call SnitchChase
Else
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " passes the quaffle to " + T1CName(t1nch) + "." + vbNewLine
SeekCnt = SeekCnt + 1
Interception = Int(2 * Rnd) + 1
If Interception = 1 Then
Call T2Intercept(T1CPassing(t1cch), 1)
Else
t1cch = t1nch
Call T1C3Pass
End If
End If
End If
End Sub
Sub T1C3Pass()
Dim Interception As Integer, Rno As Integer
If t1cch = -1 Then
t1cch = Int(Rnd * 3)
End If
t1nch = Int(Rnd * 2)
If t1cch = 0 Then
If t1nch = 0 Then
t1nch = 1
Else
t1nch = 2
End If
ElseIf t1cch = 1 Then
If t1nch = 1 Then
t1nch = 2
Else
t1nch = 0
End If
End If
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " has the quaffle." + vbNewLine
Rno = Int(3 * Rnd) + 1
If Rno = 1 Then
Rno = Int(2 * Rnd) + 1
If Rno = 1 Then
Rno = Int(2 * Rnd)
Call T2BeaterT1(T1BSaving(Rno), 2, Rno)
Else
Call T2Beater(T1CReflexes(t1cch), 2)
End If
Else
If SeekCnt = 50 Then
Text1.Text = Text1.Text + "Lee Jordan: Is it a snitch? Yes, it is! Both seekers are rushing toward the snitch!" + vbNewLine
Call SnitchChase
Else
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " shoots!" + vbNewLine
SeekCnt = SeekCnt + 1
Cal = T1CShooting(t1cch) - T2KBlocking
If Cal < 0 Then
Cal = Cal * -1
If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If

If Cal >= 0 And Cal <= 5 And Rno = 1 Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

Else
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass
End If

Else
If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If

If Cal >= 0 And Cal <= 5 And Rno = 1 Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass

ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T2KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass


Else
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T1Score = T1Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t2cch = -1
Call T2C1Pass
End If

End If
End If
End If
End Sub


Sub T2C1Pass()
Dim Interception As Integer, Rno As Integer
If t2cch = -1 Then
t2cch = Int(Rnd * 3)
End If
t2nch = Int(Rnd * 2)
If t2cch = 0 Then
If t2nch = 0 Then
t2nch = 1
Else
t2nch = 2
End If
ElseIf t2cch = 1 Then
If t2nch = 1 Then
t2nch = 2
Else
t1nch = 0
End If
End If
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " has the quaffle." + vbNewLine
Rno = Int(3 * Rnd) + 1
If Rno = 1 Then
Rno = Int(2 * Rnd) + 1
If Rno = 1 Then
Rno = Int(2 * Rnd)
Call T1BeaterT2(T2BSaving(Rno), 0, Rno)
Else
Call T1Beater(T2CReflexes(t2cch), 0)
End If
Else
If SeekCnt = 50 Then
Text1.Text = Text1.Text + "Lee Jordan: Is it a snitch? Yes, it is! Both seekers are rushing toward the snitch!" + vbNewLine
Call SnitchChase
Else
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " passes the quaffle to " + T2CName(t2nch) + "." + vbNewLine
SeekCnt = SeekCnt + 1
Interception = Int(2 * Rnd) + 1
If Interception = 1 Then
Call T1Intercept(T2CPassing(t2cch), 0)
Else
t2cch = t2nch
Call T2C2Pass
End If
End If
End If
End Sub
Sub T2C2Pass()
Dim Interception As Integer, Rno As Integer
If t2cch = -1 Then
t2cch = Int(Rnd * 3)
End If
t2nch = Int(Rnd * 2)
If t2cch = 0 Then
If t2nch = 0 Then
t2nch = 1
Else
t2nch = 2
End If
ElseIf t2cch = 1 Then
If t2nch = 1 Then
t2nch = 2
Else
t1nch = 0
End If
End If
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " has the quaffle." + vbNewLine
Rno = Int(3 * Rnd) + 1
If Rno = 1 Then
Rno = Int(2 * Rnd) + 1
If Rno = 1 Then
Rno = Int(2 * Rnd)
Call T1BeaterT2(T2BSaving(Rno), 1, Rno)
Else
Call T1Beater(T2CReflexes(t2cch), 1)
End If
Else
If SeekCnt = 50 Then
Text1.Text = Text1.Text + "Lee Jordan: Is it a snitch? Yes, it is! Both seekers are rushing toward the snitch!" + vbNewLine
Call SnitchChase
Else
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " passes the quaffle to " + T2CName(t2nch) + "." + vbNewLine
SeekCnt = SeekCnt + 1
Interception = Int(2 * Rnd) + 1
If Interception = 1 Then
Call T1Intercept(T2CPassing(t2cch), 1)
Else
t2cch = t2nch
Call T2C3Pass
End If
End If
End If
End Sub
Sub T2C3Pass()
Dim Interception As Integer, Rno As Integer
If t2cch = -1 Then
t2cch = Int(Rnd * 3)
End If
t2nch = Int(Rnd * 2)
If t2cch = 0 Then
If t2nch = 0 Then
t2nch = 1
Else
t2nch = 2
End If
ElseIf t2cch = 1 Then
If t2nch = 1 Then
t2nch = 2
Else
t1nch = 0
End If
End If
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " has the quaffle." + vbNewLine
Rno = Int(3 * Rnd) + 1
If Rno = 1 Then
Rno = Int(2 * Rnd) + 1
If Rno = 1 Then
Rno = Int(2 * Rnd)
Call T1BeaterT2(T2BSaving(Rno), 2, Rno)
Else
Call T1Beater(T2CReflexes(t2cch), 2)
End If
Else
If SeekCnt = 50 Then
Text1.Text = Text1.Text + "Lee Jordan: Is it a snitch? Yes, it is! Both seekers are rushing toward the snitch!" + vbNewLine
Call SnitchChase
Else
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " shoots!" + vbNewLine
SeekCnt = SeekCnt + 1
Cal = T2CShooting(t2cch) - T1KBlocking
If Cal < 0 Then
Cal = Cal * -1
If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If

If Cal >= 0 And Cal <= 5 And Rno = 1 Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

Else
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass
End If

Else
If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If
If Cal >= 0 And Cal <= 5 And Rno = 1 Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass

ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
Text1.Text = Text1.Text + "Lee Jordan: The keeper " + T1KName + " saves it!" + vbNewLine
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass


Else
Text1.Text = Text1.Text + "Lee Jordan: Scores!" + vbNewLine
T2Score = T2Score + 10
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
t1cch = -1
Call T1C1Pass
End If

End If
End If
End If
End Sub

Sub T2Intercept(T1Pass As Integer, T1C As Integer)
Dim Cal As Integer, Rno As Integer
t2cch = Int(3 * Rnd)
Cal = T1Pass - T2CIntercepting(t2cch)
If Cal < 0 Then
Cal = Cal * -1

If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If

If Cal >= 0 And Cal <= 5 And Rno = 1 Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If

Else
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If
End If

Else
If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If

If Cal >= 0 And Cal <= 5 And Rno = 1 Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If


ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If


ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

Else
Text1.Text = Text1.Text + "Lee Jordan: " + T2CName(t2cch) + " of " + T2Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T1C = 0 Then
t1cch = t1nch
Call T1C2Pass
ElseIf T1C = 1 Then
t1cch = t1nch
Call T1C3Pass
End If
End If

End If

End Sub

Sub T1Intercept(T2Pass As Integer, T2C As Integer)
Dim Cal As Integer, Rno As Integer
t1cch = Int(3 * Rnd)
Cal = T2Pass - T1CIntercepting(t1cch)
If Cal < 0 Then
Cal = Cal * -1

If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If

If Cal >= 0 And Cal <= 5 And Rno = 1 Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If

Else
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If
End If

Else

If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If

If Cal >= 0 And Cal <= 5 And Rno = 1 Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " intercepts the quaffle!" + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If


Else
Text1.Text = Text1.Text + "Lee Jordan: " + T1CName(t1cch) + " of " + T1Name + " tries to intercept the quaffle, but fails!" + vbNewLine
If T2C = 0 Then
t2cch = t2nch
Call T2C2Pass
ElseIf T2C = 1 Then
t2cch = t2nch
Call T2C3Pass
End If
End If

End If

End Sub
Sub T2Beater(T1Reflexes As Integer, T1C As Integer)
Dim BNo As Integer, Cal As Integer, Rno As Integer
BNo = Int(2 * Rnd)
t2cch = Int(3 * Rnd)
Cal = T1Reflexes - T2BBeating(BNo)
If Cal < 0 Then
Cal = Cal * -1
If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If

If Cal >= 0 And Cal <= 5 And Rno = 1 Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

Else
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If
End If

Else
If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If

If Cal >= 0 And Cal <= 5 And Rno = 1 Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " and it hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

Else
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1CName(t1cch) + " dodges it!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If
End If

End If

End Sub
Sub T2BeaterT1(T1Reflexes As Integer, T1C As Integer, B As Integer)
Dim BNo As Integer, Cal As Integer, Rno As Integer
BNo = Int(2 * Rnd)
t2cch = Int(3 * Rnd)
Cal = T1Reflexes - T2BBeating(BNo)
If Cal < 0 Then
Cal = Cal * -1
If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If

If Cal >= 0 And Cal <= 5 And Rno = 1 Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If

Else
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If
End If

Else
If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If

If Cal >= 0 And Cal <= 5 And Rno = 1 Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + ". " + T1BName(B) + " tries for a save, but fails. The bludgers hits! " + T1CName(t1cch) + " of " + T1Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T2C1Pass
ElseIf Rno = 1 Then
Call T2C2Pass
ElseIf Rno = 2 Then
Call T2C3Pass
End If

Else
Text1.Text = Text1.Text + "Lee Jordan: " + T2BName(BNo) + " of " + T2Name + " whacks a bludger at " + T1CName(t1cch) + " but " + T1BName(B) + " whacks it away!" + vbNewLine
If T1C = 0 Then
Call T1C1Pass
ElseIf T1C = 1 Then
Call T1C2Pass
ElseIf T1C = 2 Then
Call T1C3Pass
End If
End If

End If

End Sub

Sub T1Beater(T2Reflexes As Integer, T2C As Integer)
Dim BNo As Integer, Cal As Integer, Rno As Integer
BNo = Int(2 * Rnd)
t1cch = Int(3 * Rnd)
Cal = T2Reflexes - T1BBeating(BNo)
If Cal < 0 Then
Cal = Cal * -1
If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If

If Cal >= 0 And Cal <= 5 And Rno = 1 Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

Else
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If
End If

Else
If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If

If Cal >= 0 And Cal <= 5 And Rno = 1 Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " and it hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

Else
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2CName(t2cch) + " dodges it!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If
End If

End If

End Sub
Sub T1BeaterT2(T2Reflexes As Integer, T2C As Integer, B As Integer)
Dim BNo As Integer, Cal As Integer, Rno As Integer
BNo = Int(2 * Rnd)
t1cch = Int(3 * Rnd)
Cal = T2Reflexes - T1BBeating(BNo)
If Cal < 0 Then
Cal = Cal * -1
If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If

If Cal >= 0 And Cal <= 5 And Rno = 1 Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If

Else
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If
End If

Else
If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If

If Cal >= 0 And Cal <= 5 And Rno = 1 Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + ". " + T2BName(B) + " tries for a save, but fails. The bludgers hits! " + T2CName(t2cch) + " of " + T2Name + " has lost the quaffle." + vbNewLine
Rno = Int(3 * Rnd)
If Rno = 0 Then
Call T1C1Pass
ElseIf Rno = 1 Then
Call T1C2Pass
ElseIf Rno = 2 Then
Call T1C3Pass
End If

Else
Text1.Text = Text1.Text + "Lee Jordan: " + T1BName(BNo) + " of " + T1Name + " whacks a bludger at " + T2CName(t2cch) + " but " + T2BName(B) + " whacks it away!" + vbNewLine
If T2C = 0 Then
Call T2C1Pass
ElseIf T2C = 1 Then
Call T2C2Pass
ElseIf T2C = 2 Then
Call T2C3Pass
End If
End If

End If

End Sub

Sub SnitchChase()
Dim Cal As Integer, Rno As Integer
Cal = T1SSeeking - T2SSeeking
If Cal < 0 Then
Cal = Cal * -1

If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If

If Cal >= 0 And Cal <= 5 And Rno = 1 Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

Else
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
End If

Else
If Cal = 1 Or Cal = 0 Then
Rno = Int(2 * Rnd) + 1
Else
Rno = Int(Cal * Rnd) + 1
End If

If Cal >= 0 And Cal <= 5 And Rno = 1 Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


ElseIf (Cal >= 6 And Cal <= 10) And (Rno = 1 Or Rno = 2) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


ElseIf (Cal >= 11 And Cal <= 15) And (Rno >= 1 And Rno <= 3) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


ElseIf (Cal >= 16 And Cal <= 20) And (Rno >= 1 And Rno <= 4) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


ElseIf (Cal >= 21 And Cal <= 25) And (Rno >= 1 And Rno <= 5) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


ElseIf (Cal >= 26 And Cal <= 30) And (Rno >= 1 And Rno <= 6) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


ElseIf (Cal >= 31 And Cal <= 35) And (Rno >= 1 And Rno <= 7) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


ElseIf (Cal >= 36 And Cal <= 40) And (Rno >= 1 And Rno <= 8) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


ElseIf (Cal >= 41 And Cal <= 45) And (Rno >= 1 And Rno <= 9) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


ElseIf (Cal >= 46 And Cal <= 50) And (Rno >= 1 And Rno <= 10) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


ElseIf (Cal >= 51 And Cal <= 55) And (Rno >= 1 And Rno <= 11) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


ElseIf (Cal >= 56 And Cal <= 60) And (Rno >= 1 And Rno <= 12) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


ElseIf (Cal >= 61 And Cal <= 65) And (Rno >= 1 And Rno <= 13) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


ElseIf (Cal >= 66 And Cal <= 70) And (Rno >= 1 And Rno <= 14) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


ElseIf (Cal >= 71 And Cal <= 75) And (Rno >= 1 And Rno <= 15) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


ElseIf (Cal >= 76 And Cal <= 80) And (Rno >= 1 And Rno <= 16) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


ElseIf (Cal >= 81 And Cal <= 85) And (Rno >= 1 And Rno <= 17) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 86 And Cal <= 90) And (Rno >= 1 And Rno <= 18) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


ElseIf (Cal >= 91 And Cal <= 95) And (Rno >= 1 And Rno <= 19) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine

ElseIf (Cal >= 96 And Cal <= 100) And (Rno >= 1 And Rno <= 20) Then
T2Score = T2Score + 150
If T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! " + T2Name + " win!" + vbNewLine
ElseIf T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! But " + T1Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T2SName + " of " + T2Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine


Else
T1Score = T1Score + 150
If T1Score > T2Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! " + T1Name + " win!" + vbNewLine
ElseIf T2Score > T1Score Then
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! But " + T2Name + " win! Because they got more points!" + vbNewLine
Else
Text1.Text = Text1.Text + "Lee Jordan: And " + T1SName + " of " + T1Name + " got the snitch! And tied the match!" + vbNewLine
End If
Text1.Text = Text1.Text + "Lee Jordan: Score is: " + T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score) + vbNewLine
End If

End If
Text1.Text = Text1.Text + "Lee Jordan: What a great game! See you next time. This is Lee Jordan, signing out!"
Score.Caption = T1Name + ": " + Str(T1Score) + " " + T2Name + ": " + Str(T2Score)
End Sub
Private Sub Form_Load()
Randomize Timer
t1cch = -1
t2cch = -1
End Sub
