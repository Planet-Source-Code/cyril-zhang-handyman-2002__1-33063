VERSION 5.00
Begin VB.Form frmBattle 
   Caption         =   "Rock-Paper-Scissors Battle"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Ready"
      Height          =   1575
      Left            =   3360
      TabIndex        =   9
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ready"
      Height          =   1575
      Left            =   480
      TabIndex        =   8
      Top             =   2880
      Width           =   1455
   End
   Begin VB.PictureBox b2 
      Height          =   975
      Left            =   3000
      ScaleHeight     =   915
      ScaleWidth      =   2355
      TabIndex        =   7
      Top             =   1200
      Width           =   2415
   End
   Begin VB.PictureBox b1 
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmBattle.frx":0000
      Left            =   3240
      List            =   "frmBattle.frx":000D
      TabIndex        =   5
      Text            =   "Rock"
      Top             =   1440
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmBattle.frx":0028
      Left            =   120
      List            =   "frmBattle.frx":0035
      TabIndex        =   4
      Text            =   "Rock"
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label l2 
      Caption         =   "Do not look!"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label l1 
      Caption         =   "Do not look!"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmBattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
b2.Visible = False 'b2 is invisible
b1.Visible = True 'b1 is visible
l1.Visible = True 'l1 is visible
l2.Visible = False 'l2 is invisible
End Sub

Private Sub Command2_Click()
Dim MSG As String
MSG = frmPrepareRPS.Text1.Text & " used " & Combo1.Text & "." & vbCrLf
MSG = MSG + frmPrepareRPS.Text2.Text & " used " & Combo2.Text & "." & vbCrLf & vbCrLf
If Combo1.Text = "Rock" And Combo2.Text = "Paper" Then
MSG = MSG + frmPrepareRPS.Text2.Text & " won."
ElseIf Combo1.Text = "Rock" And Combo2.Text = "Scissors" Then
MSG = MSG + frmPrepareRPS.Text1.Text & " won."
ElseIf Combo1.Text = "Paper" And Combo2.Text = "Scissors" Then
MSG = MSG + frmPrepareRPS.Text2.Text & " won."
ElseIf Combo1.Text = "Paper" And Combo2.Text = "Rock" Then
MSG = MSG + frmPrepareRPS.Text1.Text & " won."
ElseIf Combo1.Text = "Scissors" And Combo2.Text = "Rock" Then
MSG = MSG + frmPrepareRPS.Text2.Text & " won."
ElseIf Combo1.Text = "Scissors" And Combo2.Text = "Paper" Then
MSG = MSG + frmPrepareRPS.Text1.Text & " won."
Else
MSG = MSG + "It was a draw."
End If
'That was a big chunk of code!
MsgBox MSG, vbInformation + vbOKOnly, "Who has won?" 'Choice message
b2.Visible = True 'b2 is visible
b1.Visible = False 'b1 is invisible
l1.Visible = False 'l1 is invisible
l2.Visible = True 'l2 is visible
End Sub

Private Sub Form_Load()
Label1.Caption = frmPrepareRPS.Text1.Text & ":" 'Player 1
Label2.Caption = frmPrepareRPS.Text2.Text & ":" 'Player 2
End Sub
