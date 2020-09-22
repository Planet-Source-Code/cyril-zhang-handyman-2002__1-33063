VERSION 5.00
Begin VB.Form frmPrepareRPS 
   Caption         =   "Prepare For Rock-Paper-Scissors Battle"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Text            =   "?????"
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "?????"
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go!"
      Enabled         =   0   'False
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Player 2:"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Player 1:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "frmPrepareRPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmBattle.Show 'Show frmBattle
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Or Text1.Text <> Text2.Text Then
Command1.Enabled = True 'Enable Command1
Else
Command1.Enabled = False 'Disable Command1
End If
End Sub

Private Sub Text2_Change()
If Text2.Text <> "" Or Text2.Text <> Text1.Text Then
Command1.Enabled = True 'Enable Command1
Else
Command1.Enabled = False 'Disable Command1
End If
End Sub
