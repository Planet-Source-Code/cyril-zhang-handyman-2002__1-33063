VERSION 5.00
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "VText.dll"
Begin VB.Form frmFun 
   Caption         =   "Handyman Mouth Fun. EWWW!"
   ClientHeight    =   5790
   ClientLeft      =   2685
   ClientTop       =   300
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   6090
   Begin VB.CommandButton Command7 
      Caption         =   "Talk"
      Height          =   1095
      Left            =   1080
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Red Lips"
      Height          =   1215
      Left            =   2160
      TabIndex        =   6
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Smiling Scream"
      Height          =   1095
      Left            =   5040
      TabIndex        =   5
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Pink Lips"
      Height          =   1095
      Left            =   5040
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Neutral"
      Height          =   1095
      Left            =   3960
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Smile"
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Kiss"
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin HTTSLibCtl.TextToSpeech mouth 
      Height          =   2775
      Left            =   1680
      OleObjectBlob   =   "frmFun.frx":0000
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   2280
      TabIndex        =   8
      Text            =   "Speak This:"
      Top             =   2520
      Width           =   1575
   End
End
Attribute VB_Name = "frmFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
mouth.LipTension = 255
End Sub

Private Sub Command2_Click()
mouth.LipTension = 0
End Sub

Private Sub Command3_Click()
mouth.MouthUpturn = 255
End Sub

Private Sub Command4_Click()
mouth.MouthUpturn = 0
End Sub

Private Sub Command5_Click()
mouth.LipType = 1
End Sub

Private Sub Command6_Click()
mouth.LipType = 0
End Sub

Private Sub Command7_Click()
mouth.Speak (Text1.Text)
End Sub

Private Sub Command8_Click()
mouth.MouthUpturn = 255
mouth.LipTension = 0
End Sub

