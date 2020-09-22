VERSION 5.00
Begin VB.Form frmSldsho 
   Caption         =   "Handy Slide"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   2535
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmSldsho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmShow.Show 'Show frmShow
End Sub

