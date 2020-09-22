VERSION 5.00
Begin VB.Form frmLowercase 
   Caption         =   "lower case converter"
   ClientHeight    =   3195
   ClientLeft      =   870
   ClientTop       =   1185
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "convert"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label3 
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Result:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Data:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmLowercase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label3.Caption = LCase(Text1.Text) 'Change to LCase
End Sub
