VERSION 5.00
Begin VB.Form frmUppercase 
   Caption         =   "UPPER CASE CONVERTER"
   ClientHeight    =   3720
   ClientLeft      =   6705
   ClientTop       =   1140
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   4680
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CONVERT"
      Height          =   975
      Left            =   1080
      TabIndex        =   0
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label3 
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Result:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Data."
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmUppercase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label3.Caption = UCase(Text1.Text) 'Change to UCase
End Sub
