VERSION 5.00
Begin VB.Form frmSmile 
   Caption         =   "Smilies"
   ClientHeight    =   8040
   ClientLeft      =   2265
   ClientTop       =   345
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   8355
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.ListBox List2 
      Height          =   3960
      Left            =   2760
      TabIndex        =   1
      Top             =   2160
      Width           =   5175
   End
   Begin VB.ListBox List1 
      Height          =   3960
      ItemData        =   "frmSmile.frx":0000
      Left            =   360
      List            =   "frmSmile.frx":0002
      TabIndex        =   0
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Description"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   1800
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Smiley"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
End
Attribute VB_Name = "frmSmile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
