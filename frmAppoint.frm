VERSION 5.00
Begin VB.Form frmAppoint 
   Caption         =   "Handy Appointment Maker"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   1560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make"
      Height          =   735
      Left            =   840
      TabIndex        =   4
      Top             =   1560
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmAppoint.frx":0000
      Left            =   2160
      List            =   "frmAppoint.frx":000A
      TabIndex        =   3
      Text            =   "AM or PM"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   ":"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   ":"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   720
      Width           =   135
   End
End
Attribute VB_Name = "frmAppoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
If Time = Text1.Text & ":" & Text2.Text & ":" & Text3.Text & " " & Combo1.Text Then 'Ready Time
Timer2.Enabled = True ' Enable Timer2
End If
End Sub

Private Sub Timer2_Timer()
Beep 'Beep
frmAppoint.BackColor = vbRed 'Change Form's BackColor to vbRed
'Big Change!
End Sub
