VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmFrmfrm 
   Caption         =   "Handyman 2002"
   ClientHeight    =   8205
   ClientLeft      =   3000
   ClientTop       =   585
   ClientWidth     =   5940
   Icon            =   "frmFrmfrm.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmFrmfrm.frx":030A
   ScaleHeight     =   8205
   ScaleWidth      =   5940
   Begin VB.CommandButton Command15 
      Caption         =   "Rock-Paper-Scissors"
      Height          =   495
      Left            =   2880
      TabIndex        =   41
      Top             =   6960
      Width           =   3015
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Reverse it!"
      Height          =   495
      Left            =   0
      TabIndex        =   40
      Top             =   6960
      Width           =   2895
   End
   Begin VB.CommandButton Command13 
      Caption         =   "lower case converter"
      Height          =   735
      Left            =   4080
      TabIndex        =   39
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton Command12 
      Caption         =   "UPPER CASE CONVERTER"
      Height          =   735
      Left            =   2040
      TabIndex        =   38
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Fun with a mouth? Disgusting!"
      Height          =   735
      Left            =   0
      TabIndex        =   37
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Appointments"
      Height          =   735
      Left            =   4080
      TabIndex        =   36
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Text Slide Show"
      Height          =   735
      Left            =   2040
      TabIndex        =   35
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "The Ultimate Browser"
      Height          =   735
      Left            =   0
      TabIndex        =   34
      Top             =   6240
      Width           =   2055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   12
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Clock"
      TabPicture(0)   =   "frmFrmfrm.frx":0614
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(1)=   "Timer2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Calendar"
      TabPicture(1)   =   "frmFrmfrm.frx":0630
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Calendar1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Multiplication"
      TabPicture(2)   =   "frmFrmfrm.frx":064C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label9"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Text12"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Text13"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Command7"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Division"
      TabPicture(3)   =   "frmFrmfrm.frx":0668
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command4"
      Tab(3).Control(1)=   "Text7"
      Tab(3).Control(2)=   "Text6"
      Tab(3).Control(3)=   "Label6"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Addition"
      TabPicture(4)   =   "frmFrmfrm.frx":0684
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command5"
      Tab(4).Control(1)=   "Text9"
      Tab(4).Control(2)=   "Text8"
      Tab(4).Control(3)=   "Label7"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Subtraction"
      TabPicture(5)   =   "frmFrmfrm.frx":06A0
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Command6"
      Tab(5).Control(1)=   "Text11"
      Tab(5).Control(2)=   "Text10"
      Tab(5).Control(3)=   "Label8"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Square root"
      TabPicture(6)   =   "frmFrmfrm.frx":06BC
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Command1"
      Tab(6).Control(1)=   "Text1"
      Tab(6).Control(2)=   "Label2"
      Tab(6).Control(3)=   "Label1"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "Square"
      TabPicture(7)   =   "frmFrmfrm.frx":06D8
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Command2"
      Tab(7).Control(1)=   "Text2"
      Tab(7).Control(2)=   "Label4"
      Tab(7).Control(3)=   "Label3"
      Tab(7).ControlCount=   4
      TabCaption(8)   =   "Big Text"
      TabPicture(8)   =   "frmFrmfrm.frx":06F4
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Text3"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Small Text"
      TabPicture(9)   =   "frmFrmfrm.frx":0710
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Text4"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Game"
      TabPicture(10)  =   "frmFrmfrm.frx":072C
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Command3"
      Tab(10).Control(1)=   "Timer1"
      Tab(10).Control(2)=   "Text5"
      Tab(10).Control(3)=   "Label5"
      Tab(10).ControlCount=   4
      TabCaption(11)  =   "Circle Sizer"
      TabPicture(11)  =   "frmFrmfrm.frx":0748
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Picture1"
      Tab(11).Control(1)=   "HScroll1"
      Tab(11).ControlCount=   2
      Begin MSACAL.Calendar Calendar1 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   33
         Top             =   1320
         Width           =   5295
         _Version        =   524288
         _ExtentX        =   9340
         _ExtentY        =   8493
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2002
         Month           =   1
         Day             =   3
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   -74520
         Top             =   1680
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Calculate"
         Height          =   735
         Left            =   480
         TabIndex        =   30
         Top             =   2760
         Width           =   4455
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   480
         TabIndex        =   29
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   480
         TabIndex        =   28
         Top             =   1560
         Width           =   4455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Calculate"
         Height          =   615
         Left            =   -74400
         TabIndex        =   27
         Top             =   2760
         Width           =   3255
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Calculate"
         Height          =   975
         Left            =   -73920
         TabIndex        =   26
         Top             =   3000
         Width           =   3255
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   -74400
         TabIndex        =   25
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   -74400
         TabIndex        =   24
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   -74040
         TabIndex        =   22
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   -74040
         TabIndex        =   21
         Top             =   1560
         Width           =   3375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Calculate"
         Height          =   735
         Left            =   -74040
         TabIndex        =   18
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   -74040
         TabIndex        =   17
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   -74040
         TabIndex        =   16
         Top             =   1920
         Width           =   3375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Start!"
         Height          =   735
         Left            =   -73560
         TabIndex        =   15
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   -74880
         Top             =   3000
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   -74760
         TabIndex        =   14
         Top             =   2400
         Width           =   4935
      End
      Begin VB.PictureBox Picture1 
         Height          =   3015
         Left            =   -74760
         ScaleHeight     =   2955
         ScaleWidth      =   4635
         TabIndex        =   12
         Top             =   2640
         Width           =   4695
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   -74880
         Max             =   1000
         TabIndex        =   11
         Top             =   2040
         Width           =   4935
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   -74280
         TabIndex        =   10
         Top             =   1800
         Width           =   4215
      End
      Begin VB.TextBox Text3 
         Height          =   3735
         Left            =   -74400
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1680
         Width           =   4095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Calculate"
         Height          =   615
         Left            =   -74760
         TabIndex        =   7
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -74760
         TabIndex        =   5
         Top             =   1800
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Calculate"
         Height          =   615
         Left            =   -74760
         TabIndex        =   4
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -74760
         TabIndex        =   1
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -73680
         TabIndex        =   32
         Top             =   2520
         Width           =   3735
      End
      Begin VB.Label Label9 
         Caption         =   "The result will be here."
         Height          =   255
         Left            =   480
         TabIndex        =   31
         Top             =   3960
         Width           =   4455
      End
      Begin VB.Label Label8 
         Caption         =   "The result will be here."
         Height          =   255
         Left            =   -74400
         TabIndex        =   23
         Top             =   3600
         Width           =   3375
      End
      Begin VB.Label Label7 
         Caption         =   "The result will be here."
         Height          =   255
         Left            =   -74040
         TabIndex        =   20
         Top             =   4440
         Width           =   3375
      End
      Begin VB.Label Label6 
         Caption         =   "The result will be here."
         Height          =   255
         Left            =   -74040
         TabIndex        =   19
         Top             =   4200
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "How many characters can you type in 10 seconds? (Please do not copy and paste!)"
         Height          =   615
         Left            =   -74880
         TabIndex        =   13
         Top             =   1320
         Width           =   5175
      End
      Begin VB.Label Label4 
         Caption         =   "The result will be here."
         Height          =   255
         Left            =   -74760
         TabIndex        =   8
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Enter a number:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   6
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "The result will be here."
         Height          =   255
         Left            =   -74640
         TabIndex        =   3
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Enter a number."
         Height          =   255
         Left            =   -74760
         TabIndex        =   2
         Top             =   1680
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmFrmfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Jump As Integer 'Declare Jump As Integer
Private Sub Command1_Click()
On Error Resume Next 'Ignore Errors
Dim a As Double
Dim b As Double
a = Val(Text1.Text)
b = a ^ 0.5 'B is A's Square Root
Label2.Caption = "The square root of " & a & " is " & b & "."
End Sub


Private Sub Command10_Click()
frmAppoint.Show 'Show frmAppoint
End Sub

Private Sub Command11_Click()
frmFun.Show 'Show frmFun
End Sub

Private Sub Command12_Click()
frmUppercase.Show 'Show frmUppercase
End Sub

Private Sub Command13_Click()
frmLowercase.Show 'Show frmLowercase
End Sub

Private Sub Command14_Click()
frmReverse.Show 'Show frmReverse
End Sub

Private Sub Command15_Click()
frmPrepareRPS.Show 'Show frmPrepareRPS
End Sub



Private Sub Command2_Click()
On Error Resume Next 'Ignore Errors
Dim a As Double
Dim b As Double
a = Val(Text2.Text)
b = a * a 'B is A's square
Label4.Caption = "The square of " & a & " is " & b & "."
End Sub

Private Sub Command3_Click()
Timer1.Enabled = True 'Enable Timer1
End Sub

Private Sub Command4_Click()
On Error Resume Next 'Ignore Errors
Dim a As Double
Dim b As Double
Dim c As Double
a = Val(Text6.Text)
b = Val(Text7.Text)
c = a / b 'C is A divided by B
Label6.Caption = a & " divided by " & b & " is " & c & "."
End Sub

Private Sub Command5_Click()
On Error Resume Next 'Ignore Errors
Dim a As Double
Dim b As Double
Dim c As Double
a = Val(Text8.Text)
b = Text9.Text
c = a + b 'C is A plus B
Label7.Caption = a & " plus " & b & " is " & c & "."
End Sub

Private Sub Command6_Click()
On Error Resume Next 'Ignore Errors
Dim a As Double
Dim b As Double
Dim c As Double
a = Val(Text10.Text)
b = Val(Text11.Text)
c = a - b 'C is A minus B
Label8.Caption = a & " minus " & b & " is " & c & "."
End Sub

Private Sub Command7_Click()
Dim a As Double
Dim b As Double
Dim c As Double
a = Val(Text12.Text)
b = Val(Text13.Text)
c = a * b 'C is A times B
Label9.Caption = a & " times " & b & " is " & c & "."
End Sub

Private Sub Command8_Click()
frmBrowser.Show 'Show frmBrowser
End Sub

Private Sub Command9_Click()
frmSldsho.Show 'Show frmSldsho
End Sub

Private Sub Form_Load()
modHello.a
End Sub

Private Sub HScroll1_Change()
Picture1.Cls 'Clear PictureBox
Picture1.Circle (1700, 999), (HScroll1.Value) 'Draw Circle with size according to HScroll1.Value
End Sub

Private Sub Timer1_Timer()
Jump = Jump + 1 'Add 1 to Jump
If Jump = 10 Then
MsgBox "Time's up. You can type " & Len(Text5.Text) & " characters in 10 seconds."
Timer1.Enabled = False 'Disable Timer1
Jump = 0
Text5.Text = ""
End If
End Sub

Private Sub Timer2_Timer()
Label10.Caption = Time 'Change Label10.Caption to Time
End Sub
