VERSION 5.00
Begin VB.Form motor 
   BackColor       =   &H8000000E&
   Caption         =   "Stepping Motor"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   Icon            =   "motor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6990
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton ltun 
      Caption         =   "反轉"
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   5760
   End
   Begin VB.CommandButton exit 
      Caption         =   "離開"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton rtun 
      Caption         =   "正轉"
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton stop 
      Caption         =   "停止"
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   1560
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   5400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   0
      Top             =   5040
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   5520
      X2              =   6960
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   5520
      X2              =   6960
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Image Image5 
      Height          =   975
      Left            =   0
      Picture         =   "motor.frx":030A
      Top             =   5040
      Width           =   7020
   End
   Begin VB.Image Image4 
      Height          =   285
      Left            =   3120
      Picture         =   "motor.frx":1BEF
      Top             =   1800
      Width           =   180
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   3840
      Picture         =   "motor.frx":1F0B
      Top             =   1080
      Width           =   195
   End
   Begin VB.Image Image2 
      Height          =   270
      Left            =   3120
      Picture         =   "motor.frx":223D
      Top             =   360
      Width           =   195
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   3240
      X2              =   6960
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800080&
      BorderWidth     =   2
      X1              =   3960
      X2              =   6960
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   3240
      X2              =   6960
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Image Image1 
      Height          =   1500
      Index           =   0
      Left            =   5400
      Picture         =   "motor.frx":2553
      Top             =   4680
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Image Image1 
      Height          =   1500
      Index           =   1
      Left            =   5400
      Picture         =   "motor.frx":2B3D
      Top             =   4680
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image Image1 
      Height          =   1500
      Index           =   2
      Left            =   5400
      Picture         =   "motor.frx":311B
      Top             =   4680
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Image Image1 
      Height          =   1485
      Index           =   3
      Left            =   5400
      Picture         =   "motor.frx":3706
      Top             =   4680
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "PC控 步進馬達"
      BeginProperty Font 
         Name            =   "華康新儷粗黑"
         Size            =   26.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3360
      TabIndex        =   2
      Top             =   4080
      Width           =   3615
   End
End
Attribute VB_Name = "motor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function AVIO_OUT_LPT1 Lib "AVIO.dll" (ByVal PortData As Integer) As Integer
Dim i, X As Integer
Dim bExit As Boolean
Dim Index As Integer

Private Sub stop_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
bExit = True
End Sub

Private Sub exit_Click()
Unload motor
Unload welcom
End Sub

Private Sub Form_Resize()
motor.Height = 6345
motor.Width = 7080
End Sub

Private Sub rtun_Click()
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = False
bExit = False
While Not bExit
    AVIO_OUT_LPT1 (&H19)
    Call Timer1_Timer
DoEvents
    AVIO_OUT_LPT1 (&H39)
    Call Timer1_Timer
DoEvents
    AVIO_OUT_LPT1 (&H13)
    Call Timer1_Timer
DoEvents
    AVIO_OUT_LPT1 (&H33)
    Call Timer1_Timer
DoEvents
    AVIO_OUT_LPT1 (&H16)
    Call Timer1_Timer
DoEvents
    AVIO_OUT_LPT1 (&H36)
    Call Timer1_Timer
DoEvents
    AVIO_OUT_LPT1 (&H1C)
    Call Timer1_Timer
DoEvents
    AVIO_OUT_LPT1 (&H3C)
    Call Timer1_Timer
DoEvents
Wend
End Sub

Private Sub ltun_Click()
Timer3.Enabled = True
Timer2.Enabled = False
Timer1.Enabled = True
bExit = False
While Not bExit
    AVIO_OUT_LPT1 (&H19)
    Call Timer1_Timer
DoEvents
    AVIO_OUT_LPT1 (&H39)
    Call Timer1_Timer
DoEvents
    AVIO_OUT_LPT1 (&H1C)
    Call Timer1_Timer
DoEvents
    AVIO_OUT_LPT1 (&H3C)
    Call Timer1_Timer
DoEvents
    AVIO_OUT_LPT1 (&H16)
    Call Timer1_Timer
DoEvents
    AVIO_OUT_LPT1 (&H36)
    Call Timer1_Timer
DoEvents
    AVIO_OUT_LPT1 (&H13)
    Call Timer1_Timer
DoEvents
    AVIO_OUT_LPT1 (&H33)
    Call Timer1_Timer
DoEvents
Wend
End Sub

Private Sub Timer1_Timer()
For i = 1 To 65535
X = X * 2
X = X / 2
Next
End Sub

Private Sub Timer3_Timer()
Set motor.Picture = Image1(Index).Picture
Index = Index - 1
If Index < 0 Then Index = 3
End Sub

Private Sub Timer2_Timer()
Set motor.Picture = Image1(Index).Picture
Index = Index + 1
If Index > 3 Then Index = 0
End Sub
