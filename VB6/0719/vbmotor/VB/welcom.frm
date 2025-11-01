VERSION 5.00
Begin VB.Form welcom 
   BackColor       =   &H8000000E&
   Caption         =   "welcome to AVIOSYS"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   Icon            =   "welcom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   8010
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton into 
      BackColor       =   &H00FFFFFF&
      Caption         =   "進入"
      BeginProperty Font 
         Name            =   "華康細黑體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   5
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label point 
      BackColor       =   &H8000000E&
      Caption         =   "●"
      Height          =   255
      Index           =   4
      Left            =   5400
      TabIndex        =   11
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label point 
      BackColor       =   &H8000000E&
      Caption         =   "●"
      Height          =   255
      Index           =   3
      Left            =   5760
      TabIndex        =   10
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label point 
      BackColor       =   &H8000000E&
      Caption         =   "●"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   9
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label point 
      BackColor       =   &H8000000E&
      Caption         =   "●"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   8
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label point2 
      BackColor       =   &H8000000E&
      Caption         =   "●"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label point 
      BackColor       =   &H8000000E&
      Caption         =   "●"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   6
      Top             =   3840
      Width           =   255
   End
   Begin VB.Line Line3 
      BorderStyle     =   3  '點線
      X1              =   1560
      X2              =   1560
      Y1              =   2160
      Y2              =   4200
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  '點線
      X1              =   1680
      X2              =   1680
      Y1              =   1080
      Y2              =   4200
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  '點線
      X1              =   3960
      X2              =   6240
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Image Image5 
      Appearance      =   0  '平面
      BorderStyle     =   1  '單線固定
      Height          =   1905
      Left            =   0
      Picture         =   "welcom.frx":030A
      Top             =   4200
      Width           =   2505
   End
   Begin VB.Image Image4 
      Appearance      =   0  '平面
      BorderStyle     =   1  '單線固定
      Height          =   1365
      Left            =   1920
      Picture         =   "welcom.frx":1E98
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Image Image3 
      Appearance      =   0  '平面
      BorderStyle     =   1  '單線固定
      Height          =   2310
      Left            =   2640
      Picture         =   "welcom.frx":2C3E
      Top             =   4560
      Width           =   1305
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   1680
      Picture         =   "welcom.frx":3D24
      Top             =   1080
      Width           =   5445
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   0
      Picture         =   "welcom.frx":BFF2
      Top             =   0
      Width           =   8565
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "歡迎使用睿意科技的步進馬達"
      BeginProperty Font 
         Name            =   "華康細黑體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   3240
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "如果在操作上有問題，請告訴我們"
      BeginProperty Font 
         Name            =   "華康細黑體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   3600
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "睿意科技 http://www.aviosys.com.tw"
      BeginProperty Font 
         Name            =   "華康細黑體(P)"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   3960
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "service@aviosys.com"
      BeginProperty Font 
         Name            =   "華康細黑體(P)"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "PC控  步進馬達"
      BeginProperty Font 
         Name            =   "華康新儷粗黑"
         Size            =   24
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   2640
      Width           =   3375
   End
End
Attribute VB_Name = "welcom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
welcom.Width = 8130
welcom.Height = 7455
End Sub

Private Sub into_Click()
motor.Show vbModal
End Sub

