VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   6810
   StartUpPosition =   3  '系統預設值
   Begin MSCommLib.MSComm MSComm1 
      Left            =   180
      Top             =   1035
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "結束系統"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   3750
      TabIndex        =   5
      Top             =   1590
      Width           =   1620
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "啟動偵測"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   1140
      TabIndex        =   4
      Top             =   1590
      Width           =   1620
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   210
      Top             =   465
   End
   Begin VB.Label lblMsg 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   90
      TabIndex        =   6
      Top             =   2775
      Width           =   6735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "RI"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DSR"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   3800
      TabIndex        =   2
      Top             =   135
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CTS"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2470
      TabIndex        =   1
      Top             =   120
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DCD"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1065
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape spRI 
      FillStyle       =   0  '實心
      Height          =   645
      Left            =   4980
      Shape           =   3  '圓形
      Top             =   465
      Width           =   675
   End
   Begin VB.Shape spDSR 
      FillStyle       =   0  '實心
      Height          =   645
      Left            =   3655
      Shape           =   3  '圓形
      Top             =   465
      Width           =   675
   End
   Begin VB.Shape spCTS 
      FillStyle       =   0  '實心
      Height          =   645
      Left            =   2330
      Shape           =   3  '圓形
      Top             =   465
      Width           =   675
   End
   Begin VB.Shape spDCD 
      FillStyle       =   0  '實心
      Height          =   645
      Left            =   1005
      Shape           =   3  '圓形
      Top             =   465
      Width           =   675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''
'使用命令按鈕控制項
'關閉通訊埠
'結束系統
''''''''''''''''''''''''''''''''''''''
Private Sub cmdEnd_Click()
MSComm1.PortOpen = False
End
End Sub

''''''''''''''''''''''''''''''''''''''
'命令按鈕控制項的滑鼠移動事件
'提示一個訊息給使用者了解
'將訊息顯示指示到Label控制項
''''''''''''''''''''''''''''''''''''''
Private Sub cmdEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMsg.Caption = "按下此鍵可結束本系統!"

End Sub

''''''''''''''''''''''''''''''''''''''
'命令按鈕控制項
'用以將計時器的啟動/關閉作一個切換
''''''''''''''''''''''''''''''''''''''
Private Sub cmdStart_Click()
Timer1.Enabled = Not Timer1.Enabled

End Sub

''''''''''''''''''''''''''''''''''''''
'命令按鈕控制項的滑鼠移動事件
'提示一個訊息給使用者了解
'將訊息顯示指示到Label控制項
''''''''''''''''''''''''''''''''''''''
Private Sub cmdStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMsg.Caption = "按下此鍵便可開始作偵測的工作!"
End Sub

''''''''''''''''''''''''''''''''''''''
'表單載入事件
'開啟串列通訊埠
''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
MSComm1.PortOpen = True
End Sub

''''''''''''''''''''''''''''''''''''''
'計時器控制項的Timer事件
'每隔一個固定時間會啟動一次事件程序內的程式碼
'在此則是檢查電壓狀態，並回應到Shape的填滿顏色
''''''''''''''''''''''''''''''''''''''
Private Sub Timer1_Timer()
If MSComm1.CDHolding Then               '偵測DCD腳位電位
  spDCD.FillColor = RGB(255, 0, 0)      '改變燈號為紅燈
Else
  spDCD.FillColor = RGB(0, 0, 0)        '燈號滅掉
End If
If MSComm1.CTSHolding Then              '偵測CTS腳位電位
  spCTS.FillColor = RGB(255, 0, 0)      '改變燈號為紅燈
Else
  spCTS.FillColor = RGB(0, 0, 0)        '燈號滅掉
End If
If MSComm1.DSRHolding Then              '偵測DSR腳位電位
  spDSR.FillColor = RGB(255, 0, 0)      '改變燈號為紅燈
Else
  spDSR.FillColor = RGB(0, 0, 0)        '燈號滅掉
End If
If MSComm1.CommEvent = comEvRing Then
  spRI.FillColor = RGB(255, 0, 0)      '改變燈號為紅燈
Else
  spRI.FillColor = RGB(0, 0, 0)        '燈號滅掉
End If
End Sub

