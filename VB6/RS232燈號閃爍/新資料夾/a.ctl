VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl UserControl1 
   ClientHeight    =   4605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   ScaleHeight     =   4605
   ScaleWidth      =   6210
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2880
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3000
      Top             =   2040
   End
   Begin VB.CommandButton Command3 
      Caption         =   "結束"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RTS閃爍"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DTR閃爍"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Shape spRTS 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Left            =   4080
      Shape           =   3  '圓形
      Top             =   720
      Width           =   1215
   End
   Begin VB.Shape spDTR 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Left            =   840
      Shape           =   3  '圓形
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim fDTR As Boolean  'DTR閃爍旗標
Dim fRTS As Boolean  'RTS閃爍旗標

'DTR按鈕的程式
Private Sub cmdDTR_Click()
  fDTR = Not fDTR  '改變旗標狀態
  If fDTR Then '改變顯示的文字，方便使用者了解
    cmdDTR.Caption = "DTR停止"
  Else
    cmdDTR.Caption = "DTR閃爍"
  End If
End Sub

''''''''''''''''''''''''''''''''''''''
'使用命令按鈕控制項
'關閉通訊埠
'結束系統
''''''''''''''''''''''''''''''''''''''
Private Sub cmdEND_Click()
    MSComm1.DTREnable = False
    MSComm1.RTSEnable = False
    MSComm1.PortOpen = False
    End
End Sub

'RTS按鈕的程式
Private Sub cmdRTS_Click()
  fRTS = Not fRTS  '改變旗標狀態
  If fRTS Then '改變顯示的文字，方便使用者了解
    cmdRTS.Caption = "RTS停止"
  Else
    cmdRTS.Caption = "RTS閃爍"
  End If

End Sub

''''''''''''''''''''''''''''''''''''''
'表單的載入事件
'開啟通訊埠，通訊埠的參數也可以在開啟之前作
'設定，然後再開啟通訊埠
''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
  MSComm1.PortOpen = True
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'計時器控制項的Timer事件
'用以執行間隔指令
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Timer1_Timer()
    '若DTR的旗標為真時
    If fDTR Then
        MSComm1.DTREnable = Not MSComm1.DTREnable
        If MSComm1.DTREnable Then
          spDTR.FillColor = RGB(255, 0, 0)
        Else
          spDTR.FillColor = RGB(255, 255, 255)
        End If
    Else
      MSComm1.DTREnable = False
      spDTR.FillColor = RGB(255, 255, 255)
    End If
    '若RTS的旗標為真時
    If fRTS Then
        MSComm1.RTSEnable = Not MSComm1.RTSEnable
        If MSComm1.RTSEnable Then
          spRTS.FillColor = RGB(255, 0, 0)
        Else
          spRTS.FillColor = RGB(255, 255, 255)
        End If
    Else
       MSComm1.RTSEnable = False
       spRTS.FillColor = RGB(255, 255, 255)
    End If
End Sub

Private Sub UserControl_Initialize()
  MSComm1.PortOpen = True
End Sub
