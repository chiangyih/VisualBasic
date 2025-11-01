VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   8400
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "輸出"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4050
      TabIndex        =   12
      Top             =   1170
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2790
      TabIndex        =   11
      Text            =   "0"
      Top             =   1215
      Width           =   960
   End
   Begin VB.CommandButton Command2 
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
      Height          =   510
      Left            =   6840
      TabIndex        =   9
      Top             =   1125
      Width           =   1185
   End
   Begin VB.CheckBox DataOut 
      Caption         =   "D7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   7
      Left            =   6030
      TabIndex        =   7
      Top             =   405
      Width           =   780
   End
   Begin VB.CheckBox DataOut 
      Caption         =   "D6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   6
      Left            =   5265
      TabIndex        =   6
      Top             =   405
      Width           =   780
   End
   Begin VB.CheckBox DataOut 
      Caption         =   "D5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   5
      Left            =   4500
      TabIndex        =   5
      Top             =   405
      Width           =   780
   End
   Begin VB.CheckBox DataOut 
      Caption         =   "D4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   4
      Left            =   3735
      TabIndex        =   4
      Top             =   405
      Width           =   780
   End
   Begin VB.CheckBox DataOut 
      Caption         =   "D3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   3
      Left            =   2970
      TabIndex        =   3
      Top             =   405
      Width           =   780
   End
   Begin VB.CheckBox DataOut 
      Caption         =   "D2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   2205
      TabIndex        =   2
      Top             =   405
      Width           =   780
   End
   Begin VB.CheckBox DataOut 
      Caption         =   "D1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   405
      Width           =   780
   End
   Begin VB.CheckBox DataOut 
      Caption         =   "D0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   675
      TabIndex        =   0
      Top             =   405
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "資料輸出"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   1710
      TabIndex        =   10
      Top             =   1260
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "資料輸出"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   2925
      TabIndex        =   8
      Top             =   90
      Width           =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PtrAddress As Long          '印表機埠變數
Dim DataOutValue As Integer  '資料埠變數


Private Sub Command1_Click()
      DIO_OutputByte PtrAddress, Val(Text1.Text)  '將文字框內數值送到指定的印表機位址上
End Sub

Private Sub Command2_Click()
  DIO_OutputByte PtrAddress, 0    '印表機埠資料線清空
  End
End Sub

Private Sub DataOut_Click(Index As Integer)
    If DataOut(Index).Value = 1 Then '該資料核取框是否被核取?
      DataOutValue = DataOutValue Or 2 ^ Index '若核取則將該位置的輸出變數值設定為1
    Else
      DataOutValue = DataOutValue And (255 - 2 ^ Index) '否則設定為0
    End If
    DIO_OutputByte PtrAddress, DataOutValue '再將變數值送到指定的印表機位址上
End Sub

Private Sub Form_Load()
    Dim code As Integer
    Dim rtn
    code = DIO_DriverInit()
    If code <> DIO_NoError Then
        rtn = MsgBox("driver open error!!!", , "dio error")
    End If
        
    PtrAddress = &H378   '指定印表機埠的位址
    DataOutValue = 0 '將資料埠變數值預設為0
    CtrlOutValue = 0 '將控制埠變數值預設為0
    DIO_OutputByte PtrAddress, DataOutValue '將此值設定給指定的印表機位址
End Sub

