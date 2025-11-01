VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   Caption         =   "motor"
   ClientHeight    =   5235
   ClientLeft      =   2310
   ClientTop       =   2220
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7035
   Begin VB.CommandButton Command4 
      Caption         =   "¥ªÂà"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "¥kÂà"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "exit"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "submit"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim wBase As Integer
Dim wInitialCode As Integer

Private Sub Command1_Click()
    Dim port As Long
    port = &H378
    Dim data As Integer
   
    data = Text1.Text
    DIO_OutputByte port, data
End Sub

Private Sub Command2_Click()
    DIO_OutputByte &H378, 0
    DIO_DriverClose
    End
End Sub

'motor right rotarion

Private Sub Command3_Click()
    DIO_OutputByte &H378, 0
    Dim i, j, k, l As Integer
    For i = 1 To 300
        For j = 0 To 3
            k = 2 ^ j
            DIO_OutputByte &H378, k
            For k = 1 To 44444
            Next k
        Next j
    Next i

End Sub

'motor left rotation

Private Sub Command4_Click()
    DIO_OutputByte &H378, 0
    Dim i, j, k, l As Integer
    For i = 1 To 300
        For j = 3 To 0 Step -1
            k = 2 ^ j
            DIO_OutputByte &H378, k
            For k = 1 To 44444
            Next k
        Next j
    Next i

End Sub

Private Sub Form_Load()
    Dim rtn
    Dim port As Long
    Dim datainit As Integer
    port = &H378
    datainit = 0
       
    '********************************************************************
    '* NOTICE: call DIO_DriverInit() to initialize the driver.        *
    '* Initial the device driver, and return the board number in the PC *
    '********************************************************************
    wInitialCode = DIO_DriverInit()
    If wInitialCode <> DIO_NoError Then
        rtn = MsgBox("Driver Open Error !!!", , "DIO Card Error")
    End If
    
    DIO_OutputByte port, datainit

End Sub
