VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FF80&
   Caption         =   "motor"
   ClientHeight    =   2760
   ClientLeft      =   2310
   ClientTop       =   2220
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   5280
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim rtn
    Dim winitialcode As Integer
    Port = &H378
    data_init = 0
    DIO_OutputByte Port, data_init
  
    '********************************************************************
    '* NOTICE: call DIO_DriverInit() to initialize the driver.        *
    '* Initial the device driver, and return the board number in the PC *
    '********************************************************************
    winitialcode = DIO_DriverInit()
    If winitialcode <> DIO_NoError Then
        rtn = MsgBox("Driver Open Error !!!", , "DIO Card Error")
    End If
   
End Sub

