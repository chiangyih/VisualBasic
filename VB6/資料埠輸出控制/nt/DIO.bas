Attribute VB_Name = "DIO"

Global Const DIO_NoError = 0
Global Const DIO_DriverOpenError = 1
Global Const DIO_DriverNoOpen = 2
Global Const DIO_GetDriverVersionError = 3
Global Const DIO_InstallIrqError = 4
Global Const DIO_ClearIntCountError = 5
Global Const DIO_GetIntCountError = 6
Global Const DIO_ResetError = 7
Global Const DIO_RemoveIrqError = 8


' The test functions
Declare Function DIO_ShortSub2 Lib "dio.dll" _
        (ByVal a As Integer, ByVal b As Integer) As Integer
Declare Function DIO_FloatSub2 Lib "dio.dll" _
        (ByVal a As Single, ByVal b As Single) As Single


' The DIO functions
Declare Sub DIO_OutputByte Lib "dio.dll" _
        (ByVal address As Integer, ByVal dataout As Byte)
Declare Sub DIO_OutputWord Lib "dio.dll" _
        (ByVal address As Integer, ByVal dataout As Integer)
Declare Function DIO_InputByte Lib "dio.dll" _
        (ByVal address As Integer) As Integer
Declare Function DIO_InputWord Lib "dio.dll" _
        (ByVal address As Integer) As Integer

' The Driver functions
Declare Function DIO_DriverInit Lib "dio.dll" () As Integer
Declare Sub DIO_DriverClose Lib "dio.dll" ()
Declare Function DIO_GetDllVersion Lib "dio.dll" () As Integer
Declare Function DIO_GetDriverVersion Lib "dio.dll" _
        (wDriverVersion As Integer) As Integer

' The Interrupt functions
Declare Function DIO_InstallIrq Lib "dio.dll" _
        (ByVal wBase As Integer, ByVal wIrq As Integer, _
        hEvent As Long) As Integer
Declare Function DIO_RemoveIrq Lib "dio.dll" _
        (ByVal hEvent As Long) As Integer
Declare Function DIO_GetIntCount Lib "dio.dll" _
        (dwVal As Integer) As Integer


' Declare Function DIO_Reset Lib "dio.dll" () As Integer
