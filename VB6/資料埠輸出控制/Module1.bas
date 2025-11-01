Attribute VB_Name = "Module1"
Declare Sub DIO_OutputByte Lib "dio.dll" (ByVal Address As Integer, ByVal DataOut As Integer)
Declare Function DIO_InputByte Lib "dio.dll" (ByVal Address As Integer) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long


Sub TimeDelay(TT As Long)
  Dim t As Long    '宣告一個長整數，記錄計數值
  t = GetTickCount()   '取得系統計數值
  Do   '開始迴圈
    DoEvents
    If GetTickCount - t < 0 Then t = GetTickCount '歸零
  Loop Until GetTickCount - t >= TT  '計算延遲是否到達
End Sub
