Imports System.Diagnostics

' CPU 使用率監控程式
Public Class Form1
    Private cpuCounter As PerformanceCounter '效能計數器物件
    Private WithEvents updateTimer As New Timer() '計時器物件

    ' 表單載入時初始化
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 建立 CPU 效能計數器 (_Total 代表所有核心的總和)
        cpuCounter = New PerformanceCounter("Processor", "% Processor Time", "_Total") 'new一個效能計數器物件,Processor類別,計數器名稱為% Processor Time,實例名稱為_Total
        cpuCounter.NextValue() '第一次呼叫初始化，通常回傳 0
        updateTimer.Interval = 800 '每 0.8 秒更新一次
    End Sub

    ' 啟動監控
    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        btnStart.Enabled = False
        updateTimer.Start()
    End Sub

    ' 每次計時器觸發時更新 CPU 使用率
    Private Sub updateTimer_Tick(sender As Object, e As EventArgs) Handles updateTimer.Tick
        Dim cpuUsage As Integer = CInt(Math.Round(cpuCounter.NextValue())) '定義一個整數變數cpuUsage,將效能計數器的下一個值取整數後賦值給cpuUsage
        lblCpu.Text = $"CPU: {cpuUsage} %" '更新標籤顯示 CPU 使用率
    End Sub

    ' 關閉程式
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    ' 表單關閉時釋放資源
    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        MyBase.OnFormClosing(e)
        updateTimer.Stop() '停止計時器
        cpuCounter?.Dispose() '釋放效能計數器資源
    End Sub
End Class
