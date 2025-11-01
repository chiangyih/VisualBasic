Imports System.Diagnostics

Public Class Form1
    ' 記憶體效能計數器
    Private memoryCounter As PerformanceCounter
    ' 監控狀態旗標
    Private isMonitoring As Boolean = False

    ''' <summary>
    ''' 表單載入時初始化記憶體計數器
    ''' </summary>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' 建立記憶體使用率計數器（與工作管理員一致）
            memoryCounter = New PerformanceCounter("Memory", "% Committed Bytes In Use")
            memoryCounter.NextValue() ' 初始化
        Catch ex As Exception
            MessageBox.Show("無法初始化記憶體計數器: " & ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 啟動/停止監控按鈕
    ''' </summary>
    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        ' 切換監控狀態
        isMonitoring = Not isMonitoring

        If isMonitoring Then
            ' 啟動監控
            tmrUpdate.Start()
            btnStart.Text = "停止監控"
        Else
            ' 停止監控
            tmrUpdate.Stop()
            btnStart.Text = "啟動監控"
        End If
    End Sub

    ''' <summary>
    ''' 結束程式按鈕
    ''' </summary>
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Application.Exit()
    End Sub

    ''' <summary>
    ''' 每秒更新記憶體使用率
    ''' </summary>
    Private Sub tmrUpdate_Tick(sender As Object, e As EventArgs) Handles tmrUpdate.Tick
        Try
            ' 讀取並顯示記憶體使用率（一位小數）
            Dim usage As Single = memoryCounter.NextValue()
            lblMemoryUsage.Text = usage.ToString("F1") & " %"
        Catch ex As Exception
            ' 發生錯誤時停止監控
            MessageBox.Show("讀取記憶體使用率時發生錯誤: " & ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tmrUpdate.Stop()
            isMonitoring = False
            btnStart.Text = "啟動監控"
        End Try
    End Sub

    ''' <summary>
    ''' 表單關閉時清理資源
    ''' </summary>
    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        memoryCounter?.Dispose()
        MyBase.OnFormClosing(e)
    End Sub
End Class
