<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.lblMemoryUsage = New System.Windows.Forms.Label()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.tmrUpdate = New System.Windows.Forms.Timer(Me.components)
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft JhengHei UI", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.lblTitle.Location = New System.Drawing.Point(200, 50)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(200, 24)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "記憶體使用率監控"
        '
        'lblMemoryUsage
        '
        Me.lblMemoryUsage.AutoSize = True
        Me.lblMemoryUsage.Font = New System.Drawing.Font("Microsoft JhengHei UI", 36.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.lblMemoryUsage.ForeColor = System.Drawing.Color.DarkBlue
        Me.lblMemoryUsage.Location = New System.Drawing.Point(150, 120)
        Me.lblMemoryUsage.Name = "lblMemoryUsage"
        Me.lblMemoryUsage.Size = New System.Drawing.Size(150, 61)
        Me.lblMemoryUsage.TabIndex = 1
        Me.lblMemoryUsage.Text = "0 %"
        '
        'btnStart
        '
        Me.btnStart.Font = New System.Drawing.Font("Microsoft JhengHei UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.btnStart.Location = New System.Drawing.Point(120, 250)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(150, 50)
        Me.btnStart.TabIndex = 2
        Me.btnStart.Text = "啟動監控"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Font = New System.Drawing.Font("Microsoft JhengHei UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.btnExit.Location = New System.Drawing.Point(320, 250)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(150, 50)
        Me.btnExit.TabIndex = 3
        Me.btnExit.Text = "結束程式"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'tmrUpdate
        '
        Me.tmrUpdate.Interval = 1000
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(600, 400)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.lblMemoryUsage)
        Me.Controls.Add(Me.lblTitle)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "記憶體使用率監控系統"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblTitle As Label
    Friend WithEvents lblMemoryUsage As Label
    Friend WithEvents btnStart As Button
    Friend WithEvents btnExit As Button
    Friend WithEvents tmrUpdate As Timer

End Class
