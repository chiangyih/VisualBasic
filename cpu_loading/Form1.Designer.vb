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
        lblCpu = New Label()
        btnStart = New Button()
        btnExit = New Button()
        SuspendLayout()
        ' 
        ' lblCpu
        ' 
        lblCpu.Font = New Font("Segoe UI", 24.0F)
        lblCpu.Location = New Point(12, 9)
        lblCpu.Name = "lblCpu"
        lblCpu.Size = New Size(360, 120)
        lblCpu.TabIndex = 0
        lblCpu.Text = "CPU: -- %"
        lblCpu.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' btnStart
        ' 
        btnStart.Location = New Point(60, 150)
        btnStart.Name = "btnStart"
        btnStart.Size = New Size(100, 35)
        btnStart.TabIndex = 1
        btnStart.Text = "Start"
        btnStart.UseVisualStyleBackColor = True
        ' 
        ' btnExit
        ' 
        btnExit.Location = New Point(200, 150)
        btnExit.Name = "btnExit"
        btnExit.Size = New Size(100, 35)
        btnExit.TabIndex = 2
        btnExit.Text = "Exit"
        btnExit.UseVisualStyleBackColor = True
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7.0F, 15.0F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(384, 211)
        Controls.Add(btnExit)
        Controls.Add(btnStart)
        Controls.Add(lblCpu)
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False
        MinimizeBox = False
        Name = "Form1"
        StartPosition = FormStartPosition.CenterScreen
        Text = "CPU Usage"
        ResumeLayout(False)

    End Sub

    Friend WithEvents lblCpu As System.Windows.Forms.Label
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button

End Class
