<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtYear = New System.Windows.Forms.TextBox
        Me.cmdBuildCards = New System.Windows.Forms.Button
        Me.lblStatus = New System.Windows.Forms.Label
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(27, 51)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Year:"
        Me.ToolTip1.SetToolTip(Me.Label1, "Hello Bill")
        '
        'txtYear
        '
        Me.txtYear.Location = New System.Drawing.Point(79, 49)
        Me.txtYear.Name = "txtYear"
        Me.txtYear.Size = New System.Drawing.Size(57, 20)
        Me.txtYear.TabIndex = 1
        '
        'cmdBuildCards
        '
        Me.cmdBuildCards.Location = New System.Drawing.Point(44, 98)
        Me.cmdBuildCards.Name = "cmdBuildCards"
        Me.cmdBuildCards.Size = New System.Drawing.Size(91, 24)
        Me.cmdBuildCards.TabIndex = 2
        Me.cmdBuildCards.Text = "Build Cards"
        Me.cmdBuildCards.UseVisualStyleBackColor = True
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(38, 166)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(39, 13)
        Me.lblStatus.TabIndex = 3
        Me.lblStatus.Text = "Label2"
        '
        'ToolTip1
        '
        Me.ToolTip1.Tag = "GetBent"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(323, 242)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.cmdBuildCards)
        Me.Controls.Add(Me.txtYear)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmMain"
        Me.Text = "BuildCards"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtYear As System.Windows.Forms.TextBox
    Friend WithEvents cmdBuildCards As System.Windows.Forms.Button
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip

End Class
