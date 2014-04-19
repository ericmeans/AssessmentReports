<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Options
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
        Me.components = New System.ComponentModel.Container()
        Me.btnReady = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtCurrentSemester = New System.Windows.Forms.TextBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtFirstScoreCol = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtLastScoreCol = New System.Windows.Forms.TextBox()
        Me.chkDebug = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'btnReady
        '
        Me.btnReady.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnReady.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnReady.Location = New System.Drawing.Point(193, 61)
        Me.btnReady.Name = "btnReady"
        Me.btnReady.Size = New System.Drawing.Size(110, 49)
        Me.btnReady.TabIndex = 4
        Me.btnReady.Text = "Ready"
        Me.btnReady.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Current Semester"
        '
        'txtCurrentSemester
        '
        Me.txtCurrentSemester.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCurrentSemester.Location = New System.Drawing.Point(106, 6)
        Me.txtCurrentSemester.Name = "txtCurrentSemester"
        Me.txtCurrentSemester.Size = New System.Drawing.Size(197, 20)
        Me.txtCurrentSemester.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtCurrentSemester, "Leave blank for all semesters")
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(12, 82)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(59, 28)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "Cancel"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 35)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "First Score Col:"
        '
        'txtFirstScoreCol
        '
        Me.txtFirstScoreCol.Location = New System.Drawing.Point(96, 32)
        Me.txtFirstScoreCol.Name = "txtFirstScoreCol"
        Me.txtFirstScoreCol.Size = New System.Drawing.Size(41, 20)
        Me.txtFirstScoreCol.TabIndex = 1
        Me.txtFirstScoreCol.Text = "J"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(143, 35)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Last:"
        '
        'txtLastScoreCol
        '
        Me.txtLastScoreCol.Location = New System.Drawing.Point(179, 32)
        Me.txtLastScoreCol.Name = "txtLastScoreCol"
        Me.txtLastScoreCol.Size = New System.Drawing.Size(45, 20)
        Me.txtLastScoreCol.TabIndex = 2
        Me.txtLastScoreCol.Text = "AH"
        '
        'chkDebug
        '
        Me.chkDebug.AutoSize = True
        Me.chkDebug.Location = New System.Drawing.Point(15, 59)
        Me.chkDebug.Name = "chkDebug"
        Me.chkDebug.Size = New System.Drawing.Size(117, 17)
        Me.chkDebug.TabIndex = 3
        Me.chkDebug.Text = "Include Debug Info"
        Me.chkDebug.UseVisualStyleBackColor = True
        '
        'Options
        '
        Me.AcceptButton = Me.btnReady
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(315, 122)
        Me.ControlBox = False
        Me.Controls.Add(Me.chkDebug)
        Me.Controls.Add(Me.txtLastScoreCol)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtFirstScoreCol)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txtCurrentSemester)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnReady)
        Me.Name = "Options"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Options"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnReady As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtCurrentSemester As System.Windows.Forms.TextBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtFirstScoreCol As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtLastScoreCol As System.Windows.Forms.TextBox
    Friend WithEvents chkDebug As System.Windows.Forms.CheckBox
End Class
