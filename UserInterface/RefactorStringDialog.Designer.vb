<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RefactorStringDialog
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(RefactorStringDialog))
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.uxSendFeedbackLink = New System.Windows.Forms.LinkLabel()
        Me.uxPreviewCheckbox = New System.Windows.Forms.CheckBox()
        Me.uxSaveSettings = New System.Windows.Forms.CheckBox()
        Me.options = New Microsoft.VSPowerToys.ResourceRefactor.ResourceReplaceOptions()
        Me.cmdHelp = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'OK_Button
        '
        resources.ApplyResources(Me.OK_Button, "OK_Button")
        Me.OK_Button.Name = "OK_Button"
        '
        'Cancel_Button
        '
        resources.ApplyResources(Me.Cancel_Button, "Cancel_Button")
        Me.Cancel_Button.CausesValidation = False
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Name = "Cancel_Button"
        '
        'uxSendFeedbackLink
        '
        resources.ApplyResources(Me.uxSendFeedbackLink, "uxSendFeedbackLink")
        Me.uxSendFeedbackLink.Name = "uxSendFeedbackLink"
        Me.uxSendFeedbackLink.TabStop = True
        '
        'uxPreviewCheckbox
        '
        resources.ApplyResources(Me.uxPreviewCheckbox, "uxPreviewCheckbox")
        Me.uxPreviewCheckbox.Checked = True
        Me.uxPreviewCheckbox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.uxPreviewCheckbox.Name = "uxPreviewCheckbox"
        Me.uxPreviewCheckbox.UseVisualStyleBackColor = True
        '
        'uxSaveSettings
        '
        resources.ApplyResources(Me.uxSaveSettings, "uxSaveSettings")
        Me.uxSaveSettings.Name = "uxSaveSettings"
        Me.uxSaveSettings.UseVisualStyleBackColor = True
        '
        'options
        '
        resources.ApplyResources(Me.options, "options")
        Me.options.Name = "options"
        Me.options.Options = Microsoft.VSPowerToys.ResourceRefactor.ResourceReplaceOption.ReplaceCurrentInstance
        Me.options.SelectedResourceFile = Nothing
        '
        'cmdHelp
        '
        resources.ApplyResources(Me.cmdHelp, "cmdHelp")
        Me.cmdHelp.Name = "cmdHelp"
        Me.cmdHelp.UseVisualStyleBackColor = True
        '
        'RefactorStringDialog
        '
        Me.AcceptButton = Me.OK_Button
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.Controls.Add(Me.cmdHelp)
        Me.Controls.Add(Me.uxSaveSettings)
        Me.Controls.Add(Me.uxPreviewCheckbox)
        Me.Controls.Add(Me.uxSendFeedbackLink)
        Me.Controls.Add(Me.OK_Button)
        Me.Controls.Add(Me.Cancel_Button)
        Me.Controls.Add(Me.options)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "RefactorStringDialog"
        Me.ShowInTaskbar = False
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents options As ResourceReplaceOptions
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents uxSendFeedbackLink As System.Windows.Forms.LinkLabel
    Friend WithEvents uxPreviewCheckbox As System.Windows.Forms.CheckBox
    Friend WithEvents uxSaveSettings As System.Windows.Forms.CheckBox
    Friend WithEvents cmdHelp As System.Windows.Forms.Button

End Class
