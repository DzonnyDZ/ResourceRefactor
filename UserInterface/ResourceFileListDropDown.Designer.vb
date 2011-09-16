<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ResourceFileListDropDown
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ResourceFileListDropDown))
        Me.lblResourceFile = New System.Windows.Forms.Label
        Me.cboCreateNewResxFile = New Microsoft.VSPowerToys.ResourceRefactor.ComboBoxWithCreate
        Me.SuspendLayout()
        '
        'lblResourceFile
        '
        resources.ApplyResources(Me.lblResourceFile, "lblResourceFile")
        Me.lblResourceFile.Name = "lblResourceFile"
        '
        'cboCreateNewResxFile
        '
        resources.ApplyResources(Me.cboCreateNewResxFile, "cboCreateNewResxFile")
        Me.cboCreateNewResxFile.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCreateNewResxFile.FormattingEnabled = True
        Me.cboCreateNewResxFile.MinimumSize = New System.Drawing.Size(100, 0)
        Me.cboCreateNewResxFile.Name = "cboCreateNewResxFile"
        '
        'ResourceFileListDropDown
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.cboCreateNewResxFile)
        Me.Controls.Add(Me.lblResourceFile)
        Me.MaximumSize = New System.Drawing.Size(1600, 27)
        Me.MinimumSize = New System.Drawing.Size(200, 27)
        Me.Name = "ResourceFileListDropDown"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboCreateNewResxFile As Microsoft.VSPowerToys.ResourceRefactor.ComboBoxWithCreate
    Friend WithEvents lblResourceFile As System.Windows.Forms.Label

End Class
