<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ListInstancesDialog
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ListInstancesDialog))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.uxCancelButton = New System.Windows.Forms.Button
        Me.uxOkButton = New System.Windows.Forms.Button
        Me.uxInstanceList = New Microsoft.VSPowerToys.ResourceRefactor.InstanceListView
        Me.uxPreview = New Microsoft.VSPowerToys.ResourceRefactor.CodeFilePreview
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        resources.ApplyResources(Me.SplitContainer1, "SplitContainer1")
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.uxInstanceList)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.uxPreview)
        '
        'Label2
        '
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.Name = "Label2"
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'uxCancelButton
        '
        resources.ApplyResources(Me.uxCancelButton, "uxCancelButton")
        Me.uxCancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.uxCancelButton.Name = "uxCancelButton"
        Me.uxCancelButton.UseVisualStyleBackColor = True
        '
        'uxOkButton
        '
        resources.ApplyResources(Me.uxOkButton, "uxOkButton")
        Me.uxOkButton.Name = "uxOkButton"
        Me.uxOkButton.UseVisualStyleBackColor = True
        '
        'uxInstanceList
        '
        resources.ApplyResources(Me.uxInstanceList, "uxInstanceList")
        Me.uxInstanceList.CheckBoxes = True
        Me.uxInstanceList.Name = "uxInstanceList"
        '
        'uxPreview
        '
        resources.ApplyResources(Me.uxPreview, "uxPreview")
        Me.uxPreview.Name = "uxPreview"
        Me.uxPreview.ReadOnly = True
        '
        'ListInstancesDialog
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.uxCancelButton
        Me.Controls.Add(Me.uxOkButton)
        Me.Controls.Add(Me.uxCancelButton)
        Me.Controls.Add(Me.SplitContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "ListInstancesDialog"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents uxInstanceList As Microsoft.VSPowerToys.ResourceRefactor.InstanceListView
    Friend WithEvents uxPreview As Microsoft.VSPowerToys.ResourceRefactor.CodeFilePreview
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents uxCancelButton As System.Windows.Forms.Button
    Friend WithEvents uxOkButton As System.Windows.Forms.Button
End Class
