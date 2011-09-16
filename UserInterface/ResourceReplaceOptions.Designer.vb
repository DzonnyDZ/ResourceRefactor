<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ResourceReplaceOptions
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ResourceReplaceOptions))
        Me.BottomToolStripPanel = New System.Windows.Forms.ToolStripPanel
        Me.TopToolStripPanel = New System.Windows.Forms.ToolStripPanel
        Me.RightToolStripPanel = New System.Windows.Forms.ToolStripPanel
        Me.LeftToolStripPanel = New System.Windows.Forms.ToolStripPanel
        Me.uxCreateNewResource = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.uxNewResourceName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.uxTextToReplace = New System.Windows.Forms.TextBox
        Me.uxErrorProvider = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.Label3 = New System.Windows.Forms.Label
        Me.uxRefactorChoicePanel = New System.Windows.Forms.Panel
        Me.uxOptionChangeAll = New System.Windows.Forms.RadioButton
        Me.uxOptionChangeFileOnly = New System.Windows.Forms.RadioButton
        Me.uxOptionChangeCurrentOnly = New System.Windows.Forms.RadioButton
        Me.uxResourceFileSelector = New Microsoft.VSPowerToys.ResourceRefactor.ResourceFileListDropDown
        Me.uxResourceView = New Microsoft.VSPowerToys.ResourceRefactor.ResourceGridView
        CType(Me.uxErrorProvider, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.uxRefactorChoicePanel.SuspendLayout()
        CType(Me.uxResourceView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BottomToolStripPanel
        '
        resources.ApplyResources(Me.BottomToolStripPanel, "BottomToolStripPanel")
        Me.BottomToolStripPanel.Name = "BottomToolStripPanel"
        Me.BottomToolStripPanel.Orientation = System.Windows.Forms.Orientation.Horizontal
        Me.BottomToolStripPanel.RowMargin = New System.Windows.Forms.Padding(3, 0, 0, 0)
        '
        'TopToolStripPanel
        '
        resources.ApplyResources(Me.TopToolStripPanel, "TopToolStripPanel")
        Me.TopToolStripPanel.Name = "TopToolStripPanel"
        Me.TopToolStripPanel.Orientation = System.Windows.Forms.Orientation.Horizontal
        Me.TopToolStripPanel.RowMargin = New System.Windows.Forms.Padding(3, 0, 0, 0)
        '
        'RightToolStripPanel
        '
        resources.ApplyResources(Me.RightToolStripPanel, "RightToolStripPanel")
        Me.RightToolStripPanel.Name = "RightToolStripPanel"
        Me.RightToolStripPanel.Orientation = System.Windows.Forms.Orientation.Horizontal
        Me.RightToolStripPanel.RowMargin = New System.Windows.Forms.Padding(3, 0, 0, 0)
        '
        'LeftToolStripPanel
        '
        resources.ApplyResources(Me.LeftToolStripPanel, "LeftToolStripPanel")
        Me.LeftToolStripPanel.Name = "LeftToolStripPanel"
        Me.LeftToolStripPanel.Orientation = System.Windows.Forms.Orientation.Horizontal
        Me.LeftToolStripPanel.RowMargin = New System.Windows.Forms.Padding(3, 0, 0, 0)
        '
        'uxCreateNewResource
        '
        resources.ApplyResources(Me.uxCreateNewResource, "uxCreateNewResource")
        Me.uxCreateNewResource.CausesValidation = False
        Me.uxCreateNewResource.Checked = True
        Me.uxCreateNewResource.CheckState = System.Windows.Forms.CheckState.Checked
        Me.uxCreateNewResource.Name = "uxCreateNewResource"
        Me.uxCreateNewResource.UseVisualStyleBackColor = True
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'uxNewResourceName
        '
        resources.ApplyResources(Me.uxNewResourceName, "uxNewResourceName")
        Me.uxNewResourceName.Name = "uxNewResourceName"
        '
        'Label2
        '
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.Name = "Label2"
        '
        'uxTextToReplace
        '
        resources.ApplyResources(Me.uxTextToReplace, "uxTextToReplace")
        Me.uxTextToReplace.Name = "uxTextToReplace"
        Me.uxTextToReplace.ReadOnly = True
        Me.uxTextToReplace.TabStop = False
        '
        'uxErrorProvider
        '
        Me.uxErrorProvider.ContainerControl = Me
        '
        'Label3
        '
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.Name = "Label3"
        '
        'uxRefactorChoicePanel
        '
        Me.uxRefactorChoicePanel.Controls.Add(Me.uxOptionChangeAll)
        Me.uxRefactorChoicePanel.Controls.Add(Me.uxOptionChangeFileOnly)
        Me.uxRefactorChoicePanel.Controls.Add(Me.uxOptionChangeCurrentOnly)
        resources.ApplyResources(Me.uxRefactorChoicePanel, "uxRefactorChoicePanel")
        Me.uxRefactorChoicePanel.Name = "uxRefactorChoicePanel"
        '
        'uxOptionChangeAll
        '
        resources.ApplyResources(Me.uxOptionChangeAll, "uxOptionChangeAll")
        Me.uxOptionChangeAll.Name = "uxOptionChangeAll"
        Me.uxOptionChangeAll.UseVisualStyleBackColor = True
        '
        'uxOptionChangeFileOnly
        '
        resources.ApplyResources(Me.uxOptionChangeFileOnly, "uxOptionChangeFileOnly")
        Me.uxOptionChangeFileOnly.Name = "uxOptionChangeFileOnly"
        Me.uxOptionChangeFileOnly.UseVisualStyleBackColor = True
        '
        'uxOptionChangeCurrentOnly
        '
        resources.ApplyResources(Me.uxOptionChangeCurrentOnly, "uxOptionChangeCurrentOnly")
        Me.uxOptionChangeCurrentOnly.Checked = True
        Me.uxOptionChangeCurrentOnly.Name = "uxOptionChangeCurrentOnly"
        Me.uxOptionChangeCurrentOnly.TabStop = True
        Me.uxOptionChangeCurrentOnly.UseVisualStyleBackColor = True
        '
        'uxResourceFileSelector
        '
        resources.ApplyResources(Me.uxResourceFileSelector, "uxResourceFileSelector")
        Me.uxResourceFileSelector.ExtractResourceAction = Nothing
        Me.uxResourceFileSelector.MaximumSize = New System.Drawing.Size(1600, 27)
        Me.uxResourceFileSelector.MinimumSize = New System.Drawing.Size(200, 27)
        Me.uxResourceFileSelector.Name = "uxResourceFileSelector"
        Me.uxResourceFileSelector.SelectedResourceFile = Nothing
        '
        'uxResourceView
        '
        Me.uxResourceView.AllowUserToAddRows = False
        Me.uxResourceView.AllowUserToDeleteRows = False
        Me.uxResourceView.AllowUserToResizeRows = False
        resources.ApplyResources(Me.uxResourceView, "uxResourceView")
        Me.uxResourceView.AutoGenerateColumns = False
        Me.uxResourceView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.uxResourceView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.uxResourceView.MultiSelect = False
        Me.uxResourceView.Name = "uxResourceView"
        Me.uxResourceView.ReadOnly = True
        Me.uxResourceView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        '
        'ResourceReplaceOptions
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.uxRefactorChoicePanel)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.uxTextToReplace)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.uxNewResourceName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.uxCreateNewResource)
        Me.Controls.Add(Me.uxResourceFileSelector)
        Me.Controls.Add(Me.uxResourceView)
        Me.Controls.Add(Me.BottomToolStripPanel)
        Me.Controls.Add(Me.TopToolStripPanel)
        Me.Controls.Add(Me.RightToolStripPanel)
        Me.Controls.Add(Me.LeftToolStripPanel)
        Me.Name = "ResourceReplaceOptions"
        CType(Me.uxErrorProvider, System.ComponentModel.ISupportInitialize).EndInit()
        Me.uxRefactorChoicePanel.ResumeLayout(False)
        Me.uxRefactorChoicePanel.PerformLayout()
        CType(Me.uxResourceView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BottomToolStripPanel As System.Windows.Forms.ToolStripPanel
    Friend WithEvents TopToolStripPanel As System.Windows.Forms.ToolStripPanel
    Friend WithEvents RightToolStripPanel As System.Windows.Forms.ToolStripPanel
    Friend WithEvents LeftToolStripPanel As System.Windows.Forms.ToolStripPanel
    Friend WithEvents uxResourceFileSelector As Microsoft.VSPowerToys.ResourceRefactor.ResourceFileListDropDown
    Friend WithEvents uxResourceView As Microsoft.VSPowerToys.ResourceRefactor.ResourceGridView
    Friend WithEvents uxCreateNewResource As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents uxNewResourceName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents uxTextToReplace As System.Windows.Forms.TextBox
    Friend WithEvents uxErrorProvider As System.Windows.Forms.ErrorProvider
    Friend WithEvents uxRefactorChoicePanel As System.Windows.Forms.Panel
    Friend WithEvents uxOptionChangeAll As System.Windows.Forms.RadioButton
    Friend WithEvents uxOptionChangeFileOnly As System.Windows.Forms.RadioButton
    Friend WithEvents uxOptionChangeCurrentOnly As System.Windows.Forms.RadioButton
    Friend WithEvents Label3 As System.Windows.Forms.Label

End Class
