' Copyright (c) Microsoft Corporation.  All rights reserved.
Imports EnvDTE
Imports System.Windows.Forms

''' <summary>
''' A user control that lists resource files in a ResourceFileCollection and also
''' allows user to create a new resource
''' </summary>
''' <remarks></remarks>
Public Class ResourceFileListDropDown

    ''' <summary>
    ''' Raised when resource file selection has been changed by the user
    ''' </summary>
    ''' <remarks></remarks>
    Public Event SelectionChanged As EventHandler(Of EventArgs)

    Private _resourceFiles As Common.ResourceFileCollection
    Private lastSelectedItem As Common.ResourceFile

    Private _extractResourceAction As Common.IExtractResourceAction

    Public Property SelectedResourceFile() As Common.ResourceFile
        Get
            If Me.cboCreateNewResxFile.SelectedItem IsNot Nothing Then
                Return TryCast(Me.cboCreateNewResxFile.SelectedItem, Common.ResourceFile)
            End If
            Return Nothing
        End Get
        Set(ByVal value As Common.ResourceFile)
            If (Not value Is Nothing) AndAlso Me.cboCreateNewResxFile.Items.Contains(value) Then
                Me.cboCreateNewResxFile.SelectedItem = value
            End If
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the current action object in use
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ExtractResourceAction() As Common.IExtractResourceAction
        Get
            Return Me._extractResourceAction
        End Get
        Set(ByVal value As Common.IExtractResourceAction)
            Me._extractResourceAction = value
        End Set
    End Property

    ''' <summary>
    ''' Fills the resource file combobox with entries from provided ResourceFileCollection
    ''' </summary>
    ''' <param name="resourceFiles">Resource file collection containing resource file entries</param>
    ''' <remarks></remarks>
    Public Sub FillResourceFilesDropDown(ByVal resourceFiles As Common.ResourceFileCollection)
        _resourceFiles = resourceFiles
        cboCreateNewResxFile.Items.Clear()
        If resourceFiles IsNot Nothing Then
            'Add all resource files to the "add new resource" dropdown
            For Each resourceFile As Common.ResourceFile In resourceFiles
                cboCreateNewResxFile.Items.Add(resourceFile)
            Next
            cboCreateNewResxFile.Items.Add(New CreateItem(My.Resources.Strings.CreateNewResourceText))
            If Me.cboCreateNewResxFile.Items.Count > 1 Then
                Me.cboCreateNewResxFile.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub cboCreateNewResxFile_CreateItemSelected(ByVal sender As Object, ByVal e As EventArgs) Handles cboCreateNewResxFile.CreateItemSelected
        Dim uxFileDialog As System.Windows.Forms.SaveFileDialog = New System.Windows.Forms.SaveFileDialog()
        uxFileDialog.DefaultExt = "resx"
        uxFileDialog.AddExtension = True
        uxFileDialog.Filter = "Resource File (*.resx)|*.resx"
        uxFileDialog.FilterIndex = 0
        uxFileDialog.InitialDirectory = IO.Path.GetDirectoryName(Me._resourceFiles.Project.FullName)
        If IO.Directory.Exists(IO.Path.Combine(uxFileDialog.InitialDirectory, Me._extractResourceAction.DefaultResourceFilePath)) Then
            uxFileDialog.InitialDirectory = IO.Path.Combine(uxFileDialog.InitialDirectory, Me._extractResourceAction.DefaultResourceFilePath)
        End If
        Dim newFile As Common.ResourceFile = Nothing
        Dim result As DialogResult = uxFileDialog.ShowDialog()
        If result = DialogResult.OK Then
            Dim writer As Resources.ResXResourceWriter = Nothing
            Try
                writer = New Resources.ResXResourceWriter(uxFileDialog.FileName)
                writer.Generate()
                writer.Close()
                writer = Nothing
                Try
                    Dim newProjectItem As ProjectItem = _resourceFiles.Project.ProjectItems.AddFromFile(uxFileDialog.FileName)
                    newFile = New Common.ResourceFile(newProjectItem)
                    Me._extractResourceAction.UpdateResourceFileProperties(newProjectItem)
                    Me._resourceFiles.RefreshListOfFiles()
                    Me.FillResourceFilesDropDown(Me._resourceFiles)
                Catch exception As System.Runtime.InteropServices.COMException
                    'Raised when project is locked, we don't do anything since user is already warned by TFS
                End Try
            Finally
                If writer IsNot Nothing Then
                    writer.Close()
                End If
            End Try
        End If
        If (Not newFile Is Nothing) AndAlso Me.cboCreateNewResxFile.Items.Contains(Me._resourceFiles.Item(newFile.DisplayName)) Then
            Me.cboCreateNewResxFile.SelectedItem = Me._resourceFiles.Item(newFile.DisplayName)
        ElseIf (Not Me.lastSelectedItem Is Nothing) AndAlso Me.cboCreateNewResxFile.Items.Contains(Me.lastSelectedItem) Then
            Me.cboCreateNewResxFile.SelectedItem = Me.lastSelectedItem
        End If
        RaiseEvent SelectionChanged(Me, Nothing)
    End Sub

    Private Sub cboCreateNewResxFile_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCreateNewResxFile.SelectionChangeCommitted
        RaiseEvent SelectionChanged(Me, Nothing)
    End Sub

    ''' <summary>
    ''' We save the last selected entry to return to it in case "Create new file" fails
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cboCreateNewResxFile_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCreateNewResxFile.SelectedIndexChanged
        If Not Me.cboCreateNewResxFile.SelectedItem Is Nothing Then
            Me.lastSelectedItem = TryCast(Me.cboCreateNewResxFile.SelectedItem, Common.ResourceFile)
        End If
    End Sub
End Class
