' Copyright (c) Microsoft Corporation.  All rights reserved.
Imports System.Windows.Forms
Imports System.Collections.ObjectModel

''' <summary>
''' This is the main dialog for the resource refactoring tool.
''' </summary>
''' <remarks></remarks>
Public Class RefactorStringDialog

    ''' <summary>
    ''' Hard coded string instance to refactor
    ''' </summary>
    ''' <remarks></remarks>
    Private stringToRefactor As Common.BaseHardCodedString

    ''' <summary>
    ''' Shows the refactor string dialog
    ''' </summary>
    ''' <param name="refactorSite">Site to be refactored</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shadows Function ShowDialog(ByVal refactorSite As Common.ExtractToResourceActionSite) As DialogResult
        If refactorSite Is Nothing Then
            Throw New ArgumentNullException("refactorSite")
        End If
        Me.stringToRefactor = refactorSite.StringToExtract
        Dim resourceCollection As Common.ResourceFileCollection = _
                    New Common.ResourceFileCollection(Me.stringToRefactor.Parent.ContainingProject, _
                    New Common.FilterMethod(AddressOf refactorSite.ActionObject.IsValidResourceFile))
        options.DisplayMatchOptions(refactorSite, resourceCollection)
        Me.LoadSettings(resourceCollection)
        Return MyBase.ShowDialog()
    End Function

    ''' <summary>
    ''' Handles OK button click event, prepares a collection of BaseHardCodedString instances and replaces them with the resource
    ''' user has choosen.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>As localization is not supported right now, CA1300 is suppressed.</remarks>
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions")> _
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.ValidateChildren() Then
            If Me.options.SelectedResourceName IsNot Nothing And _
                Me.options.SelectedResourceFile IsNot Nothing Then
                Dim resourceFile As Common.ResourceFile = Me.options.SelectedResourceFile
                ' Prepare the list of instances to replace
                Dim instancesToReplace As ReadOnlyCollection(Of Common.ExtractToResourceActionSite) = Nothing
                If Me.options.Options = ResourceReplaceOption.ReplaceCurrentInstance Then
                    Dim tempCollection As Collection(Of Common.ExtractToResourceActionSite)
                    tempCollection = New Collection(Of Common.ExtractToResourceActionSite)()
                    tempCollection.Add(New Common.ExtractToResourceActionSite(Me.stringToRefactor))
                    instancesToReplace = New ReadOnlyCollection(Of Common.ExtractToResourceActionSite)(tempCollection)
                ElseIf Me.options.Options = ResourceReplaceOption.ReplaceCurrentFileOnly Then
                    instancesToReplace = CreateActionSitesFromInstanceList(Common.BaseHardCodedString.FindAllInstancesInDocument(Me.stringToRefactor.Parent, Me.stringToRefactor.RawValue))
                Else
                    instancesToReplace = CreateActionSitesFromInstanceList(Common.BaseHardCodedString.FindAllInstancesInProject(Me.stringToRefactor.Parent.ContainingProject, Me.stringToRefactor.RawValue))
                End If

                ' Preview instances and get the new list
                If Me.uxPreviewCheckbox.Checked Then
                    Dim previewDialog As ListInstancesDialog = New ListInstancesDialog()
                    If previewDialog.ShowDialog(instancesToReplace, Me.options.SelectedResourceFile, Me.options.SelectedResourceName) = Windows.Forms.DialogResult.OK Then
                        instancesToReplace = previewDialog.InstancesToReplace
                    Else
                        Exit Sub
                    End If
                End If

                'Replace instances
                If (Me.ReplaceInstances(instancesToReplace, resourceFile, Me.options.SelectedResourceName)) Then
                    Me.DialogResult = System.Windows.Forms.DialogResult.OK
                    Me.SaveSettings()
                    Me.Close()
                End If
            Else
                MessageBox.Show(My.Resources.Strings.NoResourceSelected, My.Resources.Strings.ErrorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Replace all instances of text in the list with a reference to provided resource entry
    ''' </summary>
    ''' <param name="list">Collection of BaseHardCodedString objects</param>
    ''' <param name="resourceFile">File containing the resource entry</param>
    ''' <param name="resourceName">Name of the resource entry</param>
    ''' <remarks>As localization is not supported right now, CA1300 is suppressed.</remarks>
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions")> _
    Private Function ReplaceInstances(ByVal list As ReadOnlyCollection(Of Common.ExtractToResourceActionSite), ByVal resourceFile As Common.ResourceFile, ByVal resourceName As String) As Boolean
        If list Is Nothing Then
            Throw New ArgumentNullException("list")
        End If
        Dim result As Boolean = False
        If list.Count > 0 Then
            Try
                If Me.options.IsCreateResourceChecked Then
                    resourceFile.AddResource(resourceName, Me.stringToRefactor.Value, String.Empty)
                    resourceFile.SaveFile()
                End If
                For Each instance As Common.ExtractToResourceActionSite In list
                    instance.ExtractStringToResource(resourceFile, resourceName)
                Next
                result = True
            Catch exception As Common.FileCheckoutException
                MessageBox.Show(String.Format(System.Globalization.CultureInfo.CurrentUICulture, _
                                My.Resources.Strings.FileCheckoutError, exception.FileName), _
                                My.Resources.Strings.ErrorCaption, _
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            Catch exception As Common.FileReadOnlyException
                MessageBox.Show(String.Format(System.Globalization.CultureInfo.CurrentUICulture, _
                                My.Resources.Strings.FileReadOnlyError, exception.FileName), _
                                My.Resources.Strings.ErrorCaption, _
                                MessageBoxButtons.OK, MessageBoxIcon.Error)

            End Try
        End If
        Return result
    End Function

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub uxSendFeedbackLink_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles uxSendFeedbackLink.LinkClicked
        System.Diagnostics.Process.Start(My.Resources.Strings.SendFeedbackURL)
    End Sub

    ''' <summary>
    ''' Loads user settings from the assembly settings file
    ''' </summary>
    ''' <param name="resourceFiles">Initial list of resource files to get the last selected resource file</param>
    Private Sub LoadSettings(ByVal resourceFiles As Common.ResourceFileCollection)
        Me.uxPreviewCheckbox.Checked = My.Settings.PreviewChangesChecked
        Me.options.SelectedResourceFile = resourceFiles.GetResourceFile(My.Settings.MostRecentResourceFileName)
        If My.Settings.ReplaceSetting <> -1 Then
            Try
                Dim replaceOption As ResourceReplaceOption = CType(My.Settings.ReplaceSetting, ResourceReplaceOption)
                Me.options.Options = replaceOption
            Catch e As InvalidCastException
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Saves user settings to the assembly settings file
    ''' </summary>
    Private Sub SaveSettings()
        If Me.options.SelectedResourceFile IsNot Nothing Then
            My.Settings.MostRecentResourceFileName = Me.options.SelectedResourceFile.DisplayName
        End If
        If Me.uxSaveSettings.Checked Then
            My.Settings.PreviewChangesChecked = Me.uxPreviewCheckbox.Checked
            My.Settings.ReplaceSetting = Me.options.Options
        End If
        My.Settings.Save()
    End Sub

    Private Shared Function CreateActionSitesFromInstanceList(ByVal list As ReadOnlyCollection(Of Common.BaseHardCodedString)) As ReadOnlyCollection(Of Common.ExtractToResourceActionSite)
        Dim siteList As Collection(Of Common.ExtractToResourceActionSite) = New Collection(Of Common.ExtractToResourceActionSite)
        For Each instance As Common.BaseHardCodedString In list
            siteList.Add(New Common.ExtractToResourceActionSite(instance))
        Next
        Return New ReadOnlyCollection(Of Common.ExtractToResourceActionSite)(siteList)
    End Function

    Private Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click
        MsgBox(My.Resources.Help)
    End Sub
End Class
