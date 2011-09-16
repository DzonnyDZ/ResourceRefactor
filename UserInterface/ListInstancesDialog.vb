' Copyright (c) Microsoft Corporation.  All rights reserved.
Imports System.Windows.Forms
Imports System.Collections.ObjectModel
Imports EnvDTE

''' <summary>
''' Dialog used to list instances found during search and preview changes on those instances. 
''' This dialog also allows user to select which of the shown instances to refactor.
''' </summary>
''' <remarks></remarks>
Public Class ListInstancesDialog

    ''' <summary>
    ''' Selected resource file
    ''' </summary>
    ''' <remarks></remarks>
    Private _resourceFile As Common.ResourceFile

    ''' <summary>
    ''' Selected resource name
    ''' </summary>
    ''' <remarks></remarks>
    Private _resourceName As String

    ''' <summary>
    ''' Gets a collection of string instances to replace
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Value is valid only when DialogResult.OK is returned</remarks>
    Public ReadOnly Property InstancesToReplace() As ReadOnlyCollection(Of Common.ExtractToResourceActionSite)
        Get
            Return Me.uxInstanceList.SelectedInstances
        End Get
    End Property

    ''' <summary>
    ''' Shows the dialog previewing given list of instance changes for the provided resource entry
    ''' </summary>
    ''' <param name="instanceList">List of text instances to preview</param>
    ''' <param name="resourceFile">Resource file containing the resource entry</param>
    ''' <param name="resourceName">Name of the resource entry</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shadows Function ShowDialog(ByVal instanceList As ReadOnlyCollection(Of Common.ExtractToResourceActionSite), _
                                       ByVal resourceFile As Common.ResourceFile, ByVal resourceName As String) As DialogResult
        Me.uxInstanceList.BindData(instanceList)
        Me._resourceFile = resourceFile
        Me._resourceName = resourceName
        Return MyBase.ShowDialog()
    End Function


    ''' <summary>
    ''' Handles preview changes after user changes their selection
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub InstanceList_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles uxInstanceList.AfterSelect
        Dim item As ProjectItem = Me.uxInstanceList.CurrentSelectedFile
        Dim document As TextDocument = Common.BaseHardCodedString.GetDocumentForItem(item)
        If document IsNot Me.uxPreview.ActiveDocument Then
            Me.RefreshPreview()
        End If
        If TypeOf e.Node.Tag Is Common.ExtractToResourceActionSite Then
            Me.uxPreview.ShowInstance(CType(e.Node.Tag, Common.ExtractToResourceActionSite))
        End If
    End Sub

    ''' <summary>
    ''' Refreshes the preview box with the currently selected list of instaces
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub RefreshPreview()
        If Me.uxInstanceList.CurrentSelectedFile IsNot Nothing Then
            Dim item As ProjectItem = Me.uxInstanceList.CurrentSelectedFile
            Me.uxPreview.ActiveDocument = Common.BaseHardCodedString.GetDocumentForItem(item)
            Me.uxPreview.ActiveChangeList = Me.uxInstanceList.SelectedInstances
            Me.uxPreview.RefreshPreview(Me._resourceFile, Me._resourceName)
        End If
    End Sub

    ''' <summary>
    ''' Refreshes the preview view after user has changes checked state in one of the instances
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub InstanceList_AfterCheck(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles uxInstanceList.AfterCheck
        Me.RefreshPreview()
    End Sub

    Private Sub uxCancelButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles uxCancelButton.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub uxOkButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles uxOkButton.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub ListInstancesDialog_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If Not Me.DialogResult = System.Windows.Forms.DialogResult.OK Then
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        End If
    End Sub
End Class