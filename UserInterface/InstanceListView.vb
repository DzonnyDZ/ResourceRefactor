' Copyright (c) Microsoft Corporation.  All rights reserved.
Imports System.Windows.Forms
Imports System.Resources
Imports System.Collections.ObjectModel
Imports EnvDTE

''' <summary>
''' An extension to tree view for displaying hard coded string instances
''' </summary>
''' <remarks></remarks>
Public Class InstanceListView
    Inherits System.Windows.Forms.TreeView


#Region "Public Properties"

    ''' <summary>
    ''' Gets a read only collection of instances that are selected by user to be replaced
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SelectedInstances() As ReadOnlyCollection(Of Common.ExtractToResourceActionSite)
        Get
            Dim collection As Collection(Of Common.ExtractToResourceActionSite) = New Collection(Of Common.ExtractToResourceActionSite)()
            For Each fileNode As TreeNode In Me.Nodes
                For Each instanceNode As TreeNode In fileNode.Nodes
                    If instanceNode.Checked Then
                        collection.Add(CType(instanceNode.Tag, Common.ExtractToResourceActionSite))
                    End If
                Next
            Next
            Return New ReadOnlyCollection(Of Common.ExtractToResourceActionSite)(collection)
        End Get
    End Property

    ''' <summary>
    ''' Returns the file related to node that is currently selected. If no nodes are selected
    ''' Nothing is returned.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property CurrentSelectedFile() As ProjectItem
        Get
            If Me.SelectedNode IsNot Nothing Then
                If Me.SelectedNode.Parent Is Nothing Then
                    Return CType(Me.SelectedNode.Tag, ProjectItem)
                Else
                    Return CType(Me.SelectedNode.Parent.Tag, ProjectItem)
                End If
            Else
                Return Nothing
            End If
        End Get
    End Property

#End Region

    ''' <summary>
    ''' Creates a new instance of Instance List viewer
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        MyBase.New()
        Me.CheckBoxes = True
    End Sub

    ''' <summary>
    ''' Fills the list view with instances from the given collection. 
    ''' </summary>
    ''' <param name="list">A list of BaseHardCodedString object instances</param>
    ''' <remarks></remarks>
    Public Sub BindData(ByVal list As ReadOnlyCollection(Of Common.ExtractToResourceActionSite))
        Me.BeginUpdate()
        Me.Nodes.Clear()
        For Each instance As Common.ExtractToResourceActionSite In list
            Dim fileNode As TreeNode = Nothing
            If Me.Nodes.ContainsKey(instance.StringToExtract.Parent.FileNames(1)) Then
                fileNode = Me.Nodes(instance.StringToExtract.Parent.FileNames(1))
            Else
                fileNode = New TreeNode()
                fileNode.Text = instance.StringToExtract.Parent.Name
                fileNode.Tag = instance.StringToExtract.Parent
                fileNode.Name = instance.StringToExtract.Parent.FileNames(1)
                fileNode.Checked = True
                fileNode.Expand()
                Me.Nodes.Add(fileNode)
            End If
            Dim instanceNode As TreeNode = New TreeNode()
            instanceNode.Text = "Line " & (instance.StringToExtract.StartingLine + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture) & ", Column: " & (instance.StringToExtract.StartIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture)
            instanceNode.Tag = instance
            instanceNode.Checked = True
            fileNode.Nodes.Add(instanceNode)
        Next
        Me.EndUpdate()
    End Sub

    ''' <summary>
    ''' Checks if a file node is checked/unchecked and performs the behaviour on all children nodes.
    ''' </summary>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub OnAfterCheck(ByVal e As System.Windows.Forms.TreeViewEventArgs)
        If e.Node.Nodes.Count > 0 Then
            Dim checkedState As Boolean = e.Node.Checked
            For Each node As TreeNode In e.Node.Nodes
                node.Checked = checkedState
            Next
        End If
        MyBase.OnAfterCheck(e)
    End Sub

End Class
