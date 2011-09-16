' Copyright (c) Microsoft Corporation.  All rights reserved.
Imports VB = Microsoft.VisualBasic
Imports System.Resources
Imports System.Collections.Generic
Imports EnvDTE
Imports System.Collections.ObjectModel


''' <summary>
''' A class used to desribe resource match results
''' </summary>
''' <remarks></remarks>
Public Class ResourceMatch

    Private _percentage As Double

    Private _resourceNode As ResXDataNode

    Private _resourceFile As ResourceFile

    ''' <summary>
    ''' Match percentage
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Percentage() As Double
        Get
            Return Me._percentage
        End Get
    End Property

    ''' <summary>
    ''' Name of the resource
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ResourceName() As String
        Get
            Return Me._resourceNode.Name
        End Get
    End Property

    ''' <summary>
    ''' Reference to ResXDataNode to retrieve value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ResourceNode() As ResXDataNode
        Get
            Return Me._resourceNode
        End Get
    End Property

    ''' <summary>
    ''' Value of the resource
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Value() As String
        Get
            Return Me.ResourceFile.GetValue(Me.ResourceName).ToString()
        End Get
    End Property

    ''' <summary>
    ''' Reference to the resource file containing the resource matched
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ResourceFile() As ResourceFile
        Get
            Return Me._resourceFile
        End Get
    End Property

    ''' <summary>
    ''' Creates a new resource match descriptor
    ''' </summary>
    ''' <param name="percentage">Match percentage</param>
    ''' <param name="node">ResXDataNode object for the resource</param>
    ''' <param name="parent">Resource file containing the resource</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal percentage As Double, ByVal node As ResXDataNode, ByVal parent As ResourceFile)
        Me._percentage = percentage
        Me._resourceNode = node
        Me._resourceFile = parent
    End Sub

End Class

''' <summary>
''' This class represents a resource file in the project. It is used to add/read entries from resource files and generic to both
''' C# and VB.Net projects
''' </summary>
''' <remarks></remarks>
Public Class ResourceFile

#Region "Private Variables"

    Private projectItem As ProjectItem

    ''' <summary>
    ''' List of data nodes to be saved later.
    ''' </summary>
    ''' <remarks></remarks>
    Private savedDataNodes As List(Of ResXDataNode) = New List(Of ResXDataNode)

    ''' <summary>
    ''' List of metadata information read from the resource file
    ''' </summary>
    ''' <remarks></remarks>
    Private savedMetadata As New List(Of KeyValuePair(Of String, Object))

    Private entries As Dictionary(Of String, ResXDataNode)

    Private _fileNameSpace As String

    ''' <summary>
    ''' Returns the short file name form of the resource file
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private ReadOnly Property shortFileName() As String
        Get
            Return projectItem.Name
        End Get
    End Property

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets the name to be displayed to user for this resource file
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property DisplayName() As String
        Get
            If IsDefaultResXFile() Then
                Return "(Default resources)"
            Else
                Return Me.shortFileName
            End If
        End Get
    End Property

    ''' <summary>
    ''' Gets the namespace of the code file generated for this resource file
    ''' </summary>
    ''' <value></value>
    ''' <returns>Nothing if failed, namespace otherwise</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FileNamespace() As String
        Get
            Dim namespaceValue As String = Nothing
            If (Me._fileNameSpace Is Nothing) Then
                For Each item As ProjectItem In Me.Item.ProjectItems
                    If item.Name.Contains(".Designer.") Then
                        If (item.FileCodeModel IsNot Nothing) Then
                            Dim element As CodeElement = Nothing
                            For Each subElement As CodeElement In item.FileCodeModel.CodeElements
                                element = ResourceFile.FindNamespaceElement(subElement)
                                If (element IsNot Nothing) Then
                                    namespaceValue = element.FullName
                                    Me._fileNameSpace = namespaceValue
                                    Exit For
                                End If
                            Next
                        End If
                        Exit For
                    End If
                Next
            ElseIf (Me._fileNameSpace.Length > 0) Then
                namespaceValue = Me._fileNameSpace
            End If
            Return namespaceValue
        End Get
    End Property

    ''' <summary>
    ''' Gets full path file name for the resource file
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FileName() As String
        Get
            Return projectItem.FileNames(0)
        End Get
    End Property


    ''' <summary>
    ''' Gets the custom tool name property for this resource file
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CustomToolName() As String
        Get
            Return CStr(projectItem.Properties.Item("CustomTool").Value)
        End Get
        Set(ByVal value As String)
            projectItem.Properties.Item("CustomTool").Value = value
        End Set
    End Property

    ''' <summary>
    ''' Gets custom tool namespace property of the resource file object
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property CustomToolNamespace() As String
        Get
            Dim returnValue As String = String.Empty
            Try
                Dim namespaceProperty As [Property] = Me.Item.Properties.Item("CustomToolNamespace")
                If namespaceProperty IsNot Nothing Then
                    returnValue = namespaceProperty.Value.ToString()
                End If
            Catch e As ArgumentException
            End Try
            Return returnValue
        End Get
    End Property
    ''' <summary>
    ''' Returns the project item related to this resource file.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Item() As ProjectItem
        Get
            Return Me.projectItem
        End Get
    End Property

    ''' <summary>
    ''' Returns all resources found in the resource file, keyed by the resource name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>This collection should be treated as a read only collection as changes to this collection will not be 
    ''' reflected in the resource file.</remarks>
    Public ReadOnly Property Resources() As Dictionary(Of String, ResXDataNode)
        Get
            If Me.entries Is Nothing Then
                Me.Refresh()
            End If
            Return Me.entries
        End Get
    End Property

#End Region

    ''' <summary>
    ''' Creates a new instance of ResourceFile object.
    ''' </summary>
    ''' <param name="projectItem">Project item pointing to resource file in the Visual Studio project</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal projectItem As ProjectItem)
        Me.projectItem = projectItem
    End Sub

    ''' <summary>
    ''' Finds all resource entries that has a value similar to one provided
    ''' </summary>
    ''' <param name="Text">Text to search for</param>
    ''' <returns>A list of MatchResult objects</returns>
    ''' <remarks></remarks>
    Public Function GetAllMatches(ByVal text As String) As Collection(Of ResourceMatch)
        Dim Results As New Collection(Of ResourceMatch)
        For Each resource As ResXDataNode In Me.Resources.Values
            If resource.GetValueTypeName(New System.Reflection.AssemblyName() {}).Equals(GetType(String).AssemblyQualifiedName) _
            AndAlso resource.FileRef Is Nothing Then
                Dim match As ResourceMatch = New ResourceMatch(StringMatch.Match(Me.GetValue(resource.Name).ToString(), text), resource, Me)
                Results.Add(match)
            End If
        Next
        Return Results
    End Function

    ''' <summary>
    ''' Checks if this resource file is the default resource file for the project
    ''' </summary>
    Public Function IsDefaultResXFile() As Boolean
        Return Me.FileName.Contains("My Project\Resources.resx")
    End Function

    ''' <summary>
    ''' Adds a text resource to this resource file. 
    ''' </summary>
    ''' <param name="resourceName">Name of the resource</param>
    ''' <param name="resourceValue">Text that is stored in the resources</param>
    ''' <param name="resourceComment">Comment for this resource entry</param>
    ''' <remarks>Method will throw ArgumentException if a resource with the same name already exists</remarks>
    Public Sub AddResource(ByVal resourceName As String, ByVal resourceValue As String, ByVal resourceComment As String)
        If Me.Contains(resourceName) Then
            Throw New ArgumentException("Resource already exists", "resourceName")
        End If
        For Each dataNode As ResXDataNode In savedDataNodes
            If dataNode.Name.Equals(resourceName, StringComparison.InvariantCultureIgnoreCase) Then
                Throw New ArgumentException("Resource already exists", "resourceName")
            End If
        Next
        Dim node As New ResXDataNode(resourceName, resourceValue)
        node.Comment = resourceComment
        savedDataNodes.Add(node)
    End Sub

    ''' <summary>
    ''' Check if resource file contains a resource with the provided name
    ''' </summary>
    ''' <param name="resourceName">Name of the resource to look for</param>
    ''' <returns></returns>
    ''' <remarks>Check is performed on the dictionary stored in Entries parameter. It does not include unsaved resource entries</remarks>
    Public Function Contains(ByVal resourceName As String) As Boolean
        Return Me.Resources.ContainsKey(resourceName)
    End Function

    ''' <summary>
    ''' Gets the value of a resource defined in the resource file.
    ''' </summary>
    ''' <param name="resourceName">Name of the resource</param>
    ''' <returns>The value for the resource, or ArgumentException if resource is not defined</returns>
    ''' <remarks></remarks>
    ''' <exception cref="System.ArgumentException">Throws ArgumentException if resource could not be located as resource values can be null</exception>
    Public Function GetValue(ByVal resourceName As String) As Object
        If Not Me.Contains(resourceName) Then
            Throw New ArgumentException("Resource could not be found", "resourceName")
        End If
        Return Me.Resources(resourceName).GetValue(New System.Reflection.AssemblyName() {})
    End Function

    ''' <summary>
    ''' Saves changes to the resource file.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SaveFile()
        Dim writer As ResXResourceWriter = Nothing
        Try
            If Not ExtensibilityMethods.CheckoutItem(Me.Item) Then
                Throw New FileCheckoutException(Me.Item.Name)
            End If
            Dim attributes As IO.FileAttributes = IO.File.GetAttributes(Me.FileName)
            If (attributes And IO.FileAttributes.ReadOnly) = IO.FileAttributes.ReadOnly Then
                Throw New FileReadOnlyException(Me.Item.Name)
            End If
            writer = New ResXResourceWriter(Me.FileName)
            Me.Refresh()
            For Each node As ResXDataNode In Me.Resources.Values
                writer.AddResource(node)
            Next

            'Write all saved resource entries
            For Each node As ResXDataNode In savedDataNodes
                Debug.Assert(node IsNot Nothing)
                If Not Me.Resources.ContainsKey(node.Name) Then
                    writer.AddResource(node)
                End If
            Next

            'Next, write all the saved metadata.
            For Each entry As KeyValuePair(Of String, Object) In savedMetadata
                writer.AddMetadata(entry.Key, entry.Value)
            Next
        Finally
            If writer IsNot Nothing Then
                writer.Generate()
                writer.Close()
                Me.savedDataNodes.Clear()
            End If
        End Try
        Me.RunCustomTool()
    End Sub

    ''' <summary>
    ''' Request the custom tool for project item assigned to this resource to be executed
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RunCustomTool()
        If Me.projectItem IsNot Nothing Then
            Dim VsProjectItem As VSLangProj.VSProjectItem = TryCast(projectItem.Object, VSLangProj.VSProjectItem)
            If VsProjectItem IsNot Nothing Then
                VsProjectItem.RunCustomTool()
            End If
        End If
    End Sub

    ''' <summary>
    ''' Refresh the dictionary of resource entries from the latest copy of the file
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Refresh()
        Me.entries = New Dictionary(Of String, ResXDataNode)(System.StringComparer.InvariantCultureIgnoreCase)
        Using reader As ResXResourceReader = New ResXResourceReader(Me.FileName)
            reader.UseResXDataNodes = True
            Dim resources As IDictionaryEnumerator = reader.GetEnumerator()
            While resources.MoveNext()
                entries.Add(resources.Key.ToString(), TryCast(resources.Value, ResXDataNode))
            End While

            Dim metadatas As IDictionaryEnumerator = reader.GetMetadataEnumerator()
            While metadatas.MoveNext()
                Me.savedMetadata.Add(New KeyValuePair(Of String, Object)(metadatas.Key.ToString(), metadatas.Value))
            End While
        End Using
    End Sub

    Public Overrides Function ToString() As String
        Return Me.DisplayName
    End Function

    ''' <summary>
    ''' Recurses in to code elements to find a Namespace element
    ''' </summary>
    ''' <param name="element">Code element to recurse into</param>
    ''' <returns>Namespace code element if one found, Nothing (null) otherwise</returns>
    ''' <remarks></remarks>
    Private Shared Function FindNamespaceElement(ByVal element As CodeElement) As CodeElement
        Dim returnElement As CodeElement = Nothing
        If element.Kind = vsCMElement.vsCMElementNamespace Then
            returnElement = element
        Else
            For Each subElement As CodeElement In element.Children
                returnElement = FindNamespaceElement(subElement)
                If (returnElement IsNot Nothing) Then
                    Exit For
                End If
            Next
        End If
        Return returnElement
    End Function
End Class

