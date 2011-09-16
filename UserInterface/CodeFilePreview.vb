' Copyright (c) Microsoft Corporation.  All rights reserved.
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.ObjectModel
Imports EnvDTE
Imports System.ComponentModel.Design
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Diagnostics.CodeAnalysis


''' <summary>
''' A preview window to preview multiple changes in a code file.
''' </summary>
Public Class CodeFilePreview
    Inherits RichTextBox

#Region "InstanceInformation definition"

    ''' <summary>
    ''' A structure to hold information about instance preview. Not all values may be valid at all times.
    ''' </summary>
    ''' <remarks></remarks>
    Private Class InstanceInformation

        Private _startingIndex As Integer
        Private _endingIndex As Integer
        Private _previewLength As Integer

        ''' <summary>
        ''' Starting index of the text related to this instance (relevant to start of the document)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property StartingIndex() As Integer
            Get
                Return _startingIndex
            End Get
            Set(ByVal value As Integer)
                _startingIndex = value
            End Set
        End Property

        ''' <summary>
        ''' Ending index of the text related to this instance
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property EndingIndex() As Integer
            Get
                Return _endingIndex
            End Get
            Set(ByVal value As Integer)
                _endingIndex = value
            End Set
        End Property

        ''' <summary>
        ''' Length of the reference preview related to this instance
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property PreviewLength() As Integer
            Get
                Return _previewLength
            End Get
            Set(ByVal value As Integer)
                _previewLength = value
            End Set
        End Property

    End Class
#End Region

    Private _activeChangeList As Collection(Of Common.ExtractToResourceActionSite)
    Private _activeDocument As EnvDTE.TextDocument
    Private _lastShownInstance As Common.ExtractToResourceActionSite

    ''' <summary>
    ''' A dictionary mapping line numbers to a sorted list of BaseHardCodedString objects
    ''' </summary>
    ''' <remarks></remarks>
    Private _previewEntries As SortedList(Of Integer, Common.ExtractToResourceActionSite) = _
        New SortedList(Of Integer, Common.ExtractToResourceActionSite)

    ''' <summary>
    ''' Dictionary mapping BaseHardCodedString instances to number of characters they add extra to that line
    ''' </summary>
    ''' <remarks></remarks>
    Private _instanceInformation As Dictionary(Of Common.ExtractToResourceActionSite, InstanceInformation) = _
        New Dictionary(Of Common.ExtractToResourceActionSite, InstanceInformation)()

    ''' <summary>
    ''' Collection of instances to be preview. All instances should belong to document references by ActiveDocument property 
    ''' otherwise they will be ignored.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>CA2227: Collection has to be specified from outside sources thus it can't only have "get" method.</remarks>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
    <SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")> _
    Public Property ActiveChangeList() As ReadOnlyCollection(Of Common.ExtractToResourceActionSite)
        Get
            If Me._activeChangeList Is Nothing Then
                Me._activeChangeList = New Collection(Of Common.ExtractToResourceActionSite)()
            End If
            Return New ReadOnlyCollection(Of Common.ExtractToResourceActionSite)(Me._activeChangeList)
        End Get
        Set(ByVal value As ReadOnlyCollection(Of Common.ExtractToResourceActionSite))
            If value IsNot Nothing Then
                Me._activeChangeList = New Collection(Of Common.ExtractToResourceActionSite)(value)
            End If
        End Set
    End Property

    ''' <summary>
    ''' Current active document shown on the preview window. Instances not in this document will be ignored
    ''' by the preview window
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
    Public Property ActiveDocument() As EnvDTE.TextDocument
        Get
            Return Me._activeDocument
        End Get
        Set(ByVal value As EnvDTE.TextDocument)
            Me._activeDocument = value
        End Set
    End Property

    ''' <summary>
    ''' Creates a new instance of preview window
    ''' </summary>
    ''' <remarks>CA2122: Native method has to be called to setup the text \box correctly.</remarks>
    <SuppressMessage("Microsoft.Security", "CA2122:DoNotIndirectlyExposeMethodsWithLinkDemands")> _
    Public Sub New()
        Me.ReadOnly = True
        Me.ScrollBars = RichTextBoxScrollBars.Both
        Me.WordWrap = False
        Me.Font = New Font(System.Drawing.FontFamily.GenericMonospace, 8)

        ' Set line numbers
        Dim param As New PARAFORMAT2
        param.cbSize = Marshal.SizeOf(param)
        NativeMethods.SendMessage(New HandleRef(Me, Me.Handle), NativeMethods.GetParaFormatMessage, IntPtr.Zero, param)
        param.dwMask = NativeMethods.NumberingValid Or NativeMethods.NumberingStyleValid Or NativeMethods.NumberingStart
        param.wNumbering = NativeMethods.NumberingBullet
        param.wNumberingStyle = &H100&
        param.wNumberingStart = 1
        NativeMethods.SendMessage(New HandleRef(Me, Me.Handle), NativeMethods.SetParaFormatMessage, IntPtr.Zero, param)
    End Sub


    ''' <summary>
    ''' Refreshes the preview window contents replacing all instances selected with a reference to provided resource
    ''' </summary>
    ''' <param name="resourceFile"></param>
    ''' <param name="resourceName"></param>
    ''' <remarks></remarks>
    Public Sub RefreshPreview(ByVal resourceFile As Common.ResourceFile, ByVal resourceName As String)
        If Me.ActiveDocument IsNot Nothing Then
            Me._lastShownInstance = Nothing
            Me.Text = Common.ExtensibilityMethods.GetDocumentText(Me.ActiveDocument)
            If Me._activeChangeList IsNot Nothing Then
                Me._previewEntries.Clear()
                Me._instanceInformation.Clear()
                For Each instance As Common.ExtractToResourceActionSite In Me._activeChangeList
                    If instance.StringToExtract.Parent Is Me.ActiveDocument.Parent.ProjectItem Then
                        Me.AddEntryToPreviewList(instance)
                        Dim info As InstanceInformation = New InstanceInformation
                        info.PreviewLength = instance.PreviewChanges(resourceFile, resourceName).Length + 1
                        Me._instanceInformation.Add(instance, info)
                    End If
                Next
                For Each instance As Common.ExtractToResourceActionSite In Me._activeChangeList
                    If instance.StringToExtract.Parent Is Me.ActiveDocument.Parent.ProjectItem Then
                        ' Get starting index on the line
                        Dim startIndex As Integer = instance.StringToExtract.AbsoluteStartIndex
                        Dim length As Integer = instance.StringToExtract.AbsoluteEndIndex - instance.StringToExtract.AbsoluteStartIndex
                        For Each item As Common.ExtractToResourceActionSite In Me._previewEntries.Values
                            If item IsNot instance Then
                                startIndex += Me._instanceInformation(item).PreviewLength
                            Else
                                Exit For
                            End If
                        Next
                        ' Replace the string with its preview
                        Me._instanceInformation(instance).StartingIndex = startIndex
                        Me.Select(startIndex, length)
                        Me.ReplaceSelection(Me.SelectedText, Color.Blue, FontStyle.Strikeout)
                        Me.Select(startIndex + length, 0)
                        Me.ReplaceSelection(My.Resources.Strings.PreviewChangeSeperator, Color.Blue, FontStyle.Regular)
                        Me.Select(startIndex + length + 1, 0)
                        Me.ReplaceSelection(instance.PreviewChanges(resourceFile, resourceName), Color.Red, FontStyle.Bold Or FontStyle.Underline)
                        Me._instanceInformation(instance).EndingIndex = Me.SelectionStart + Me.SelectedText.Length
                    End If
                Next
            End If
        End If

    End Sub

    ''' <summary>
    ''' Highlights and sets the focus on the given instance of text
    ''' </summary>
    ''' <param name="instance">Instance to set focus</param>
    ''' <remarks>Will have no effect if changes for the instance is not shown in the text</remarks>
    Public Sub ShowInstance(ByVal instance As Common.ExtractToResourceActionSite)
        If Me.ActiveDocument.Parent.ProjectItem Is instance.StringToExtract.Parent And _
            Me._instanceInformation.ContainsKey(instance) Then
            If Me._lastShownInstance IsNot Nothing Then
                Me.HighlightInstance(Me._lastShownInstance, Me.BackColor)
            End If
            Me.Focus()
            Me.HighlightInstance(instance, Color.Yellow)
            Me._lastShownInstance = instance
        End If
    End Sub

    ''' <summary>
    ''' Highlights the text related to provided instance of string with the given color
    ''' </summary>
    ''' <param name="instance">Instance to highlight</param>
    ''' <param name="color">Color to highlight with</param>
    ''' <remarks></remarks>
    Private Sub HighlightInstance(ByVal instance As Common.ExtractToResourceActionSite, ByVal color As Drawing.Color)
        Dim info As InstanceInformation = Nothing
        If Me._instanceInformation.ContainsKey(instance) Then
            info = Me._instanceInformation(instance)
            Me.Select(info.StartingIndex, info.EndingIndex - info.StartingIndex)
            Me.SelectionBackColor = color
        End If
    End Sub

    ''' <summary>
    ''' Adds an entry to preview entries dictionary
    ''' </summary>
    ''' <param name="instance">Instance to add to dictionary</param>
    ''' <remarks></remarks>
    Private Sub AddEntryToPreviewList(ByVal instance As Common.ExtractToResourceActionSite)
        Me._previewEntries(instance.StringToExtract.AbsoluteStartIndex) = instance
    End Sub

    ''' <summary>
    ''' Replaces the selection of text with the provided text and style
    ''' </summary>
    ''' <param name="text">New text</param>
    ''' <param name="color">New color of the selection</param>
    ''' <param name="style">New style of the selection</param>
    ''' <remarks></remarks>
    Private Sub ReplaceSelection(ByVal text As String, ByVal color As Color, ByVal style As FontStyle)
        With Me
            .SelectionColor = color
            .SelectionFont = New Font(.Font, style)
            .SelectedText = text
        End With
    End Sub
End Class