' Copyright (c) Microsoft Corporation.  All rights reserved.
Imports System.Windows.Forms

''' <summary>
''' This form collects all the information needed to replace strings with resource references in the code.
''' It contains resource file dropdowns and resource entry selection controls.
''' </summary>
''' <remarks></remarks>
Public Class ResourceReplaceOptions

    ''' <summary>
    ''' This event is raised when user changes their resource selection
    ''' </summary>
    ''' <remarks></remarks>
    Public Event ResourceSelectionChanged As EventHandler(Of ResourceSelectionChangedEventArgs)

    Private _stringInstance As Common.BaseHardCodedString
    Private _resources As Common.ResourceFileCollection
    Private resourceNameRegExp As System.Text.RegularExpressions.Regex = _
        New System.Text.RegularExpressions.Regex(My.Resources.Strings.ResourceNameValidationRegExp, _
        System.Text.RegularExpressions.RegexOptions.Compiled)

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        Me.uxNewResourceName.Enabled = Me.uxCreateNewResource.Checked
    End Sub

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the selected resource file.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SelectedResourceFile() As Common.ResourceFile
        Get
            Return Me.uxResourceFileSelector.SelectedResourceFile
        End Get
        Set(ByVal value As Common.ResourceFile)
            If value IsNot Nothing Then
                Me.uxResourceFileSelector.SelectedResourceFile = value
                Me.SetResourceGridView()
            End If
        End Set
    End Property

    ''' <summary>
    ''' Gets the resource entry name to be used during refactoring
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SelectedResourceName() As String
        Get
            If Me.uxCreateNewResource.Checked Then
                Return Me.uxNewResourceName.Text
            Else
                Return Me.uxResourceView.SelectedResourceName
            End If
        End Get
    End Property

    ''' <summary>
    ''' Gets if user has selected to create the resource entry
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IsCreateResourceChecked() As Boolean
        Get
            Return Me.uxCreateNewResource.Checked
        End Get
    End Property

    ''' <summary>
    ''' Gets or sets the options selected by user in the dialog
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Options() As ResourceReplaceOption
        Get
            If Me.uxOptionChangeAll.Checked Then
                Return ResourceReplaceOption.ReplaceAllInstances
            ElseIf Me.uxOptionChangeFileOnly.Checked Then
                Return ResourceReplaceOption.ReplaceCurrentFileOnly
            Else
                Return ResourceReplaceOption.ReplaceCurrentInstance
            End If
        End Get
        Set(ByVal value As ResourceReplaceOption)
            If value = ResourceReplaceOption.ReplaceAllInstances Then
                Me.uxOptionChangeAll.Checked = True
            ElseIf value = ResourceReplaceOption.ReplaceCurrentFileOnly Then
                Me.uxOptionChangeFileOnly.Checked = True
            Else
                Me.uxOptionChangeCurrentOnly.Checked = True
            End If
        End Set
    End Property

#End Region

    ''' <summary>
    ''' Initializes the control to display resource files in the provided collection and show options
    ''' for hard coded string user selected
    ''' </summary>
    ''' <param name="refactorSite">Site object defining current string to refactor and action to perform</param>
    ''' <param name="resourceFiles">Resource file collection for the project</param>
    ''' <remarks></remarks>
    Public Sub DisplayMatchOptions(ByVal refactorSite As Common.ExtractToResourceActionSite, ByVal resourceFiles As Common.ResourceFileCollection)
        If refactorSite Is Nothing Then
            Throw New ArgumentNullException("refactorSite")
        End If
        Me._resources = resourceFiles
        Me._stringInstance = refactorSite.StringToExtract
        Me.uxResourceFileSelector.ExtractResourceAction = refactorSite.ActionObject
        Me.uxResourceFileSelector.FillResourceFilesDropDown(Me._resources)
        If refactorSite.StringToExtract IsNot Nothing Then
            Me.uxNewResourceName.Text = CreateSuggestedResourceName(refactorSite.StringToExtract.Value)
            Me.uxTextToReplace.Text = refactorSite.StringToExtract.Value
        End If
        Me.SetResourceGridView()
    End Sub

    ''' <summary>
    ''' Returns a suggested resource name for the given value of the resource
    ''' </summary>
    ''' <param name="stringValue">Value of the resource to generate a name for</param>
    ''' <returns>a String that can be used as a resource name</returns>
    ''' <remarks></remarks>
    Private Shared Function CreateSuggestedResourceName(ByVal stringValue As String) As String
        Dim sb As New System.Text.StringBuilder(stringValue.Length)
        Dim firstLetterOfWord As Boolean = True
        Dim firstLetter As Boolean = True
        For Each c As Char In stringValue
            If firstLetter And Not Char.IsLetter(c) Then
                Continue For
            End If
            firstLetter = False
            If Char.IsLetterOrDigit(c) Then
                If firstLetterOfWord Then
                    If sb.Length > My.Settings.SuggestedResourceNameMaxLength Then
                        Exit For
                    End If
                    sb.Append(Char.ToUpper(c, Threading.Thread.CurrentThread.CurrentUICulture))
                Else
                    sb.Append(c)
                End If
                firstLetterOfWord = False
            Else
                firstLetterOfWord = True
            End If
        Next
        Return sb.ToString()
    End Function

    Private Sub OnResourceFileSelectionChange(ByVal sender As Object, ByVal e As EventArgs) Handles uxResourceFileSelector.SelectionChanged
        Me.SetResourceGridView()
    End Sub

    ''' <summary>
    ''' Sets the resource grid view according to resource file selection and filtering option selected
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetResourceGridView()
        Me.uxResourceView.BindData(uxResourceFileSelector.SelectedResourceFile, Me._stringInstance)
        Me.uxCreateNewResource.Checked = _
            (Me.uxResourceView.Rows.Count = 0) OrElse _
            (Not Me.uxResourceView.ExactMatchSelected)
        Me.ValidateChildren()
    End Sub

    Private Sub uxCreateNewResource_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles uxCreateNewResource.CheckedChanged
        Me.uxNewResourceName.Enabled = Me.uxCreateNewResource.Checked
        If (Me.uxCreateNewResource.Checked) Then
            Me.ValidateChildren()
        Else
            Me.uxErrorProvider.Clear()
        End If
        Me.RaiseSelectionChangedEvent()
    End Sub

    Private Sub ResourceView_SelectionChanges(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles uxResourceView.SelectionChanged
        Me.RaiseSelectionChangedEvent()
    End Sub

    ''' <summary>
    ''' Raises SelectionChanged Event with correct arguments
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub RaiseSelectionChangedEvent()
        If Me.uxResourceFileSelector.SelectedResourceFile IsNot Nothing Then
            Dim name As String
            If Me.uxCreateNewResource.Checked Then
                name = Me.uxNewResourceName.Text
            Else
                name = Me.uxResourceView.SelectedResourceName
            End If
            If name IsNot Nothing Then
                RaiseEvent ResourceSelectionChanged(Me, New ResourceSelectionChangedEventArgs(Me.uxResourceFileSelector.SelectedResourceFile, name))
            End If
        End If
    End Sub

    Private Sub uxNewResourceName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles uxNewResourceName.TextChanged
        If Me.uxCreateNewResource.Checked Then
            Me.RaiseSelectionChangedEvent()
        End If
    End Sub

    ''' <summary>
    ''' Validates the resource name using regular expressions.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub uxNewResourceName_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles uxNewResourceName.Validating
        If Me.uxCreateNewResource.Checked Then
            If (Not Me.resourceNameRegExp.IsMatch(Me.uxNewResourceName.Text)) Then
                e.Cancel = True
                Me.uxErrorProvider.SetError(Me.uxNewResourceName, My.Resources.Strings.InvalidResourceNameError)
                Exit Sub
            ElseIf Me.SelectedResourceFile IsNot Nothing AndAlso Me.SelectedResourceFile.Contains(Me.uxNewResourceName.Text) Then
                e.Cancel = True
                Me.uxErrorProvider.SetError(Me.uxNewResourceName, My.Resources.Strings.ResourceAlreadyExists)
                Exit Sub
            End If
        End If
        e.Cancel = False
        Me.uxErrorProvider.SetError(Me.uxNewResourceName, String.Empty)
    End Sub
End Class

Public Class ResourceSelectionChangedEventArgs
    Inherits EventArgs

    Private _file As Common.ResourceFile
    Private _name As String

    ''' <summary>
    ''' Resource file user has selected
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property ResourceFile() As Common.ResourceFile
        Get
            Return _file
        End Get
    End Property

    ''' <summary>
    ''' Resource name user has selected
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property ResourceName() As String
        Get
            Return _name
        End Get
    End Property

    ''' <summary>
    ''' Creates a new argument object
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal resourceFile As Common.ResourceFile, ByVal resourceName As String)
        MyBase.New()
        Me._file = resourceFile
        Me._name = resourceName
    End Sub

End Class

''' <summary>
''' Enum to define options for resource refactoring
''' </summary>
''' <remarks></remarks>
Public Enum ResourceReplaceOption
    ReplaceCurrentInstance
    ReplaceCurrentFileOnly
    ReplaceAllInstances
End Enum