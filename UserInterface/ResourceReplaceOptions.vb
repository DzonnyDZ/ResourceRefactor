Imports System.Windows.Forms
Imports System.Text.RegularExpressions

''' <summary>This form collects all the information needed to replace strings with resource references in the code.</summary>
''' <remarks>It contains resource file dropdowns and resource entry selection controls.</remarks>
Public Class ResourceReplaceOptions

    ''' <summary>This event is raised when user changes their resource selection</summary>
    Public Event ResourceSelectionChanged As EventHandler(Of ResourceSelectionChangedEventArgs)

    Private _stringInstance As Common.BaseHardCodedString
    Private _resources As Common.ResourceFileCollection
    Private resourceNameRegExp As Regex = New System.Text.RegularExpressions.Regex(My.Resources.Strings.ResourceNameValidationRegExp, RegexOptions.Compiled)

    ''' <summary>CTor - creates a new instance of the <see cref="ResourceReplaceOptions"/> class</summary>
    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        Me.uxNewResourceName.Enabled = Me.uxCreateNewResource.Checked
    End Sub

#Region "Public Properties"

    ''' <summary>Gets or sets the selected resource file.</summary>
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

    ''' <summary>Gets the resource entry name to be used during refactoring</summary>
    Public ReadOnly Property SelectedResourceName() As String
        Get
            If Me.uxCreateNewResource.Checked Then
                Return Me.uxNewResourceName.Text
            Else
                Return Me.uxResourceView.SelectedResourceName
            End If
        End Get
    End Property

    ''' <summary>Gets if user has selected to create the resource entry</summary>
    Public ReadOnly Property IsCreateResourceChecked() As Boolean
        Get
            Return Me.uxCreateNewResource.Checked
        End Get
    End Property

    ''' <summary>Gets or sets the options selected by user in the dialog</summary>
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

    ''' <summary>Initializes the control to display resource files in the provided collection and show options for hard coded string user selected</summary>
    ''' <param name="refactorSite">Site object defining current string to refactor and action to perform</param>
    ''' <param name="resourceFiles">Resource file collection for the project</param>
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

    ''' <summary>Returns a suggested resource name for the given value of the resource</summary>
    ''' <param name="stringValue">Value of the resource to generate a name for</param>
    ''' <returns>a String that can be used as a resource name</returns>
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

    Private Sub uxResourceFileSelector_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles uxResourceFileSelector.SelectionChanged
        Me.SetResourceGridView()
    End Sub

    ''' <summary>Sets the resource grid view according to resource file selection and filtering option selected</summary>
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

    ''' <summary>Raises <see cref="ResourceSelectionChanged"/> event with correct arguments</summary>
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

    ''' <summary>Validates the resource name using regular expressions.</summary>
    ''' <param name="sender">Source of the event</param>
    ''' <param name="e">Event arguments</param>
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

''' <summary>Event arguments of an event raised when resource selection changes</summary>
Public Class ResourceSelectionChangedEventArgs
    Inherits EventArgs

    Private _file As Common.ResourceFile
    Private _name As String

    ''' <summary>Gets resource file user has selected</summary>
    Public ReadOnly Property ResourceFile() As Common.ResourceFile
        Get
            Return _file
        End Get
    End Property

    ''' <summary>Gets resource name user has selected</summary>
    Public ReadOnly Property ResourceName() As String
        Get
            Return _name
        End Get
    End Property

    ''' <summary>Gets creates a new argument object</summary>
    Public Sub New(ByVal resourceFile As Common.ResourceFile, ByVal resourceName As String)
        MyBase.New()
        Me._file = resourceFile
        Me._name = resourceName
    End Sub

End Class

''' <summary>Enum to define options for resource refactoring</summary>
Public Enum ResourceReplaceOption
    ''' <summary>Replace olny surrent instance</summary>
    ReplaceCurrentInstance
    ''' <summary>Replace all occurences in current file</summary>
    ReplaceCurrentFileOnly
    ''' <summary>Replaces all instances in current project</summary>
    ReplaceAllInstances
End Enum