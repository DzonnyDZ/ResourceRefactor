Imports System
Imports System.Collections.Generic
Imports System.Text
Imports EnvDTE

''' <summary>Specifies project types to be used by implementations of <see cref="T:Microsoft.VSPowerToys.ResourceRefactor.Common.BaseHardCodedString"/></summary>
Public Enum ProjectType
    ''' <summary>C# project</summary>
    CSharp = 1
    ''' <summary>Visual Basic project</summary>
    VB = 2
    ''' <summary>Web project</summary>
    WebProject = 3
    ''' <summary>No project type selected</summary>
    None = 0
End Enum

''' <summary>Contains common methods related to <see cref="EnvDTE"/></summary>
Public NotInheritable Class ExtensibilityMethods
    ''' <summary>Hides the default constructor and makes "static" class</summary>
    Private Sub New()
    End Sub

    ''' <summary>Check outs a project item if it is under source control and not checked out.</summary>
    ''' <param name="item">An item to be checked out</param>
    ''' <exception cref="ArgumentNullException"><paramref name="item"/> is null</exception>
    Public Shared Function CheckoutItem(ByVal item As ProjectItem) As Boolean
        If item Is Nothing Then
            Throw New ArgumentNullException("item")
        End If
        Dim result As Boolean = True
        Dim control As SourceControl = item.DTE.SourceControl
        If (control IsNot Nothing) Then
            Dim itemName As String = item.FileNames(1)
            If (control.IsItemUnderSCC(itemName) AndAlso Not control.IsItemCheckedOut(itemName)) Then
                Try
                    result = control.CheckOutItem(itemName)
                Finally
                    result = control.IsItemCheckedOut(itemName)
                End Try
            End If
        End If
        ' If not under source return true
        Return result
    End Function

    ''' <summary>Gets the project type of a project</summary>
    ''' <param name="project">Project to query for its type</param>
    ''' <returns>Type of project</returns>
    ''' <exception cref="ArgumentNullException"><paramref name="project"/> is null</exception>
    Public Shared Function GetProjectType(ByVal project As Project) As ProjectType
        If project Is Nothing Then
            Throw New ArgumentNullException("project")
        End If
        Dim kind As String = project.Kind
        If (kind.Equals(My.Resources.ProjectKindCSharp)) Then
            Return ProjectType.CSharp
        ElseIf (kind.Equals(My.Resources.ProjectKindVB)) Then
            Return ProjectType.VB
        ElseIf (kind.Equals(My.Resources.ProjectKindWeb)) Then
            Return ProjectType.WebProject
        Else
            Return ProjectType.None
        End If
    End Function

    ''' <summary>Gets the text for the whole document object</summary>
    ''' <param name="document">A document to get text of</param>
    ''' <returns>text of the document</returns>
    ''' <exception cref="ArgumentNullException"><paramref name="document"/> is null</exception>
    Public Shared Function GetDocumentText(ByVal document As TextDocument) As String
        If document Is Nothing Then
            Throw New ArgumentNullException("document")
        End If
        Dim start As EditPoint = document.StartPoint.CreateEditPoint()
        Return start.GetText(document.EndPoint.AbsoluteCharOffset)
    End Function
End Class
