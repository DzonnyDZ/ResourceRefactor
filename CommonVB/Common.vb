' Copyright (c) Microsoft Corporation.  All rights reserved.

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports EnvDTE

''' <summary>
''' Specifies project types to be used by implementations of IHardCodedString
''' </summary>
Public Enum ProjectType
    CSharp = 1
    VB = 2
    WebProject = 3
    None = 0
End Enum

''' <summary>
''' Contains common methods related to EnvDTE
''' </summary>
Public NotInheritable Class ExtensibilityMethods

    ''' <summary>
    ''' Hides the default constructor
    ''' </summary>
    Private Sub New()

    End Sub

    ''' <summary>
    ''' Check outs a project item if it is under source control and not checked out.
    ''' </summary>
    ''' <param name="item"></param>
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

    ''' <summary>
    ''' Gets the project type of a project
    ''' </summary>
    ''' <param name="project">Project to query for its type</param>
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

    ''' <summary>
    ''' Gets the text for the whole document object
    ''' </summary>
    ''' <param name="document"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetDocumentText(ByVal document As TextDocument) As String
        If document Is Nothing Then
            Throw New ArgumentNullException("document")
        End If
        Dim start As EditPoint = document.StartPoint.CreateEditPoint()
        Return start.GetText(document.EndPoint.AbsoluteCharOffset)
    End Function
End Class
