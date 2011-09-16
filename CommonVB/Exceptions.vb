' Copyright (c) Microsoft Corporation.  All rights reserved.

''' <summary>
''' Exception raised when a file could not be checked out using Source Control object
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class FileCheckoutException
    Inherits Exception

    Private _fileName As String

    ''' <summary>
    ''' Name of the file that couldn't be checked out.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FileName() As String
        Get
            Return _fileName
        End Get
    End Property

    ''' <summary>
    ''' Constructor created for FxCop compatibility
    ''' </summary>
    Public Sub New()
        MyBase.New()
    End Sub

    ''' <summary>
    ''' Creates a new instance of FileCheckoutException
    ''' </summary>
    ''' <param name="fileName">Name of the file</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal fileName As String)
        MyBase.New(My.Resources.FileCheckOutWarning)
        Me._fileName = fileName
    End Sub

    ''' <summary>
    ''' Constructor created for FxCop compatibility
    ''' </summary>
    Public Sub New(ByVal message As String, ByVal inner As Exception)
        MyBase.New(message, inner)
    End Sub

    ''' <summary>
    ''' Constructor created for FxCop compatibility
    ''' </summary>
    Protected Sub New( _
        ByVal info As System.Runtime.Serialization.SerializationInfo, _
        ByVal context As System.Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub

    <System.Security.Permissions.SecurityPermission(Security.Permissions.SecurityAction.Demand, SerializationFormatter:=True)> _
    <System.Security.Permissions.SecurityPermission(Security.Permissions.SecurityAction.LinkDemand, Flags:=Security.Permissions.SecurityPermissionFlag.SerializationFormatter)> _
    Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
        MyBase.GetObjectData(info, context)
    End Sub
End Class

''' <summary>
''' Exception raised when a file is read only and cannot be edited
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class FileReadOnlyException
    Inherits Exception

    Private _fileName As String

    ''' <summary>
    ''' Name of the file that couldn't be checked out.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FileName() As String
        Get
            Return _fileName
        End Get
    End Property

    ''' <summary>
    ''' Constructor created for FxCop compatibility
    ''' </summary>
    Public Sub New()
        MyBase.New()
    End Sub

    ''' <summary>
    ''' Creates a new instance of FileCheckoutException
    ''' </summary>
    ''' <param name="fileName">Name of the file</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal fileName As String)
        MyBase.New(My.Resources.FileCheckOutWarning)
        Me._fileName = fileName
    End Sub

    ''' <summary>
    ''' Constructor created for FxCop compatibility
    ''' </summary>
    Public Sub New(ByVal message As String, ByVal inner As Exception)
        MyBase.New(message, inner)
    End Sub

    ''' <summary>
    ''' Constructor created for FxCop compatibility
    ''' </summary>
    Protected Sub New( _
        ByVal info As System.Runtime.Serialization.SerializationInfo, _
        ByVal context As System.Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub

    <System.Security.Permissions.SecurityPermission(Security.Permissions.SecurityAction.Demand, SerializationFormatter:=True)> _
    <System.Security.Permissions.SecurityPermission(Security.Permissions.SecurityAction.LinkDemand, Flags:=Security.Permissions.SecurityPermissionFlag.SerializationFormatter)> _
    Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
        MyBase.GetObjectData(info, context)
    End Sub
End Class