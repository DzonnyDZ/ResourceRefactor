''' <summary>Exception raised when a file could not be checked out using Source Control object</summary>
<Serializable()>
Public Class FileCheckoutException
    Inherits Exception

    Private _fileName As String

    ''' <summary>Gets Name of the file that couldn't be checked out.</summary>
    Public ReadOnly Property FileName() As String
        Get
            Return _fileName
        End Get
    End Property

    ''' <summary>Creates a new instance of <see cref="FileCheckoutException"/></summary>
    ''' <param name="fileName">Name of the file</param>
    Public Sub New(ByVal fileName As String)
        MyBase.New(My.Resources.FileCheckOutWarning)
        Me._fileName = fileName
    End Sub
End Class

''' <summary>Exception raised when a file is read only and cannot be edited</summary>
<Serializable()>
Public Class FileReadOnlyException
    Inherits Exception

    Private _fileName As String

    ''' <summary>Gets name of the file that couldn't be checked out.</summary>
    Public ReadOnly Property FileName() As String
        Get
            Return _fileName
        End Get
    End Property

    ''' <summary>Creates a new instance of <see cref="FileCheckoutException"/></summary>
    ''' <param name="fileName">Name of the file</param>
    Public Sub New(ByVal fileName As String)
        MyBase.New(My.Resources.FileCheckOutWarning)
        Me._fileName = fileName
    End Sub

End Class