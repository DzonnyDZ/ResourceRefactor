' Copyright (c) Microsoft Corporation.  All rights reserved.

''' <summary>
''' Contains method to match strings approximatlely
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class StringMatch

    'A match is rated from 0 (not at all) to 1.0 (exact)
    Public Const NoMatch As Double = 0.0
    Public Const ExactMatch As Double = 1.0

    'Private-use constants - external clients should simply sort the values and not check against these constants
    Private Const ExactMatchDiffersByCase As Double = 0.95

    ''' <summary>
    ''' Private constructor to avoid any public constructors
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub New()

    End Sub

    ''' <summary>
    ''' Calculates a match value for two strings (using Levinsthein algorithm)
    ''' </summary>
    ''' <param name="original">Original string to compare against</param>
    ''' <param name="other">String to compare</param>
    ''' <returns>A double indicating match result (1.0 being an exact match)</returns>
    ''' <remarks>Arguments are already validated by the first if statement, thus message is suppressed</remarks>
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062:ValidateArgumentsOfPublicMethods")> _
    Public Shared Function Match(ByVal original As String, ByVal other As String) As Double
        If String.IsNullOrEmpty(original) Or String.IsNullOrEmpty(other) Then
            'If either string is empty, we don't want to consider it a match for our purposes
            Return NoMatch
        End If

        'Try exact match
        If original.Equals(other, StringComparison.CurrentCulture) Then
            Return ExactMatch
        End If

        'Try case-independent
        If original.Equals(other, StringComparison.CurrentCultureIgnoreCase) Then
            Return ExactMatchDiffersByCase
        End If
        Dim maxLength As Double
        If original.Length < other.Length Then
            maxLength = other.Length
        Else
            maxLength = original.Length
        End If
        Dim distance As Double = CalculateDistance(original, other)
        Return 1 - (distance / maxLength)
    End Function

    ''' <summary>
    ''' Returns the Levenshtein difference between 2 strings
    ''' </summary>
    ''' <param name="original">String to start with</param>
    ''' <param name="other">String trying to end up with</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CalculateDistance(ByVal original As String, ByVal other As String) As Integer
        If original Is Nothing Then
            Throw New ArgumentNullException("original")
        End If
        If other Is Nothing Then
            Throw New ArgumentNullException("other")
        End If

        Dim upperRow(other.Length) As Integer
        Dim lowerRow(other.Length) As Integer
        Dim originalChars() As Char = original.ToCharArray()
        Dim otherChars() As Char = other.ToCharArray()

        'Initialize array
        For i As Integer = 0 To other.Length
            upperRow(i) = i
        Next
        lowerRow(0) = 1
        For i As Integer = 1 To original.Length
            For j As Integer = 1 To other.Length
                Dim insert As Integer = 2
                If originalChars(i - 1) = otherChars(j - 1) Then
                    insert = 0
                ElseIf Char.ToLowerInvariant(originalChars(i - 1)) = Char.ToLowerInvariant(otherChars(j - 1)) Then
                    insert = 1
                End If
                lowerRow(j) = Min(upperRow(j) + 2, lowerRow(j - 1) + 2, upperRow(j - 1) + insert)
            Next
            upperRow = lowerRow
            ReDim lowerRow(other.Length)
            lowerRow(0) = i + 1
        Next
        Return CType(upperRow(other.Length) / 2, Integer)
    End Function

    ''' <summary>
    ''' Returns the minimum of 3 numbers
    ''' </summary>
    ''' <param name="d1"></param>
    ''' <param name="d2"></param>
    ''' <param name="d3"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function Min(ByVal d1 As Integer, ByVal d2 As Integer, ByVal d3 As Integer) As Integer
        If d1 <= d2 And d1 <= d3 Then
            Return d1
        ElseIf d2 <= d1 And d2 <= d3 Then
            Return d2
        Else
            Return d3
        End If
    End Function

End Class
