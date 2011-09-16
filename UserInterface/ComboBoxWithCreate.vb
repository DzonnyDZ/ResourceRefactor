' Copyright (c) Microsoft Corporation.  All rights reserved.
Imports System.Windows.Forms

''' <summary>
''' A combobox with a custom "Create Item" entry. 
''' </summary>
''' <remarks></remarks>
Public Class ComboBoxWithCreate
    Inherits ComboBox

    ''' <summary>
    ''' Raised when Create Item selection is selected
    ''' </summary>
    ''' <remarks></remarks>
    Public Event CreateItemSelected As EventHandler(Of EventArgs)

    Protected Overrides Sub OnSelectedIndexChanged(ByVal e As System.EventArgs)
        If TypeOf Me.SelectedItem Is CreateItem Then
            'Do not allow the "Create" item to be selected
            Me.SelectedIndex = -1
            RaiseEvent CreateItemSelected(Me, Nothing)
        Else
            MyBase.OnSelectedIndexChanged(e)
        End If
    End Sub

End Class

''' <summary>
''' Item representing "Create item" entry. This must be manually added
''' to items collection by user for it to be selectable.
''' </summary>
''' <remarks></remarks>
Public Class CreateItem

    ''' <summary>
    ''' Displayed text for the entry
    ''' </summary>
    ''' <remarks></remarks>
    Private _text As String

    ''' <summary>
    ''' Creates a new instance of the itme
    ''' </summary>
    ''' <param name="text">Text to be displayed in the combo box</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal text As String)
        _text = text
    End Sub

    Public Overrides Function ToString() As String
        Return _text
    End Function
End Class