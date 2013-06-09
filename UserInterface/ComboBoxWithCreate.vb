' Copyright (c) Microsoft Corporation.  All rights reserved.
Imports System.Windows.Forms

''' <summary>A combobox with a custom "Create Item" entry. </summary>
Public Class ComboBoxWithCreate
    Inherits ComboBox

    ''' <summary>Raised when Create Item selection is selected</summary>
    Public Event CreateItemSelected As EventHandler(Of EventArgs)

    ''' <summary>Raises the <see cref="E:System.Windows.Forms.ComboBox.SelectedIndexChanged" /> event.</summary>
    ''' <param name="e">An <see cref="T:System.EventArgs" /> that contains the event data. </param>
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

''' <summary>Item representing "Create item" entry. This must be manually added to items collection by user for it to be selectable.</summary>
Public Class CreateItem

    ''' <summary>Displayed text for the entry</summary>
    Private _text As String

    ''' <summary>Creates a new instance of the item</summary>
    ''' <param name="text">Text to be displayed in the combo box</param>
    Public Sub New(ByVal text As String)
        _text = text
    End Sub

    ''' <summary>Returns a string that represents the current object.</summary>
    ''' <returns>A string that represents the current object.</returns>
    ''' <filterpriority>2</filterpriority>
    Public Overrides Function ToString() As String
        Return _text
    End Function
End Class