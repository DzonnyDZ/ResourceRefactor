' Copyright (c) Microsoft Corporation.  All rights reserved.
Imports System.Windows.Forms
Imports System.Resources

''' <summary>
''' An extension to grid view for displayin resource entries
''' </summary>
''' <remarks></remarks>
Public Class ResourceGridView
    Inherits System.Windows.Forms.DataGridView

    ''' <summary>
    ''' Returns the selected resource entry's name
    ''' </summary>
    ''' <value></value>
    ''' <returns>Returns a non empty string for name, or Nothing if no row is selected</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SelectedResourceName() As String
        Get
            If Me.SelectedRows.Count > 0 Then
                Dim currentRow As System.Data.DataRowView = CType(Me.SelectedRows.Item(0).DataBoundItem, System.Data.DataRowView)
                Return CType(currentRow("Object"), Common.ResourceMatch).ResourceName
            End If
            Return Nothing
        End Get
    End Property

    ''' <summary>
    ''' Gets if currently selected row is an exact match to the text being refactored
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ExactMatchSelected() As Boolean
        Get
            Dim result As Boolean = False
            If Me.SelectedRows.Count > 0 Then
                Dim currentRow As System.Data.DataRowView = CType(Me.SelectedRows.Item(0).DataBoundItem, System.Data.DataRowView)
                result = currentRow("Percentage").Equals(Common.StringMatch.ExactMatch)
            End If
            Return result
        End Get
    End Property

    ''' <summary>
    ''' Initializes the default settings for the grid view
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub InitLayout()
        MyBase.InitLayout()
        Me.AllowUserToAddRows = False
        Me.AllowUserToDeleteRows = False
        Me.AllowUserToResizeRows = False
        Me.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.MultiSelect = False
        Me.ReadOnly = True
        Me.AutoSize = True
        Me.AutoGenerateColumns = False
    End Sub

    ''' <summary>
    ''' Binds data from resource file to grid view
    ''' </summary>
    ''' <param name="resources">Resource file to use for getting data</param>
    ''' <param name="stringInstance">If not Nothing, resource list is filtered by this hard coded string to list only similar ones</param>
    ''' <remarks></remarks>
    Public Sub BindData(ByVal resources As Common.ResourceFile, ByVal stringInstance As Common.BaseHardCodedString)
        If resources IsNot Nothing Then
            Me.DataSource = Nothing
            Me.Rows.Clear()
            Me.Columns.Clear()
            Me.Columns.Add("Name", "Name")
            Me.Columns.Add("Value", "Value")
            Me.Columns.Add("Match", "Similarity")
            With Me.Columns.Item(0)
                .DataPropertyName = "Name"
                .SortMode = DataGridViewColumnSortMode.Automatic
            End With
            With Me.Columns.Item(1)
                .DataPropertyName = "Value"
                .SortMode = DataGridViewColumnSortMode.Automatic
            End With
            With Me.Columns.Item(2)
                .DataPropertyName = "Percentage"
                .SortMode = DataGridViewColumnSortMode.Automatic
            End With
            Dim textToMatch As String = String.Empty
            If stringInstance IsNot Nothing Then
                textToMatch = stringInstance.Value
            End If
            Dim table As System.Data.DataTable = New System.Data.DataTable()
            table.Locale = System.Globalization.CultureInfo.CurrentCulture
            table.Columns.Add("Name")
            table.Columns.Add("Value")
            table.Columns.Add("Percentage", New Double().GetType())
            table.Columns.Add("Object", New Common.ResourceMatch(0, Nothing, Nothing).GetType())
            For Each match As Common.ResourceMatch In resources.GetAllMatches(textToMatch)
                table.Rows.Add(match.ResourceName, match.Value, match.Percentage, match)
            Next
            Me.DataSource = table
            Me.Sort(Me.Columns.Item(2), System.ComponentModel.ListSortDirection.Descending)
        End If
    End Sub

    ''' <summary>
    ''' Formats the similarity column and determines what to display based on similarity percentage
    ''' </summary>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub OnCellFormatting(ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs)
        If e.ColumnIndex = 2 Then
            Dim matchPercentage As Double = CType(e.Value, Double)
            If matchPercentage = Common.StringMatch.ExactMatch Then
                e.Value = "Same"
            ElseIf matchPercentage >= 0.8 Then
                e.Value = "Very Similar"
            ElseIf matchPercentage >= 0.6 Then
                e.Value = "Similar"
            Else
                e.Value = "Not Similar"
            End If
        End If
        MyBase.OnCellFormatting(e)
    End Sub

    ''' <summary>
    ''' Checks if a row added is an exact match and if so selects the row
    ''' </summary>
    ''' <param name="e"></param>
    ''' <remarks>CA1817: Suppressed because currentRowView changes in each iteration.</remarks>
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1817:DoNotCallPropertiesThatCloneValuesInLoops")> _
    Protected Overrides Sub OnRowsAdded(ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs)
        For i As Integer = 0 To e.RowCount - 1
            Dim currentRowView As System.Windows.Forms.DataGridViewRow = Me.Rows(e.RowIndex + i)
            Dim currentRow As System.Data.DataRowView = CType(currentRowView.DataBoundItem, System.Data.DataRowView)
            If currentRow("Percentage").Equals(Common.StringMatch.ExactMatch) Then
                currentRowView.Selected = True
                Return
            End If
        Next
        MyBase.OnRowsAdded(e)
    End Sub
End Class
