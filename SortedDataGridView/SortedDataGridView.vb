Imports System.ComponentModel

Public Class SortedDataGridView
    Inherits DataGridView

    Public _columnSorter As DataGridComparer

    Public Sub New()
        _columnSorter = New DataGridComparer(Me)
    End Sub

    Protected Overrides Sub OnColumnHeaderMouseClick(e As DataGridViewCellMouseEventArgs)
        If e.Button = MouseButtons.Left Then
            _columnSorter.SetSortColumn(e.ColumnIndex, ModifierKeys)

            Sort(_columnSorter)

            Columns(e.ColumnIndex).SortMode = DataGridViewColumnSortMode.Programmatic
        End If
        MyBase.OnColumnHeaderMouseClick(e)

    End Sub

    Public ReadOnly Property SortOrderDescription
        Get
            Return _columnSorter.SortOrderDescription
        End Get
    End Property
End Class