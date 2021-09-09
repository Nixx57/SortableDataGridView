Imports System.Text

Public Class DataGridComparer
    Implements IComparer

    Private _grid As DataGridView
    Private _sortedColumns As List(Of SortColDefn)

    Private Structure SortColDefn
        Friend colNum As Int16
        Friend order As SortOrder

        Friend Sub New(columnNum As Integer, sortOrder As SortOrder)
            colNum = Convert.ToInt16(columnNum)
            order = sortOrder
        End Sub
    End Structure

    Public Sub New(datagrid As DataGridView)
        _grid = datagrid

        'Si à 0 => Pas de limite de colonnes trié

        _sortedColumns = New List(Of SortColDefn)
    End Sub

    Public Function SetSortColumn(columnIndex As Integer, ModifierKeys As Keys) As SortOrder
        Dim colDefn As SortColDefn

        Dim sortPriority As Integer = _sortedColumns.FindIndex(Function(cd As SortColDefn)
                                                                   Return cd.colNum = columnIndex
                                                               End Function)

        If sortPriority <> -1 Then 'Si existe dans l'array
            Return ReverseSort(sortPriority)
        End If

        colDefn = New SortColDefn(columnIndex, SortOrder.Ascending)
        _grid.Columns(colDefn.colNum).HeaderCell.SortGlyphDirection = SortOrder.Ascending

        _sortedColumns.Add(colDefn)

        _grid.Columns(colDefn.colNum).HeaderCell.SortGlyphDirection = SortOrder.Ascending

        Return SortOrder.Ascending

    End Function

    Private Function ReverseSort(sortPriority As Integer) As SortOrder
        Dim colDefn As SortColDefn = _sortedColumns(sortPriority)

        Dim sortOrder As SortOrder

        If colDefn.order = SortOrder.Ascending Then
            colDefn.order = SortOrder.Descending
            _grid.Columns(colDefn.colNum).HeaderCell.SortGlyphDirection = SortOrder.Descending
        ElseIf colDefn.order = SortOrder.Descending Then
            colDefn.order = SortOrder.None
            _grid.Columns(colDefn.colNum).HeaderCell.SortGlyphDirection = SortOrder.None
            _sortedColumns.RemoveAt(sortPriority)
        ElseIf colDefn.order = SortOrder.None Then
            colDefn.order = SortOrder.Ascending
            _grid.Columns(colDefn.colNum).HeaderCell.SortGlyphDirection = SortOrder.Ascending
        End If
        If colDefn.order <> SortOrder.None Then
            _sortedColumns(sortPriority) = colDefn
        End If


        sortOrder = colDefn.order
        If sortPriority <> 0 Then
            Return sortOrder
        End If

        _grid.Columns(colDefn.colNum).HeaderCell.SortGlyphDirection = sortOrder

        Return sortOrder

    End Function

    Public ReadOnly Property SortOrderDescription
        Get
            Dim sb As StringBuilder = New StringBuilder("Trié par : ")

            For Each item As SortColDefn In _sortedColumns
                sb.Append(_grid.Columns(item.colNum).HeaderText)
                If item.order = SortOrder.Ascending Then
                    sb.Append(" ASC | ")
                ElseIf item.order = SortOrder.Descending Then
                    sb.Append(" DESC | ")
                ElseIf item.order = SortOrder.None Then
                    sb.Append(" NONE | ")
                End If
            Next

            sb.Length -= 2

            Return sb.ToString()
        End Get
    End Property

    Public Function Compare(lhs As DataGridViewCellCollection, rhs As DataGridViewCellCollection) As Integer
        For Each item As SortColDefn In _sortedColumns
            If item.order <> SortOrder.None Then

                Dim retval As Integer = Comparer(Of Object).Default.Compare(
                lhs(item.colNum).Value,
                rhs(item.colNum).Value)

                If retval <> 0 Then
                    Return IIf(item.order = SortOrder.Ascending, retval, -retval)
                End If
            End If
        Next
        Return 0
    End Function

    Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
        Dim lhs As DataGridViewRow = x
        Dim rhs As DataGridViewRow = y

        Return Compare(lhs.Cells, rhs.Cells)
    End Function
End Class
