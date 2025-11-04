Imports System.ComponentModel
Imports System.Collections.Generic
Public Class DGV_MulCapRowMerge
    Inherits DataGridView
    Private _columnCaptions As String()
    Private _ColCap As String(,)
    Private _cellHeight As Integer = 35
    Private _columnDeep As Integer = 1
    Private _HeaderFont As Font = Me.Font
    Private _mergecolumnname As List(Of String) = New List(Of String)()


    ''' <summary>
    ''' 设置或获得数据列合并的集合
    ''' </summary>
    ''' <returns></returns>
    Public Property MergeColumnNames() As List(Of String)
        Get
            Return _mergecolumnname
        End Get
        Set(ByVal Value As List(Of String))
            _mergecolumnname = Value
        End Set
    End Property

    ''' <summary>
    ''' 设置或获得标头的字体
    ''' </summary>
    ''' <returns></returns>
    Public Property ColHeaderFont() As Font
        Get
            Return _HeaderFont
        End Get
        Set(ByVal value As Font)
            _HeaderFont = value
        End Set
    End Property

    ''' <summary>
    ''' 设置或获得合并表头的深度
    ''' </summary>
    ''' <returns></returns>
    Public Property ColumnDeep() As Integer
        Get
            If Me.Columns.Count = 0 Then
                _columnDeep = 1
            End If
            Me.ColumnHeadersHeight = _cellHeight * _columnDeep

            Return _columnDeep
        End Get
        Set(ByVal value As Integer)
            If value < 1 Then
                _columnDeep = 1
            Else
                _columnDeep = value
            End If
            Me.ColumnHeadersHeight = _cellHeight * _columnDeep
        End Set
    End Property

    ''' <summary>
    ''' 添加合并式单元格绘制的所需要的对象
    ''' </summary>
    ''' <returns></returns>
    Public Property ColumnSCaption() As String()
        Get
            Return _columnCaptions
        End Get
        Set(ByVal value As String())
            _columnCaptions = value
            GetCaption()
        End Set
    End Property
    Private Sub GetCaption()
        Dim iRow As Integer = _columnCaptions.GetLength(0)
        If iRow > 0 Then '表示设置OK
            On Error Resume Next
            Dim strCap As String() = _columnCaptions(0).Split(",")
            Dim iCol As Integer = strCap.GetLength(0) - 1
            '     Me.Columns.Clear()
            ReDim Preserve _ColCap(iRow, iCol)
            For iR As Integer = 0 To iRow - 1
                For iC As Integer = 0 To iCol
                    _ColCap(iR, iC) = strCap(iC)
                Next
                If iR + 1 < iRow Then
                    strCap = _columnCaptions(iR + 1).Split(",")
                End If
            Next
        End If
    End Sub
    Protected Overrides Sub OnCellPainting(ByVal e As DataGridViewCellPaintingEventArgs)
        If e.RowIndex > -1 And e.ColumnIndex > -1 Then  '数据区重新回执
            DrawCell(e)
        ElseIf e.RowIndex = -1 And e.ColumnIndex > -1 Then '列头重新绘制
            If _columnDeep = 1 Then
                MyBase.OnCellPainting(e)
                '  e.Handled = True
                Exit Sub
            End If
            Dim ILev As Integer = ColumnDeep
            Do While ILev > 0
                PaintUnitHeader(e, ILev)
                ILev = ILev - 1
            Loop
            e.Handled = True
        ElseIf e.ColumnIndex = -1 Then
            MyBase.OnCellPainting(e)
            Exit Sub
        End If
    End Sub


    Private Sub DrawCell(ByVal e As DataGridViewCellPaintingEventArgs)
        If e.CellStyle.Alignment = DataGridViewContentAlignment.NotSet Then
            e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End If
        Dim gridBrush As Brush = New SolidBrush(Me.GridColor)
        Dim backBrush As SolidBrush = New SolidBrush(Color.White)
        Dim fontBrush As SolidBrush = New SolidBrush(e.CellStyle.ForeColor)
        Dim cellwidth As Integer
        Dim UpRows As Integer = 0
        Dim DownRows As Integer = 0
        Dim count As Integer = 0
        Try
            If Me.MergeColumnNames.Contains(Me.Columns(e.ColumnIndex).Name) And e.RowIndex <> -1 Then '如果本列需要合并
                cellwidth = e.CellBounds.Width
                Dim gridLinePen As Pen = New Pen(gridBrush)
                Dim curValue As String = CType(e.Value, String)
                If curValue Is Nothing Then Exit Sub

                Dim curSelected As String = CType(Me.CurrentRow.Cells(e.ColumnIndex).Value, String)
                If curSelected Is Nothing Then Exit Sub

                For i As Integer = e.RowIndex To Me.Rows.Count - 1 Step 1
                    If Me.Rows(i).Cells(e.ColumnIndex).Value.ToString().Equals(curValue) Then
                        If e.ColumnIndex > 0 Then
                            If Me.MergeColumnNames.Contains(Me.Columns(e.ColumnIndex - 1).Name) And i >= 1 Then '如果前一列合并
                                If Not Me.Rows(i).Cells(e.ColumnIndex - 1).Value.ToString().Equals(Me.Rows(i - 1).Cells(e.ColumnIndex - 1).Value.ToString()) Then Exit For
                            End If
                        End If
                        DownRows = DownRows + 1
                    Else
                        Exit For
                    End If
                Next

                For j As Integer = e.RowIndex To 0 Step -1
                    If Me.Rows(j).Cells(e.ColumnIndex).Value.ToString().Equals(curValue) Then
                        If e.ColumnIndex > 0 Then
                            If Me.MergeColumnNames.Contains(Me.Columns(e.ColumnIndex - 1).Name) And j < e.RowIndex Then '如果前一列合并
                                If Not Me.Rows(j).Cells(e.ColumnIndex - 1).Value.ToString().Equals(Me.Rows(j + 1).Cells(e.ColumnIndex - 1).Value.ToString()) Then Exit For
                            End If
                        End If
                        UpRows = UpRows + 1
                    Else
                        Exit For
                    End If
                Next
                count = DownRows + UpRows - 1
                If count < 2 Then Return

                If Me.Rows(e.RowIndex).Selected Then
                    backBrush.Color = e.CellStyle.SelectionBackColor
                    fontBrush.Color = e.CellStyle.SelectionForeColor
                End If
                '填充清空
                e.Graphics.FillRectangle(backBrush, e.CellBounds.X, e.CellBounds.Y - e.CellBounds.Height * (UpRows - 1), cellwidth, e.CellBounds.Height * count)
                '底线
                e.Graphics.DrawLine(gridLinePen, e.CellBounds.Left, e.CellBounds.Bottom + (DownRows - 1) * e.CellBounds.Height - 1, e.CellBounds.Right - 1, e.CellBounds.Bottom + (DownRows - 1) * e.CellBounds.Height - 1)
                '右线
                e.Graphics.DrawLine(gridLinePen, e.CellBounds.Right - 1, e.CellBounds.Top - e.CellBounds.Height * (UpRows - 1), e.CellBounds.Right - 1, e.CellBounds.Bottom + e.CellBounds.Height * (DownRows - 1))
                '写内容
                Dim fontheight As Integer = CType(e.Graphics.MeasureString(e.Value.ToString(), e.CellStyle.Font).Height, Integer)
                Dim fontwidth As Integer = CType(e.Graphics.MeasureString(e.Value.ToString(), e.CellStyle.Font).Width, Integer)
                Dim cellheight As Integer = e.CellBounds.Height
                Select Case e.CellStyle.Alignment
                    Case DataGridViewContentAlignment.BottomCenter
                        e.Graphics.DrawString(CType(e.Value, String), e.CellStyle.Font, fontBrush, CType(e.CellBounds.X + (cellwidth - fontwidth) / 2, Single), e.CellBounds.Y + cellheight * DownRows - fontheight)
                    Case DataGridViewContentAlignment.BottomLeft
                        e.Graphics.DrawString(CType(e.Value, String), e.CellStyle.Font, fontBrush, e.CellBounds.X, e.CellBounds.Y + cellheight * DownRows - fontheight)
                    Case DataGridViewContentAlignment.BottomRight
                        e.Graphics.DrawString(CType(e.Value, String), e.CellStyle.Font, fontBrush, e.CellBounds.X + cellwidth - fontwidth, e.CellBounds.Y + cellheight * DownRows - fontheight)
                    Case DataGridViewContentAlignment.MiddleCenter
                        e.Graphics.DrawString(CType(e.Value, String), e.CellStyle.Font, fontBrush, CType(e.CellBounds.X + (cellwidth - fontwidth) / 2, Single), CType(e.CellBounds.Y + cellheight * (DownRows - UpRows) + (cellheight * count - fontheight) / 2, Single))
                    Case DataGridViewContentAlignment.MiddleLeft
                        e.Graphics.DrawString(CType(e.Value, String), e.CellStyle.Font, fontBrush, e.CellBounds.X, CType(e.CellBounds.Y - cellheight * (UpRows - 1) + (cellheight * count - fontheight) / 2, Single))
                    Case DataGridViewContentAlignment.MiddleRight
                        e.Graphics.DrawString(CType(e.Value, String), e.CellStyle.Font, fontBrush, e.CellBounds.X + cellwidth - fontwidth, CType(e.CellBounds.Y - cellheight * (UpRows - 1) + (cellheight * count - fontheight) / 2, Single))
                    Case DataGridViewContentAlignment.TopCenter
                        e.Graphics.DrawString(CType(e.Value, String), e.CellStyle.Font, fontBrush, e.CellBounds.X + CType((cellwidth - fontwidth) / 2, Single), e.CellBounds.Y - cellheight * (UpRows - 1))
                    Case DataGridViewContentAlignment.TopLeft
                        e.Graphics.DrawString(CType(e.Value, String), e.CellStyle.Font, fontBrush, e.CellBounds.X, e.CellBounds.Y - cellheight * (UpRows - 1))
                    Case DataGridViewContentAlignment.TopRight
                        e.Graphics.DrawString(CType(e.Value, String), e.CellStyle.Font, fontBrush, e.CellBounds.X + cellwidth - fontwidth, e.CellBounds.Y - cellheight * (UpRows - 1))
                    Case Else
                        e.Graphics.DrawString(CType(e.Value, String), e.CellStyle.Font, fontBrush, e.CellBounds.X + CType((cellwidth - fontwidth) / 2, Single), CType(e.CellBounds.Y - cellheight * (UpRows - 1) + (cellheight * count - fontheight) / 2, Single))
                End Select

                e.Handled = True
            End If


        Catch
        End Try
    End Sub

    ''' <summary>
    ''' 绘制合并表头
    ''' </summary>
    ''' <param name="e">绘图参数集</param>
    ''' <param name="level">结点深度</param>
    ''' <param name="IsDeep"></param>
    Private Sub PaintUnitHeader(
                    ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs,
                    ByVal level As Integer, Optional ByVal IsDeep As Boolean = True)
        Dim iCol As Integer = e.ColumnIndex
        If level = 0 And iCol < 0 Then
            Return
        End If
        '处理

        Me.ColumnHeadersHeight = _cellHeight * _columnDeep
        Dim uhRectangle As Rectangle
        Dim uhWidth As Integer
        Dim gridBrush As New SolidBrush(Me.GridColor)
        Dim backColorBrush As New SolidBrush(e.CellStyle.BackColor)
        Dim gridLinePen As New Pen(gridBrush)
        Dim textFormat As New StringFormat()
        Dim iDeepH As Integer
        textFormat.Alignment = StringAlignment.Center


        uhWidth = GetUnitHeaderWidth(iCol, level)
        If uhWidth <= 0 Then Exit Sub
        iDeepH = IsSingleChildNode(iCol, level)
        If iDeepH < 0 Then Exit Sub


        uhRectangle = New Rectangle(
                                     e.CellBounds.Left,
                                     e.CellBounds.Top + (level - 1 - iDeepH) * _cellHeight,
                                     uhWidth - 1,
                                     _cellHeight * (iDeepH + 1) - 1)
        If e.ColumnIndex = -1 Then
            uhRectangle = Me.GetCellDisplayRectangle(-1, -1, True)
        End If
        With e.Graphics
            '    On Error Resume Next
            .FillRectangle(backColorBrush, uhRectangle)
            '划底线
            .DrawLine(gridLinePen _
                                , uhRectangle.Left _
                                , uhRectangle.Bottom _
                                , uhRectangle.Right _
                                , uhRectangle.Bottom)
            '划右端线
            .DrawLine(gridLinePen _
                                , uhRectangle.Right _
                                , uhRectangle.Top _
                                , uhRectangle.Right _
                                , uhRectangle.Bottom)



            If e.ColumnIndex > -1 Then .DrawString(_ColCap(_columnDeep - level, iCol) _
                                  , _HeaderFont _
                                  , Brushes.Black _
                                  , uhRectangle.Left +
                                  uhRectangle.Width / 2 -
                                  .MeasureString(_ColCap(_columnDeep - level, iCol), _HeaderFont).Width / 2 - 1 _
                                  , uhRectangle.Top +
                                  uhRectangle.Height / 2 -
                                  .MeasureString(_ColCap(_columnDeep - level, iCol), _HeaderFont).Height / 2)


        End With
    End Sub
    '------------------------------------------------------
    '返回上下的跨度,0表示无跨度
    '-----------------------------------------------------
    Private Function IsSingleChildNode(ByVal ColIndex As Integer, ByVal Deep As Integer) As Integer
        Dim strTemp As String = _ColCap(_columnDeep - Deep, ColIndex)
        Dim IRc As Integer = 0
        If Deep < _columnDeep Then
            If strTemp = _ColCap(_columnDeep - Deep - 1, ColIndex) Then Return -1
        End If
        For iD As Integer = _columnDeep - Deep + 1 To _columnDeep - 1
            If strTemp <> _ColCap(iD, ColIndex) Then Exit For
            IRc += 1
        Next
        Return IRc
    End Function


    ''' <summary>
    ''' 获得合并标题字段的宽度
    ''' </summary>
    ''' <param name="ColIndex"></param>
    ''' <param name="Deep"></param>
    ''' <returns>字段宽度</returns>
    Private Function GetUnitHeaderWidth(ByVal ColIndex As Integer, ByVal Deep As Integer) As Integer
        '获得字段的宽度
        Dim uhWidth As Integer = 0

        If ColIndex < 0 Then Return -1
        If Deep = _columnDeep Then '最底层
            Return Me.Columns(ColIndex).Width
        Else
            If ColIndex = 0 Then
                Dim strTemp As String = _ColCap(_columnDeep - Deep, 0)
                For iC As Integer = 0 To Me.Columns.Count - 1
                    uhWidth += Me.Columns(iC).Width
                    If iC + 1 < Me.Columns.Count Then
                        If strTemp <> _ColCap(_columnDeep - Deep, iC + 1) Then Exit For
                    End If
                Next
            ElseIf ColIndex + 1 = Me.Columns.Count Then
                Return Me.Columns(ColIndex).Width
            Else
                Dim strTemp As String = _ColCap(_columnDeep - Deep, ColIndex)
                If strTemp = _ColCap(_columnDeep - Deep, ColIndex - 1) Then
                    Return -1 '因为和前面一致，不再重画
                Else
                    For iC As Integer = ColIndex To Me.Columns.Count - 1
                        uhWidth += Me.Columns(iC).Width
                        If iC + 1 < Me.Columns.Count Then
                            If strTemp <> _ColCap(_columnDeep - Deep, iC + 1) Then Exit For
                        End If
                    Next
                End If
            End If
            Return uhWidth
        End If
    End Function

End Class
