VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''' <summary>
''' Этот класс содержит методы для поиска диапазона таблицы на листе Excel.
''' </summary>

''' <summary>
''' Находит диапазон таблицы в заданном диапазоне Range.
''' </summary>
''' <param name="rng">Диапазон Range, в котором нужно найти таблицу.</param>
''' <returns>Диапазон Range, представляющий таблицу.</returns>
''' <example>
''' Пример использования:
''' <code>
''' Dim tableRange As Range
''' Set tableRange = FindTableRange(Range("A1:F10"))
''' </code>
''' </example>
Public Function FindTableRange(ByVal rng As Range) As Range
    Dim top As Long, bottom As Long, left As Long, right As Long
    
    If WorksheetFunction.CountA(rng) = 0 Then
        Set FindTableRange = Nothing
        Exit Function
    End If
    
    top = FindBoundary(rng, False, xlNext)
    bottom = FindBoundary(rng, False, xlPrevious)
    left = FindBoundary(rng, True, xlNext)
    right = FindBoundary(rng, True, xlPrevious)
    
    If top > 0 And bottom > 0 And left > 0 And right > 0 Then
        Set FindTableRange = rng.Parent.Range(rng.Cells(top, left), rng.Cells(bottom, right))
    End If
End Function

''' <summary>
''' Находит границы таблицы в заданном диапазоне Range по заданному направлению.
''' </summary>
''' <param name="rng">Диапазон Range, в котором нужно найти границы таблицы.</param>
''' <param name="isRow">Указывает, ищутся ли границы строк.</param>
''' <param name="Direction">Направление поиска границ таблицы.</param>
''' <returns>Число, представляющее границу таблицы.</returns>
Private Function FindBoundary(ByVal rng As Range, ByVal isRow As Boolean, ByVal Direction As XlSearchDirection) As Long
    Dim Arr() As Long
    Arr = FindBoundaryArray(rng, isRow, Direction)
    
    If UBound(Arr) > 0 Then
        If isRow Then
            FindBoundary = WorksheetFunction.Median(Arr)
        Else
            FindBoundary = WorksheetFunction.Max(Arr)
        End If
    Else
        FindBoundary = 0
    End If
End Function

''' <summary>
''' Находит массив границ таблицы в заданном диапазоне Range по заданному направлению.
''' </summary>
''' <param name="rng">Диапазон Range, в котором нужно найти границы таблицы.</param>
''' <param name="isRow">Указывает, ищутся ли границы строк.</param>
''' <param name="Direction">Направление поиска границ таблицы.</param>
''' <returns>Массив чисел, представляющий границы таблицы.</returns>
Private Function FindBoundaryArray(ByVal rng As Range, ByVal isRow As Boolean, ByVal Direction As XlSearchDirection) As Variant
    Dim i As Long
    Dim Arr() As Long
    
    If WorksheetFunction.CountA(rng) = 0 Then
        FindBoundaryArray = 0
        Exit Function
    End If
    
    If isRow And Direction = xlNext Then
        ReDim Arr(1 To rng.rows.Count)
        For i = 1 To rng.rows.Count
            Arr(i) = FindLeftBoundary(rng, i)
        Next i
    ElseIf isRow And Direction = xlPrevious Then
        ReDim Arr(1 To rng.rows.Count)
        For i = 1 To rng.rows.Count
            Arr(i) = FindRightBoundary(rng, i)
        Next i
    ElseIf Not isRow And Direction = xlNext Then
        ReDim Arr(1 To rng.Columns.Count)
        For i = 1 To rng.Columns.Count
            Arr(i) = FindTopBoundary(rng, i)
        Next i
    ElseIf Not isRow And Direction = xlPrevious Then
        ReDim Arr(1 To rng.Columns.Count)
        For i = 1 To rng.Columns.Count
            Arr(i) = FindBottomBoundary(rng, i)
        Next i
    End If
    
    FindBoundaryArray = Arr
End Function

''' <summary>
''' Находит правую границу таблицы в заданной строке диапазона Range.
''' </summary>
''' <param name="rng">Диапазон Range, в котором нужно найти границу таблицы.</param>
''' <param name="rowIdx">Индекс строки, в которой нужно найти границу таблицы.</param>
''' <returns>Число, представляющее правую границу таблицы.</returns>
Private Function FindRightBoundary(ByVal rng As Range, ByVal rowIdx As Long) As Long
    Dim colIdx As Long
    For colIdx = rng.Columns.Count To 1 Step -1
        If Not IsEmpty(rng(rowIdx, colIdx)) Then
            FindRightBoundary = colIdx
            Exit Function
        End If
    Next colIdx
    FindRightBoundary = 0
End Function

''' <summary>
''' Находит нижнюю границу таблицы в заданном столбце диапазона Range.
''' </summary>
''' <param name="rng">Диапазон Range, в котором нужно найти границу таблицы.</param>
''' <param name="colIdx">Индекс столбца, в котором нужно найти границу таблицы.</param>
''' <returns>Число, представляющее нижнюю границу таблицы.</returns>
Private Function FindBottomBoundary(ByVal rng As Range, ByVal colIdx As Long) As Long
    Dim rowIdx As Long
    For rowIdx = rng.rows.Count To 1 Step -1
        If Not IsEmpty(rng(rowIdx, colIdx)) Then
            FindBottomBoundary = rowIdx
            Exit Function
        End If
    Next rowIdx
    FindBottomBoundary = 0
End Function

''' <summary>
''' Находит левую границу таблицы в заданной строке диапазона Range.
''' </summary>
''' <param name="rng">Диапазон Range, в котором нужно найти границу таблицы.</param>
''' <param name="rowIdx">Индекс строки, в которой нужно найти границу таблицы.</param>
''' <returns>Число, представляющее левую границу таблицы.</returns>
Private Function FindLeftBoundary(ByVal rng As Range, ByVal rowIdx As Long) As Long
    Dim colIdx As Long
    For colIdx = 1 To rng.Columns.Count Step 1
        If Not IsEmpty(rng(rowIdx, colIdx)) Then
            FindLeftBoundary = colIdx
            Exit Function
        End If
    Next colIdx
    FindLeftBoundary = 0
End Function

''' <summary>
''' Находит верхнюю границу таблицы в заданном столбце диапазона Range.
''' </summary>
''' <param name="rng">Диапазон Range, в котором нужно найти границу таблицы.</param>
''' <param name="colIdx">Индекс столбца, в котором нужно найти границу таблицы.</param>
''' <returns>Число, представляющее верхнюю границу таблицы.</returns>
Private Function FindTopBoundary(ByVal rng As Range, ByVal colIdx As Long) As Long
    Dim rowIdx As Long
    For rowIdx = 1 To rng.rows.Count Step 1
        If Not IsEmpty(rng(rowIdx, colIdx)) Then
            FindTopBoundary = rowIdx
            Exit Function
        End If
    Next rowIdx
    FindTopBoundary = 0
End Function

