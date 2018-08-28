Sub MatrixToArray()
Dim rng1 As Range
Dim matStr1, matStr2, productMatStr, ansMatStr As String
Dim matrix1(), matrix2(), productMatrix(), ansMatrix() As Variant


Set rng1 = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)

ReDim matrix1(rng1.Rows.Count, rng1.Columns.Count) As Variant

i = 1
For Each Row In rng1.Rows
    j = 1
    For Each RowCell In Row.Cells
        matrix1(i, j) = RowCell.Value
        matStr1 = matStr1 & RowCell.Value
        j = j + 1
    Next RowCell
    matStr1 = matStr1 & vbNewLine
    i = i + 1
Next Row

MsgBox matStr1

Set rng2 = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)

ReDim matrix2(rng1.Rows.Count, rng1.Columns.Count) As Variant

n = 1
For Each Row In rng2.Rows
    m = 1
    For Each RowCell In Row.Cells
        matrix2(n, m) = RowCell.Value
        matStr2 = matStr2 & RowCell.Value
        m = m + 1
    Next RowCell
    matStr2 = matStr2 & vbNewLine
    n = n + 1
Next Row

MsgBox matStr2

MsgBox matrix1(1, 1)

ReDim productMatrix(3) As Variant
ReDim ansMatrix(n - 1, m - 1) As Variant





x = 1
For a = 1 To i - 1
    y = 1
    For b = 1 To j - 1
        For c = 1 To m - 1
            productMatrix(c) = matrix1(a, c) * matrix2(c, b)
            productMatStr = productMatStr & productMatrix(x)
        Next c
        For Each Item In productMatrix
            Sum = Sum + Item
        Next Item
        ansMatrix(x, y) = Sum
        ansMatStr = ansMatStr & " " & ansMatrix(x, y)
        Sum = 0
        y = y + 1
    Next b
    x = x + 1
    ansMatStr = ansMatStr & vbNewLine
Next a

MsgBox ansMatStr




End Sub
