Sub MatrixToArray()
Dim rng1 As Range
Dim matStr1 As String
Dim matrix1() As Variant


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
