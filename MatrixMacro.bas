Attribute VB_Name = "MatrixMacro"
Option Explicit
Option Base 1

Sub MatrixFunction()

Dim matrix1 As Variant
Dim matrix2 As Variant
Dim numcols As Integer
Dim numrows As Integer
Dim numcols2 As Integer
Dim numrows2 As Integer
Dim Minput As Integer
Dim Minput2 As Integer


numrows = InputBox("Input Number of Rows in Matrix 1")
numcols = InputBox("Input Number of Columns in Matrix 1")

numrows2 = InputBox("Input Number of Rows in Matrix 2")
numcols2 = InputBox("Input Number of Columns in Matrix 2")

Dim i As Integer
Dim j As Integer

ReDim matrix1(1 To numrows, 1 To numcols)

For j = 1 To numrows
     For i = 1 To numcols
         Minput = InputBox("Input Matrix 1 Numbers in Order Left to Right")
         matrix1(j, i) = Minput
     Next i
Next j

ReDim matrix2(1 To numrows, 1 To numcols)

For j = 1 To numrows
     For i = 1 To numcols
         Minput2 = InputBox("Input Matrix 2 Numbers in Order Left to Right")
         matrix2(j, i) = Minput2
     Next i
Next j

Dim ftype As String
Dim solution As Variant

ftype = InputBox("Enter Add, Subtract, Multiply, or Divide")

If ftype = "Subtract" Then

    Call MatrixSub(matrix1, matrix2, numrows, numcols)

ElseIf ftype = "Add" Then

    Call MatrixAdd(matrix1, matrix2, numrows, numcols)

ElseIf ftype = "Multiply" And numcols = numrows2 Then

    Call MatrixMult(matrix1, matrix2, numrows, numcols, numrows2, numcols2)
    
ElseIf ftype = "Divide" And numcols = numrows2 Then

    Call MatrixDiv(matrix1, matrix2, numrows, numcols, numrows2, numcols2)
    
End If


End Sub
