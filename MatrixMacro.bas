Attribute VB_Name = "MatrixMacro"
Option Explicit
Option Base 1

Function MatrixSub(matrixa As Variant, matrixb As Variant, rows As Integer, cols As Integer) As Variant

ReDim diff(1 To rows, 1 To cols)

Dim i As Integer
Dim j As Integer

For j = 1 To rows
     For i = 1 To cols
         diff(j, i) = matrixa(j, i) - matrixb(j, i)
     Next i
Next j

MatrixSub = diff

Dim p As Integer
Dim pp As Integer
Dim msg As String

For p = 1 To rows
    For pp = 1 To cols
    msg = msg & diff(p, pp) & vbTab
    Next pp
    msg = msg & vbCrLf
Next p
    MsgBox msg


End Function

Function MatrixMult(matrixa As Variant, matrixb As Variant, rows As Integer, cols As Integer, rows2 As Integer, cols2 As Integer) As Variant

ReDim Product(1 To rows, 1 To cols2)
ReDim Product2(1 To rows, 1 To cols2)

Dim i As Integer
Dim j As Integer
Dim b As Integer

For j = 1 To rows
     For i = 1 To cols2
        For b = 1 To rows2
         Product(j, i) = matrixa(j, b) * matrixb(b, i)
         Product2(j, i) = Product2(j, i) + Product(j, i)
        Next b
    Next i
Next j


MatrixMult = Product2

Dim p As Integer
Dim pp As Integer
Dim msg As String

For p = 1 To rows
    For pp = 1 To cols
    msg = msg & Product2(p, pp) & vbTab
    Next pp
    msg = msg & vbCrLf
Next p
    MsgBox msg
    
End Function

Function MatrixAdd(matrixa As Variant, matrixb As Variant, rows As Integer, cols As Integer) As Variant

Dim sum1() As Variant
ReDim sum1(1 To rows, 1 To cols)


Dim i As Integer
Dim j As Integer

For j = 1 To rows
     For i = 1 To cols
         sum1(j, i) = matrixa(j, i) + matrixb(j, i)
     Next i
Next j

MatrixAdd = sum1()

Dim p As Integer
Dim pp As Integer
Dim msg As String

For p = 1 To rows
    For pp = 1 To cols
    msg = msg & sum1(p, pp) & vbTab
    Next pp
    msg = msg & vbCrLf
Next p
    MsgBox msg


End Function

Function MatrixDiv(matrixa As Variant, matrixb As Variant, rows As Integer, cols As Integer, rows2 As Integer, cols2 As Integer) As Variant

Dim MatrixI() As Variant
ReDim MatrixI(1 To rows2, 1 To cols2)

MatrixI = WorksheetFunction.MInverse(matrixb)

ReDim Div(1 To rows, 1 To cols2)
ReDim Div2(1 To rows, 1 To cols2)

Dim i As Integer
Dim j As Integer
Dim b As Integer

For j = 1 To rows
     For i = 1 To cols2
        For b = 1 To rows2
         Div(j, i) = matrixa(j, b) * MatrixI(b, i)
         Div2(j, i) = Div2(j, i) + Div(j, i)
        Next b
    Next i
Next j


MatrixDiv = Div2

Dim p As Integer
Dim pp As Integer
Dim msg As String

For p = 1 To rows
    For pp = 1 To cols
    msg = msg & Div2(p, pp) & vbTab
    Next pp
    msg = msg & vbCrLf
Next p
    MsgBox msg
    
End Function

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
