Option Explicit On
Option Strict On

'Program for solution of linear algebraic equations Ax=B
'using gauss elimination on symmetric banded matrices.
'This class requres that the pivots (A(i,i)) elements being
'non zero... further, it works only on symmetric matrices
'supplied in a half band matrix form. The assembly of matrix
'to half band matrix form is to be done by the caller.
'Date: 25th March 2017

'Some good references:
'http://www.codewithc.com/gauss-elimination-method-algorithm-flowchart/
'http://nptel.ac.in/courses/122104019/numerical-analysis/kadalbajoo/lec1/fnode5.html
'http://www.docentes.unal.edu.co/jmruizv/docs/AnalisisNumerico/clase04.pdf
'https://math.dartmouth.edu/archive/m23s06/public_html/handouts/row_reduction_examples.pdf
'http://web.mit.edu/10.001/Web/Course_Notes/GaussElimPivoting.html
'http://www.math.iitb.ac.in/~neela/partialpivot.pdf

'Find unknown Data
Public Class btGauss

    Private Const Tiny As Double = 0.0000000001

    Private nVariables As Integer
    Private HalfBandWidth As Integer 'Half band width
    Private a(,) As Double 'The coefficient matrix a

    Public Err As Boolean = False


    'We accept the rhs by reference to avoid copying data..
    'we wont modify it nonetheless.
    Public Sub SolveSerial(ByRef x() As Double)
        Dim j, k As Integer
        Dim pRow As Integer 'Counter that runs from 1st to last row - the pivot row
        Dim pivot As Double
        Dim sum As Double

        'Start the elimination loop over all rows
        For pRow = 0 To nVariables - 2
            pivot = a(pRow, 0) 'The entry on principal diagonal.. a(cRows, cRows)
            'work on the submatrix now
            Dim row As Integer
            'The following loop can be parallelized
            Dim aik, akj As Double

            Dim rMax As Integer
            rMax = pRow + HalfBandWidth
            If rMax > nVariables Then rMax = nVariables

            'rMax = pRow + HalfBandWidth
            'If rMax > HalfBandWidth Then rMax = HalfBandWidth


            For row = pRow + 1 To rMax - 1
                'aik = a(row, pRow) -> to be converted to half band repreentation
                'for symmetric matrices, aik = aki..
                aik = a(pRow, row - pRow)

                For col = 0 To HalfBandWidth - 2
                    'akj = a(pRow, col) -> to be converted to half band representation

                    If row - pRow + col < HalfBandWidth Then
                        akj = a(pRow, row - pRow + col)
                        'akj = a(pRow, col + 1)
                        a(row, col) = a(row, col) - aik / pivot * akj
                    End If
                Next
                'Factorize the rhs vector
                'x(pRow) = x(pRow) / pivot
                x(row) = x(row) - aik / pivot * x(pRow)

            Next


        Next

        'Factorize x(n)
        pivot = a(nVariables - 1, 0)
        x(nVariables - 1) = x(nVariables - 1) / pivot

        'Back substitution  
        Dim uBound As Integer

        For k = nVariables - 2 To 0 Step -1
            sum = 0.0
            uBound = HalfBandWidth - 1
            If uBound + k >= nVariables Then uBound = nVariables - k - 1
            For j = 1 To uBound
                sum = sum + a(k, j) * x(j + k)
            Next
            x(k) = 1.0 / a(k, 0) * (x(k) - sum)
        Next

    End Sub

    Public Sub New(BandedMatrix(,) As Double)
        nVariables = BandedMatrix.GetLength(0)
        HalfBandWidth = BandedMatrix.GetLength(1)

        ReDim a(nVariables - 1, HalfBandWidth - 1)
        a = BandedMatrix
    End Sub


    Public Function Solve(ByRef rhs() As Double) As Double()
        Dim j, k As Integer
        Dim pRow As Integer 'Counter that runs from 1st to last row - the pivot row
        Dim pivot As Double
        Dim sum As Double

        Dim x(rhs.Length - 1) As Double
        Array.Copy(rhs, x, rhs.Length)

        'Start the elimination loop over all rows
        For pRow = 0 To nVariables - 2
            pivot = a(pRow, 0) 'The entry on principal diagonal.. a(cRows, cRows)
            'work on the submatrix now

            'The following loop can be parallelized

            Dim rMax As Integer
            rMax = pRow + HalfBandWidth
            If rMax > nVariables Then rMax = nVariables

            Parallel.For(pRow + 1, rMax,
                         Sub(row As Integer)
                             Dim aik, akj As Double


                             'aik = a(row, pRow) -> to be converted to half band repreentation
                             'for symmetric matrices, aik = aki..
                             aik = a(pRow, row - pRow) / pivot

                             'If Math.Abs(aik) > Tiny Then
                             For col = 0 To HalfBandWidth - 2
                                 'akj = a(pRow, col) -> to be converted to half band representation
                                 'akj = a(pRow, col + 1)

                                 If row - pRow + col < HalfBandWidth Then
                                     akj = a(pRow, row - pRow + col)
                                     'akj = a(pRow, col + 1)
                                     a(row, col) = a(row, col) - aik * akj
                                 End If

                                 'a(row, col) = a(row, col) - aik / pivot * akj
                             Next
                             'Factorize the rhs vector
                             'x(pRow) = x(pRow) / pivot
                             x(row) = x(row) - aik * x(pRow)

                             'End If
                         End Sub)

        Next

        'Factorize x(n)
        pivot = a(nVariables - 1, 0)
        x(nVariables - 1) = x(nVariables - 1) / pivot

        'Back substitution  
        Dim uBound As Integer

        For k = nVariables - 2 To 0 Step -1
            sum = 0.0
            uBound = HalfBandWidth - 1
            If uBound + k >= nVariables Then uBound = nVariables - k - 1
            For j = 1 To uBound
                sum = sum + a(k, j) * x(j + k)
            Next
            x(k) = 1.0 / a(k, 0) * (x(k) - sum)
        Next
        Return x
    End Function

End Class

