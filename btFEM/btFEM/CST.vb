Option Explicit On
Option Strict On

Public Class CST
    'Implementation of the constant strain triangle

    'Element properties
    Private x1, y1, x2, y2, x3, y3, Thickness, ElasticityModulus, PoissonRatio As Double

    'Distances
    Private x13, x32, x21, x23, y13, y23, y31, y12 As Double


    Private Enum ProblemTypes
        PlaneStress
        PlaneStrain
    End Enum
    Private ProblemType As ProblemTypes = ProblemTypes.PlaneStress

    ''' <summary>
    ''' Initializes a new instance of CST element
    ''' </summary>
    ''' <param name="x1"></param>
    ''' <param name="y1"></param>
    ''' <param name="x2"></param>
    ''' <param name="y2"></param>
    ''' <param name="x3"></param>
    ''' <param name="y3"></param>
    ''' <param name="Thickness"></param>
    ''' <param name="ElasticityModulus"></param>
    ''' <param name="PoissonRatio"></param>
    Public Sub New(x1 As Double, y1 As Double,
                                 x2 As Double, y2 As Double,
                                 x3 As Double, y3 As Double,
                                 Thickness As Double,
                                 ElasticityModulus As Double,
                                 PoissonRatio As Double)
        Me.x1 = x1 : Me.y1 = y1
        Me.x2 = x2 : Me.y2 = y2
        Me.x3 = x3 : Me.y3 = y3
        Me.Thickness = Thickness
        Me.ElasticityModulus = ElasticityModulus
        Me.PoissonRatio = PoissonRatio

        'calculate the distances. they will be required
        'for computing the jacobian and B matrices
        x13 = x1 - x3
        x32 = x3 - x2
        x21 = x2 - x1
        x23 = x2 - x3
        y13 = y1 - y3
        y23 = y2 - y3
        y31 = y3 - y1
        y12 = y1 - y2

    End Sub


    ''' <summary>
    ''' Returns the determinant of the Jacobian
    ''' </summary>
    ''' <returns></returns>
    Private Function getDetJ() As Double
        Return (x13 * y23 - x23 * y13)
    End Function

    ''' <summary>
    ''' Returns the [B] matrix
    ''' </summary>
    ''' <returns></returns>
    Public Function getB() As Double(,)
        Dim B(2, 5) As Double 'initialize a 3x6 array
        Dim detJ As Double = getDetJ()

        B(0, 0) = y23 / detJ
        B(1, 0) = 0.0
        B(2, 0) = x32 / detJ

        B(0, 1) = 0.0
        B(1, 1) = x32 / detJ
        B(2, 1) = y23 / detJ

        B(0, 2) = y31 / detJ
        B(1, 2) = 0.0
        B(2, 2) = x13 / detJ

        B(0, 3) = 0.0
        B(1, 3) = x13 / detJ
        B(2, 3) = y31 / detJ

        B(0, 4) = y12 / detJ
        B(1, 4) = 0.0
        B(2, 4) = x21 / detJ

        B(0, 5) = 0.0
        B(1, 5) = x21 / detJ
        B(2, 5) = y12 / detJ

        Return B
    End Function

    Private Function getAe() As Double
        Return getDetJ() * 0.5
    End Function

    Public Function getD() As Double(,)
        Dim D(2, 2) As Double
        Dim c As Double
        If ProblemType = ProblemTypes.PlaneStress Then
            c = ElasticityModulus / (1.0 - (PoissonRatio * PoissonRatio))

            D(0, 0) = 1.0 * c
            D(1, 0) = PoissonRatio * c
            D(2, 0) = 0.0

            D(0, 1) = PoissonRatio * c
            D(1, 1) = 1.0 * c
            D(2, 1) = 0.0

            D(0, 2) = 0.0
            D(1, 2) = 0.0
            D(2, 2) = (1 - PoissonRatio) * 0.5 * c

        Else
            'plane strain
            c = ElasticityModulus / ((1 + PoissonRatio) * (1.0 - 2 * PoissonRatio))
            D(0, 0) = (1.0 - PoissonRatio) * c
            D(1, 0) = PoissonRatio * c
            D(2, 0) = 0.0

            D(0, 1) = PoissonRatio * c
            D(1, 1) = (1 - PoissonRatio) * c
            D(2, 1) = 0.0

            D(0, 2) = 0.0
            D(1, 2) = 0.0
            D(2, 2) = (0.5 - PoissonRatio) * c
        End If
        Return D
    End Function

    Public Function getKe() As Double(,)

        Dim B(,) As Double = getB()
        Dim BT(,) As Double = getBT(B)
        Dim D(,) As Double = getD()
        Dim teAe As Double = Thickness * getAe()

        'Compute Ke now
        Dim BTD(,) As Double = MultiplyMatrices(BT, D)
        Dim Ke(,) As Double = MultiplyMatrices(BTD, B)

        'multiply ke by teAe
        Dim i, j As Integer
        For i = 0 To 5
            For j = 0 To 5
                ke(i, j) = ke(i, j) * teAe
            Next
        Next
        Return ke
    End Function


    Private Function getBT(ByRef B(,) As Double) As Double(,)
        'returns the transpose of [B]
        Dim BT(5, 2) As Double
        Dim i, j As Integer
        For i = 0 To 2
            For j = 0 To 5
                BT(j, i) = B(i, j)
            Next
        Next
        Return BT
    End Function

    ''' <summary>
    ''' Multiplies matrix a by b and returns the product
    ''' </summary>
    ''' <param name="a"></param>
    ''' <param name="b"></param>
    ''' <returns></returns>
    Private Function MultiplyMatrices(ByRef a(,) As Double, ByRef b(,) As Double) As Double(,)

        Dim aCols As Integer = a.GetLength(1)
        Dim bCols As Integer = b.GetLength(1)
        Dim aRows As Integer = a.GetLength(0)
        Dim ab(aRows - 1, bCols - 1) As Double
        Dim sum As Double
        For i As Integer = 0 To aRows - 1
            For j As Integer = 0 To bCols - 1
                sum = 0.0
                For k As Integer = 0 To aCols - 1
                    sum += a(i, k) * b(k, j)
                Next
                ab(i, j) += sum
            Next
        Next

        Return ab
    End Function

End Class
