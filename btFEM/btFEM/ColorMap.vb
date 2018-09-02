Option Explicit On
Option Strict On

Public Class ColorMap
    'A simple implementation of colormap..
    'only red-blue map is implemented.

    Public Property Max As Double
    Public Property Min As Double

    Private nSegs As Integer = 250
    Private Colors(nSegs) As Color

    Public Function getColor(Value As Double) As Color

        'value exists somewhere between min and max..
        'we find color between min color and max color corresponding to value and return
        Dim rangeVal As Double = Max - Min
        Dim dv As Double = Value - Min
        Dim ratio = dv / rangeVal
        Dim colIndex As Integer = CInt(ratio * nSegs)
        Return Colors(colIndex)
    End Function

    Private Sub setColors()

        'Dim mainColors() As Color = {Color.Red, Color.Yellow, Color.LightGreen,
        '    Color.Cyan, Color.Blue}
        Dim mainColors() As Color = {Color.Blue, Color.Cyan, Color.LightGreen,
            Color.Yellow, Color.Red}

        Dim nColors As Integer = mainColors.Length
        'this is from top to bottom.

        Dim percentStep As Double = 100 / (nColors - 1)
        Dim Indices(nColors - 1) As Integer 'we will have these many main indices

        Dim i, j As Integer

        For i = 0 To nColors - 1
            Indices(i) = CInt(i * percentStep / 100 * nSegs)
        Next
        Indices(0) = 0
        Indices(Indices.Length - 1) = nSegs
        ReDim Colors(nSegs)
        Dim c() As Color
        For i = 0 To Indices.Length - 2
            c = getSteppedColors(mainColors(i + 1), mainColors(i), Indices(i + 1) - Indices(i))
            For j = 0 To c.Length - 1
                Colors(Indices(i) + j) = c(j)
            Next
        Next

    End Sub

    Private Function getSteppedColors(cTop As Color, cBottom As Color, nPoints As Integer) As Color()

        Dim btm As Color = cBottom
        Dim top As Color = cTop
        Dim cols(nPoints) As Color

        Dim i As Integer
        Dim dr, dg, db As Integer
        dr = CInt(top.R) - CInt(btm.R)
        dg = CInt(top.G) - (btm.G)
        db = CInt(top.B) - (btm.B)

        'these differences have to be spanned in the nPoints-1 steps
        Dim stepR, stepG, stepB As Double
        stepR = dr / (nPoints)
        stepG = dg / (nPoints)
        stepB = db / (nPoints)

        Dim r, g, b As Integer
        For i = 0 To nPoints
            r = CInt(btm.R + i * stepR)
            g = CInt(btm.G + i * stepG)
            b = CInt(btm.B + i * stepB)

            cols(i) = Color.FromArgb(r, g, b)
        Next
        Return cols
    End Function

    Private Sub setRBColors()
        Dim btm As Color = Color.Blue
        Dim top As Color = Color.Red

        Dim i As Integer
        Dim dr, dg, db As Integer
        dr = CInt(top.R) - CInt(btm.R)
        dg = CInt(top.G) - (btm.G)
        db = CInt(top.B) - (btm.B)

        'these differences have to be spanned in the nSeg steps
        Dim stepR, stepG, stepB As Double
        stepR = dr / nSegs
        stepG = dg / nSegs
        stepB = db / nSegs

        Dim r, g, b As Integer
        For i = 0 To nSegs
            r = CInt(btm.R + i * stepR)
            g = CInt(btm.G + i * stepG)
            b = CInt(btm.B + i * stepB)

            Colors(i) = Color.FromArgb(r, g, b)
        Next

    End Sub



    Public Sub New(Max As Double, Min As Double)
        Me.Max = Max
        Me.Min = Min
        setColors()
    End Sub
End Class
