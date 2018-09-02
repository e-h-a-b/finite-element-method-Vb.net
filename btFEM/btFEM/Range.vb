Option Explicit On
Option Strict On

Public Class Range
    Public Property Max As Double = 0
    Public Property Min As Double = 0

    Public Sub AddValue(value As Double)
        If value > Max Then Max = value
        If value < Min Then Min = value
    End Sub
    Public Sub AddValue(values() As Double)
        Dim i As Integer
        For i = 0 To values.Length - 1
            If values(i) > Max Then Max = values(i)
            If values(i) < Min Then Min = values(i)
        Next
    End Sub

    Public Function Span() As Double
        Return Max - Min
    End Function


End Class
