Option Explicit On
Option Strict On

Public Class Element
    Public Property ElementNumber As Integer
    Public Property Node1 As Integer
    Public Property Node2 As Integer
    Public Property Node3 As Integer
    Public Stresses(2) As Double
    Public Strains(2) As Double

    Public Sub New()

    End Sub

    Public Sub New(n As Integer, n1 As Integer, n2 As Integer, n3 As Integer)
        _ElementNumber = n
        _Node1 = n1
        _Node2 = n2
        _Node3 = n3
    End Sub

End Class
