Option Explicit On
Option Strict On

Public Class PointLoad
    Public Property Node As Integer
    Public Property PointLoadNo As Integer
    Public Property Fx As Double
    Public Property Fy As Double

    Public Sub New(PointLoadNo As Integer, Node As Integer, Fx As Double, Fy As Double)
        _PointLoadNo = PointLoadNo
        _Node = Node
        _Fx = Fx
        _Fy = Fy

    End Sub

    Public Sub New()

    End Sub


End Class
