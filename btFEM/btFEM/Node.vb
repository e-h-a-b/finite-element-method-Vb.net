Option Explicit On
Option Strict On

Public Class Node
    Public Property NodeNumber As Integer
    Public Property x As Double
    Public Property y As Double


    Public Sub New()

    End Sub

    Public Sub New(nn As Integer, x As Double, y As Double)
        _NodeNumber = nn
        _x = x
        _y = y
    End Sub


End Class
