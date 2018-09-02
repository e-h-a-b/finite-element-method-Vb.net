Option Explicit On
Option Strict On

Public Class Support
    Public Property SupportNo As Integer
    Public Property Node As Integer
    Public Property RestraintX As Integer ' 1 or 0
    Public Property RestraintY As Integer

    Public Sub New(SupportNo As Integer, Node As Integer, RestraintX As Integer, RestraintY As Integer)
        _SupportNo = SupportNo
        _Node = Node
        _RestraintX = RestraintX
        _RestraintY = RestraintY

    End Sub

    Public Sub New()

    End Sub

End Class
