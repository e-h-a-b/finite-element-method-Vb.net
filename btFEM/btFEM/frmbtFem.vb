Option Explicit On
Option Strict On

Imports System
Imports System.ComponentModel
Imports System.IO

Public Class frmbtFem
    'This program shows the complete implementation
    'of a finite element program. Constant strain
    'triangular element has been chosen for
    'demonstration purpose. Any other element may
    'be used with suitable modifications. Please
    'note that Constant Strain Triangles (CST) are
    'not recommended for stress analysis of practical
    'real time problems.

    'Written by Dr. Bhairav Thakkar
    '23rd September 2017
    'Halifax, NS
    'Canada

    Private NNodes, NElements, NPointLoads, NSupports As Integer
    Private ElasticityModulus, Thickness, PoissonRatio As Double

    Private Nodes() As Node
    Private Elements() As Element
    Private PointLoads() As PointLoad
    Private Supports() As Support
    Private Deformations() As Double 'the deformation vector


    Private ShowDeformations As Boolean = True
    Private ShowModel As Boolean = True
    Private ShowNodeNumbers As Boolean = False
    Private ShowElementsOnDeformedShape As Boolean = True

    Private colorMap As ColorMap
    Private SigmaXRange, SigmaYRange, TauXYRange As Range
    Private EpsilonXRange, EpsilonYRange, GammaXYRange As Range


    Private Enum Results
        None
        SigmaX
        SigmaY
        TauXY
        EpsilonX
        EpsilonY
        GammaXY
    End Enum
    Private ShowResult As Results = Results.None



    'Graphics
    Private zoom As Double = 1.0
    Private DeformationZoom As Double = 50.0
    Private gr As Graphics

    Private Sub frmbtFem_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bmp = New Bitmap(pbModel.Width, pbModel.Height, pbModel.CreateGraphics)
        gr = Graphics.FromImage(bmp)
    End Sub

    Private bmp As Bitmap


    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Dim d As New OpenFileDialog
        d.Title = "Open FEM file"
        d.Filter = "FEm Files (*.fem)|*.fem|All Files|*.*"
        If d.ShowDialog = DialogResult.OK Then
            ClearAnalysisOutput()
            If ReadFile(d.FileName) = False Then Return
            DrawModel()
        End If
    End Sub

    Private Sub ClearAnalysisOutput()
        Erase Deformations 'Clear the earlier analysis output
    End Sub

    Private Sub Analyse()
        'Calculate statistics
        Dim nDof As Integer = NNodes * 2


        'make global stiffness matrix
        'we create elemental stiffness matrix and place it in proper position
        'in the global stiffness matrix
        'saving as a square matrix would be too taxing on ram.
        'hence, we save the stiffness matrix in half band form

        Dim HalfBandWidth As Integer = getHalfBandWidth()
        Dim Kg(nDof - 1, HalfBandWidth - 1) As Double 'Global stiffness matrix

        Dim cst As CST
        For i = 0 To NElements - 1
            cst = New CST(Nodes(Elements(i).Node1 - 1).x, Nodes(Elements(i).Node1 - 1).y,
                          Nodes(Elements(i).Node2 - 1).x, Nodes(Elements(i).Node2 - 1).y,
                          Nodes(Elements(i).Node3 - 1).x, Nodes(Elements(i).Node3 - 1).y,
                          Thickness, ElasticityModulus, PoissonRatio)


            If testKe(cst.getKe) = False Then
                MsgBox("ke test failed")
            End If

            AssembleKg(cst.getKe, Kg, i)

        Next
        'now, we have assembled Kg in banded form

        'make the load vector
        Dim r(nDof - 1) As Double
        Dim dof As Integer 'dof
        For i = 0 To NPointLoads - 1
            dof = getDOFx(PointLoads(i).Node - 1)
            r(dof) += PointLoads(i).Fx

            dof = getDOFy(PointLoads(i).Node - 1)
            r(dof) += PointLoads(i).Fy

        Next


        'Apply boundary conditions by penalty approach
        Dim Large As Double
        Dim pow As Double = MaxKgiiPower(Kg)
        Large = 10 ^ (Math.Ceiling(pow) + 10)
        For i = 0 To NSupports - 1
            If Supports(i).RestraintX = 1 Then
                dof = getDOFx(Supports(i).Node - 1)
                Kg(dof, 0) = Kg(dof, 0) * Large
            End If
            If Supports(i).RestraintY = 1 Then
                dof = getDOFy(Supports(i).Node - 1)
                Kg(dof, 0) = Kg(dof, 0) * Large
            End If

        Next


        'Solve the set of equations to get displacements.

        'Deformations = SolveWithPivoting(Kg, r)


        Dim bs As New btGauss(Kg)
        bs.SolveSerial(r)
        Deformations = r

        'Now that we have deformations, lets calculate stresses and strains..
        SigmaXRange = New Range : SigmaYRange = New Range : TauXYRange = New Range
        EpsilonXRange = New Range : EpsilonYRange = New Range : GammaXYRange = New Range

        'first is to calculate strains
        'strain = [B] * d.
        'since the element is a constant strain triangle, the strain value will be 
        'constant for each element

        Dim B(2, 5) As Double
        Dim n1, n2, n3 As Integer
        Dim dofs(5) As Integer
        Dim disp(5) As Double
        Dim DMatrix(,) As Double
        For i = 0 To NElements - 1
            cst = New CST(Nodes(Elements(i).Node1 - 1).x, Nodes(Elements(i).Node1 - 1).y,
                          Nodes(Elements(i).Node2 - 1).x, Nodes(Elements(i).Node2 - 1).y,
                          Nodes(Elements(i).Node3 - 1).x, Nodes(Elements(i).Node3 - 1).y,
                          Thickness, ElasticityModulus, PoissonRatio)
            B = cst.getB

            'extract the connectivity nodes
            n1 = Elements(i).Node1 - 1
            n2 = Elements(i).Node2 - 1
            n3 = Elements(i).Node3 - 1

            'find the degree of freedoms associated
            dofs(0) = getDOFx(n1) : dofs(1) = getDOFy(n1)
            dofs(2) = getDOFx(n2) : dofs(3) = getDOFy(n2)
            dofs(4) = getDOFx(n3) : dofs(5) = getDOFy(n3)

            'extract the displacements
            For j = 0 To 5
                disp(j) = Deformations(dofs(j))
            Next

            'Calculate the strain vector
            'epsilon = {ex, ey, gamma xy}T
            Elements(i).Strains = MultiplyMatrixWithVector(B, disp)

            'Calculate stresses
            'Sigma = D * Epsilon
            DMatrix = cst.getD

            Elements(i).Stresses = MultiplyMatrixWithVector(DMatrix, Elements(i).Strains)
            SigmaXRange.AddValue(Elements(i).Stresses(0))
            SigmaYRange.AddValue(Elements(i).Stresses(1))
            TauXYRange.AddValue(Elements(i).Stresses(2))

            EpsilonXRange.AddValue(Elements(i).Strains(0))
            EpsilonYRange.AddValue(Elements(i).Strains(1))
            GammaXYRange.AddValue(Elements(i).Strains(2))

        Next

        'now, we have all results
        'enjoy! :o)

    End Sub

    Private Function MaxKgiiPower(ByRef Kg(,) As Double) As Double
        Dim Max As Double = Double.MinValue
        Dim i As Integer
        For i = 0 To Kg.GetLength(0) - 1
            If Kg(i, 0) > Max Then Max = Kg(i, 0)
        Next
        Dim p As Double = Math.Log10(Max)
        Return p
    End Function

    ''' <summary>
    ''' Multiplies matrix a by vector b and returns the product
    ''' </summary>
    ''' <param name="a"></param>
    ''' <param name="b"></param>
    ''' <returns></returns>
    Private Function MultiplyMatrixWithVector(ByRef a(,) As Double, ByRef b() As Double) As Double()

        Dim aRows As Integer = a.GetLength(0)
        Dim aCols As Integer = a.GetLength(1)
        Dim ab(aRows - 1) As Double 'output will be a vector
        For i As Integer = 0 To aRows - 1
            ab(i) = 0.0
            For j As Integer = 0 To aCols - 1
                ab(i) += a(i, j) * b(j)
            Next
        Next

        Return ab
    End Function


    Private Function testKe(ByRef Ke(,) As Double) As Boolean
        Dim i, j As Integer
        For i = 0 To 5
            If Ke(i, i) < 0.0000001 Then Return False
        Next
        For i = 0 To 5
            For j = 0 To 5
                If Math.Abs(Ke(i, j) - Ke(j, i)) > 0.00001 Then Return False
            Next
        Next

        Return True
    End Function




    Private Sub AssembleKg(ByRef Ke(,) As Double, ByRef Kg(,) As Double, ElementNo As Integer)
        Dim i, j As Integer
        Dim dofs() As Integer = {getDOFx(Elements(ElementNo).Node1 - 1),
                                 getDOFy(Elements(ElementNo).Node1 - 1),
                                 getDOFx(Elements(ElementNo).Node2 - 1),
                                 getDOFy(Elements(ElementNo).Node2 - 1),
                                 getDOFx(Elements(ElementNo).Node3 - 1),
                                 getDOFy(Elements(ElementNo).Node3 - 1)}

        'Place the upper triangle of the elemental stiffness matrix in the global
        'matrix in proper position
        Dim dofi, dofj As Integer
        For i = 0 To 5 'each dof of the ke
            dofi = dofs(i)
            For j = 0 To 5
                dofj = dofs(j) - dofi
                If dofj >= 0 Then
                    Kg(dofi, dofj) = Kg(dofi, dofj) + Ke(i, j)
                End If
            Next
        Next

    End Sub

    Private Function getDOFx(NodeNo As Integer) As Integer
        Dim nDofsPerNode As Integer = 2
        Return (NodeNo) * nDofsPerNode
    End Function
    Private Function getDOFy(NodeNo As Integer) As Integer
        Dim nDofsPerNode As Integer = 2
        Return NodeNo * nDofsPerNode + 1
    End Function

    Private Function getHalfBandWidth() As Integer
        Dim i As Integer
        Dim n(2) As Integer
        Dim MaxDiff As Integer = 0
        Dim diff As Integer
        For i = 0 To NElements - 1
            n(0) = Elements(i).Node1
            n(1) = Elements(i).Node2
            n(2) = Elements(i).Node3
            diff = n.Max - n.Min
            If MaxDiff < diff Then MaxDiff = diff
        Next

        'now we have maxdiff
        'half band width is maxdiff * 2. 2 because there are 2 dofs per node

        Dim hbw As Integer = (MaxDiff + 1) * 2
        If hbw > NNodes * 2 Then hbw = NNodes * 2

        Return hbw
    End Function


    Private Function ReadFile(f As String) As Boolean
        Try
            Dim sr As New StreamReader(f)
            Dim s As String
            Dim Arr() As String
            s = readLine(sr)
            Arr = Split(s, ","c)

            Try
                NNodes = Integer.Parse(Arr(0))
                NElements = Integer.Parse(Arr(1))
                NPointLoads = Integer.Parse(Arr(2))
                NSupports = Integer.Parse(Arr(3))

                ElasticityModulus = Double.Parse(Arr(4))
                PoissonRatio = Double.Parse(Arr(5))
                Thickness = Double.Parse(Arr(6))
            Catch ex As Exception
                MsgBox(ex.Message)
                Return False
            End Try

            'got the basics.. now, lets read each entity
            ReDim Nodes(NNodes - 1)
            ReDim Elements(NElements - 1)

            Dim i As Integer
            Dim n, n1, n2, n3 As Integer
            Dim x, y As Double
            For i = 0 To NNodes - 1
                s = readLine(sr)
                s = s.Replace(vbTab, "")

                Try
                    Arr = Split(s, ","c)
                    n = Integer.Parse(Arr(0))
                    x = Double.Parse(Arr(1))
                    y = Double.Parse(Arr(2))
                    Nodes(i) = New Node(n, x, y)
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Return False
                End Try
            Next

            'read elements
            For i = 0 To NElements - 1
                s = readLine(sr)
                Try
                    Arr = Split(s, ","c)
                    n = Integer.Parse(Arr(0))
                    n1 = Integer.Parse(Arr(1))
                    n2 = Integer.Parse(Arr(2))
                    n3 = Integer.Parse(Arr(3))

                    Elements(i) = New Element(n, n1, n2, n3)

                Catch ex As Exception
                    MsgBox(ex.Message)
                    Return False
                End Try
            Next

            'Read Point Loads
            ReDim PointLoads(NPointLoads - 1)
            Dim Fx, Fy As Double
            Dim ln As Integer
            For i = 0 To NPointLoads - 1
                s = readLine(sr)
                Try
                    Arr = Split(s, ","c)
                    ln = Integer.Parse(Arr(0))
                    n = Integer.Parse(Arr(1))
                    Fx = Double.Parse(Arr(2))
                    Fy = Double.Parse(Arr(3))

                    PointLoads(i) = New PointLoad(ln, n, Fx, Fy)

                Catch ex As Exception
                    MsgBox(ex.Message)
                    Return False
                End Try
            Next

            'Read supports
            ReDim Supports(NSupports - 1)
            Dim sn, Rx, Ry As Integer
            For i = 0 To NSupports - 1
                s = readLine(sr)
                Try
                    Arr = Split(s, ","c)
                    sn = Integer.Parse(Arr(0))
                    n = Integer.Parse(Arr(1))
                    Rx = Integer.Parse(Arr(2))
                    Ry = Integer.Parse(Arr(3))

                    Supports(i) = New Support(sn, n, Rx, Ry)

                Catch ex As Exception
                    MsgBox(ex.Message)
                    Return False
                End Try
            Next


        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
        Return True
    End Function

    Private Sub btnResetZoom_Click(sender As Object, e As EventArgs) Handles btnResetZoom.Click
        zoom = 1
        hs.Value = 0
        vs.Value = 0
        DrawModel()
    End Sub

    Private Sub vs_Scroll(sender As Object, e As ScrollEventArgs) Handles vs.Scroll
        DrawModel()
    End Sub

    Private Sub hs_Scroll(sender As Object, e As ScrollEventArgs) Handles hs.Scroll
        DrawModel()
    End Sub

    Private Function readLine(ByRef sr As StreamReader) As String
        Dim s As String
        If sr.EndOfStream = True Then Return ""
        While sr.EndOfStream = False
            s = sr.ReadLine.Trim
            If s.Length > 0 Then
                If s.Substring(0, 1) <> "%" Then
                    Return s
                End If
            End If
        End While
        Return ""
    End Function

    Private Sub DrawModel()
        gr.Clear(Color.White)

        Dim eleColor As Color = Color.Lavender
        Dim eleColorDeformed As Color = Color.Turquoise
        Dim supportColor As Color = Color.Black
        Dim pointLoadColor As Color = Color.Red

        Dim MarginX, MarginY As Integer
        MarginX = 100
        MarginY = 100

        'Draw the triangles.
        Dim shiftx As Integer = MarginX - hs.Value
        Dim shifty As Integer = pbModel.Height - MarginY - vs.Value

        Dim i As Integer
        Dim n1, n2, n3 As Integer

        Dim ptsf(2) As PointF


        'Draw the undeformed model

        If ShowModel = True Then
            'Draw the elements .
            For i = 0 To NElements - 1

                n1 = Elements(i).Node1 - 1
                n2 = Elements(i).Node2 - 1
                n3 = Elements(i).Node3 - 1

                ptsf(0) = New PointF(CSng(Nodes(n1).x * zoom + shiftx), CSng(-Nodes(n1).y * zoom + shifty))
                ptsf(1) = New PointF(CSng(Nodes(n2).x * zoom + shiftx), CSng(-Nodes(n2).y * zoom + shifty))
                ptsf(2) = New PointF(CSng(Nodes(n3).x * zoom + shiftx), CSng(-Nodes(n3).y * zoom + shifty))

                'Draw the element
                gr.FillPolygon(New SolidBrush(eleColor), ptsf)
                gr.DrawPolygon(New Pen(Color.Black), ptsf)

            Next

            'Draw the node numbers
            If ShowNodeNumbers = True Then
                Dim p As New Pen(Color.Black)
                Dim nx, ny As Integer
                Dim fnt As New Font("Arial", 8)
                Dim brsh As New SolidBrush(Color.Red)
                For i = 0 To NNodes - 1
                    nx = CInt(Nodes(i).x * zoom + shiftx)
                    ny = CInt(-Nodes(i).y * zoom + shifty)
                    Dim r As New Rectangle(nx - 3, ny - 3, 6, 6)
                    gr.DrawEllipse(p, r)

                    gr.DrawString((i + 1).ToString, fnt, brsh, New PointF(nx + 5, ny + 5))

                Next
            End If

            'Draw the supports
            'we draw squares for supports
            ReDim ptsf(3)
            Dim squareHalfSize As Integer = 5
            For i = 0 To NSupports - 1
                n1 = Supports(i).Node - 1
                ptsf(0) = New PointF(CSng(Nodes(n1).x * zoom + shiftx - squareHalfSize),
                                     CSng(-Nodes(n1).y * zoom + shifty - squareHalfSize))
                ptsf(1) = New PointF(CSng(Nodes(n1).x * zoom + shiftx + squareHalfSize),
                                     CSng(-Nodes(n1).y * zoom + shifty - squareHalfSize))
                ptsf(2) = New PointF(CSng(Nodes(n1).x * zoom + shiftx + squareHalfSize),
                                     CSng(-Nodes(n1).y * zoom + shifty + squareHalfSize))
                ptsf(3) = New PointF(CSng(Nodes(n1).x * zoom + shiftx - squareHalfSize),
                                     CSng(-Nodes(n1).y * zoom + shifty + squareHalfSize))
                gr.FillPolygon(New SolidBrush(supportColor), ptsf)

            Next

            'draw the loads
            Dim px, py As Integer
            For i = 0 To NPointLoads - 1
                n1 = PointLoads(i).Node - 1
                If Math.Abs(PointLoads(i).Fx) > 0.0001 Then
                    'if the pointload is significant
                    'draw the arrow
                    px = CInt(Nodes(n1).x * zoom + shiftx)
                    py = CInt(-Nodes(n1).y * zoom + shifty)
                    If PointLoads(i).Fx > 0 Then
                        DrawHArrow(px, py, pointLoadColor, False)
                    Else
                        DrawHArrow(px, py, pointLoadColor, True)
                    End If
                End If

                If Math.Abs(PointLoads(i).Fy) > 0.0001 Then
                    px = CInt(Nodes(n1).x * zoom + shiftx)
                    py = CInt(-Nodes(n1).y * zoom + shifty)
                    If PointLoads(i).Fy > 0 Then
                        DrawVArrow(px, py, pointLoadColor, False)
                    Else
                        DrawVArrow(px, py, pointLoadColor, True)
                    End If

                End If


            Next
        End If
        'Now, plot results if available
        If Deformations IsNot Nothing AndAlso Deformations.Length > 0 Then
            Dim eColor As Color
            'we can plot now.
            If ShowDeformations = True Then
                ReDim ptsf(2)

                'compute the colormap

                'set the colormap if required
                Select Case ShowResult
                    Case Results.None
                        'create a dummy colormap
                        colorMap = New ColorMap(1, 0)
                    Case Results.SigmaX
                        colorMap = New ColorMap(SigmaXRange.Max, SigmaXRange.Min)
                    Case Results.SigmaY
                        colorMap = New ColorMap(SigmaYRange.Max, SigmaYRange.Min)
                    Case Results.TauXY
                        colorMap = New ColorMap(TauXYRange.Max, TauXYRange.Min)
                    Case Results.EpsilonX
                        colorMap = New ColorMap(EpsilonXRange.Max, EpsilonXRange.Min)
                    Case Results.EpsilonY
                        colorMap = New ColorMap(EpsilonYRange.Max, EpsilonYRange.Min)
                    Case Results.GammaXY
                        colorMap = New ColorMap(GammaXYRange.Max, GammaXYRange.Min)
                    Case Else
                        'just for completeness.. otherwise this block wont ever execute
                        'create a dummy colormap
                        colorMap = New ColorMap(1, 0)
                End Select

                For i = 0 To NElements - 1

                    'ecolor depends on the value of result to be shown

                    Select Case ShowResult
                        Case Results.None
                            eColor = eleColorDeformed
                        Case Results.SigmaX
                            eColor = colorMap.getColor(Elements(i).Stresses(0))
                        Case Results.SigmaY
                            eColor = colorMap.getColor(Elements(i).Stresses(1))
                        Case Results.TauXY
                            eColor = colorMap.getColor(Elements(i).Stresses(2))
                        Case Results.EpsilonX
                            eColor = colorMap.getColor(Elements(i).Strains(0))
                        Case Results.EpsilonY
                            eColor = colorMap.getColor(Elements(i).Strains(1))
                        Case Results.GammaXY
                            eColor = colorMap.getColor(Elements(i).Strains(2))
                        Case Else
                            eColor = eleColor
                    End Select


                    n1 = Elements(i).Node1 - 1
                    n2 = Elements(i).Node2 - 1
                    n3 = Elements(i).Node3 - 1

                    ptsf(0) = New PointF(CSng((Nodes(n1).x * zoom + getDeformationX(n1) * DeformationZoom) + shiftx),
                                         CSng((-Nodes(n1).y * zoom - getDeformationY(n1) * DeformationZoom) + shifty))
                    ptsf(1) = New PointF(CSng((Nodes(n2).x * zoom + getDeformationX(n2) * DeformationZoom) + shiftx),
                                         CSng((-Nodes(n2).y * zoom - getDeformationY(n2) * DeformationZoom) + shifty))
                    ptsf(2) = New PointF(CSng((Nodes(n3).x * zoom + getDeformationX(n3) * DeformationZoom) + shiftx),
                                         CSng((-Nodes(n3).y * zoom - getDeformationY(n3) * DeformationZoom) + shifty))

                    'Draw the element
                    gr.FillPolygon(New SolidBrush(eColor), ptsf)
                    If ShowElementsOnDeformedShape = True Then
                        gr.DrawPolygon(New Pen(Color.Black), ptsf)
                    End If

                Next

                'Draw Colormap at right top of the picturebox
                'only if required
                If ShowResult <> Results.None Then
                    Dim cmap As New ColorMap(250, 0)
                    Dim c As Color
                    Dim x1, x2, y As Integer
                    y = 50
                    x1 = bmp.Width - 100
                    x2 = x1 + 50
                    For i = 250 To 0 Step -1
                        c = cmap.getColor(i)
                        y += 1
                        gr.DrawLine(New Pen(c), New Point(x1, y), New Point(x2, y))
                    Next
                    'write the max, min and mid values
                    'max

                    Dim strWidth As Integer
                    Dim MaxWidth, MinWidth, AvgWidth As Single
                    Dim Max, Min, Avg As String
                    Dim barFont As New Font("Arial", 9)
                    Dim barBrush As New SolidBrush(Color.Black)
                    Max = colorMap.Max.ToString
                    Min = colorMap.Min.ToString
                    Avg = ((colorMap.Max + colorMap.Min) / 2).ToString
                    MaxWidth = gr.MeasureString(Max, barFont).Width
                    MinWidth = gr.MeasureString(Min, barFont).Width
                    AvgWidth = gr.MeasureString(Avg, barFont).Width

                    strWidth = CInt(Math.Max(Math.Max(MaxWidth, MinWidth), AvgWidth)) + 10

                    gr.DrawLine(Pens.Black, x1, 50, x1 - 5, 50)
                    gr.DrawString(Max, barFont, barBrush, New PointF(x1 - strWidth, 30))

                    gr.DrawLine(Pens.Black, x1, 175, x1 - 5, 175)
                    gr.DrawString(Avg, barFont, barBrush, New PointF(x1 - strWidth, 155))

                    gr.DrawLine(Pens.Black, x1, 300, x1 - 5, 300)
                    gr.DrawString(Min, barFont, barBrush, New PointF(x1 - strWidth, 320))
                End If


            End If
        End If
        pbModel.Refresh()
    End Sub


    Private Function getDeformationX(Node As Integer) As Double
        Dim dof As Integer = getDOFx(Node)
        Return Deformations(dof)
    End Function
    Private Function getDeformationY(Node As Integer) As Double
        Dim dof As Integer = getDOFy(Node)
        Return Deformations(dof)
    End Function


    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        If MsgBox("Are you sure you want to exit?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
            End
        End If
    End Sub

    Private Sub DrawHArrow(px As Integer, py As Integer, c As Color, leftward As Boolean)
        Dim p As New Pen(c)
        Dim size As Integer = 15
        If leftward = False Then
            'rightward arrow
            gr.DrawLine(p, New Point(px, py), New Point(px - size, py))
            gr.DrawLine(p, New Point(px, py), New Point(CInt(px - size / 4), CInt(py - size / 4)))
            gr.DrawLine(p, New Point(px, py), New Point(CInt(px - size / 4), CInt(py + size / 4)))
        Else
            'leftward arrow
            gr.DrawLine(p, New Point(px, py), New Point(px + size, py))
            gr.DrawLine(p, New Point(px, py), New Point(CInt(px + size / 4), CInt(py - size / 4)))
            gr.DrawLine(p, New Point(px, py), New Point(CInt(px + size / 4), CInt(py + size / 4)))
        End If


    End Sub

    Private Sub AnalyseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnalyseToolStripMenuItem.Click

        'check if there is a proper model
        If NElements <= 0 OrElse NNodes <= 0 Then
            'there are no elements defined
            MsgBox("Please open a proper finite element model")
            Return
        End If

        'perform analysis using the finite elemeent method
        Analyse()

        'show the output
        DrawModel()
    End Sub

    Private Sub SetDeformationZoomToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SetDeformationZoomToolStripMenuItem.Click
        Dim ret As String = DeformationZoom.ToString

        ret = InputBox("Set Deformation Zoom:", "Deformation Zoom", ret)
        If ret.Length = 0 Then Return
        While IsNumeric(ret) = False
            ret = InputBox("Set Deformation Zoom:", "Deformation Zoom", ret)
            If ret.Length = 0 Then Return
        End While
        DeformationZoom = Double.Parse(ret)
        DrawModel()
    End Sub



    Private Sub ShowNodeNumbersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowNodeNumbersToolStripMenuItem.Click
        ShowNodeNumbers = Not ShowNodeNumbers
        DrawModel()
    End Sub

    Private Sub ResultsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResultsToolStripMenuItem.Click
        Dim d As New dlgResults
        Dim sb As New System.Text.StringBuilder
        Dim s As String
        Dim dof As Integer


        sb.AppendLine("Model Statistics:")
        sb.AppendLine("Number of Nodes: " & NNodes.ToString)
        sb.AppendLine("Number of Elements: " & NElements.ToString)
        sb.AppendLine("Number of Variables: " & (NNodes * 2).ToString)
        sb.AppendLine("")

        'Deformations
        sb.AppendLine("Deformations:")
        sb.AppendLine("Node" & vbTab & vbTab & "Ux" & vbTab & vbTab & "Uy")
        For i = 0 To NNodes - 1
            s = (i + 1).ToString & vbTab & vbTab
            dof = getDOFx(i)
            s = s & Deformations(dof).ToString & vbTab & vbTab
            s = s & Deformations(dof + 1).ToString

            sb.AppendLine(s)
        Next

        'Element Stresses and Strains
        sb.AppendLine("Stresses and Strains:")
        sb.AppendLine("Element" & vbTab & vbTab & "sx" & vbTab & vbTab & "sy" & vbTab & vbTab & "txy" &
        vbTab & vbTab & "ex" & vbTab & vbTab & "ey" & vbTab & vbTab & "gamma_xy")
        For i = 0 To NElements - 1
            s = (i + 1).ToString & vbTab & vbTab

            s = s & Elements(i).Stresses(0).ToString & vbTab & vbTab
            s = s & Elements(i).Stresses(1).ToString & vbTab & vbTab
            s = s & Elements(i).Stresses(2).ToString & vbTab & vbTab

            s = s & Elements(i).Strains(0).ToString & vbTab & vbTab
            s = s & Elements(i).Strains(1).ToString & vbTab & vbTab
            s = s & Elements(i).Strains(2).ToString

            sb.AppendLine(s)
        Next

        'Write maximum displacements
        Dim uxMax As Double = Double.MinValue
        Dim uyMax As Double = Double.MinValue
        Dim uxMin As Double = Double.MaxValue
        Dim uyMin As Double = Double.MaxValue

        For i = 0 To Deformations.Length - 1 Step 2
            If uxMax < Deformations(i) Then uxMax = Deformations(i)
            If uxMin > Deformations(i) Then uxMin = Deformations(i)
            If uyMax < Deformations(i + 1) Then uyMax = Deformations(i + 1)
            If uyMin > Deformations(i + 1) Then uyMin = Deformations(i + 1)
        Next

        sb.AppendLine("")
        sb.AppendLine("Maximum Displacement in X direction = " & uxMax.ToString)
        sb.AppendLine("Minimum Displacement in X direction = " & uxMin.ToString)
        sb.AppendLine("Maximum Displacement in Y direction = " & uyMax.ToString)
        sb.AppendLine("Minimum Displacement in Y direction = " & uyMin.ToString)

        sb.AppendLine("")
        sb.AppendLine("Output generated by btFEM at " & DateTime.Now.ToLongDateString)


        d.txtResults.Text = sb.ToString


        d.ShowDialog()

    End Sub

    Private Sub ShowModelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowModelToolStripMenuItem.Click
        ShowModel = Not ShowModel
        DrawModel()
    End Sub

    Private Sub ShowElementsOnDeformedShapeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowElementsOnDeformedShapeToolStripMenuItem.Click
        ShowElementsOnDeformedShape = Not ShowElementsOnDeformedShape
        DrawModel()
    End Sub

    Private Sub SaveImageAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveImageAsToolStripMenuItem.Click
        Dim d As New SaveFileDialog
        d.Title = "Save Image As"
        d.Filter = "PNG Files|*.png|JPEG Files|*.jpg|BMP Files|*.bmp|TIFF FIles|*.tiff|All Files|*.*"
        If d.ShowDialog = DialogResult.OK Then
            Try

                Select Case d.FilterIndex
                    Case 1 'png
                        bmp.Save(d.FileName, Imaging.ImageFormat.Png)
                    Case 2 'jpg
                        bmp.Save(d.FileName, Imaging.ImageFormat.Jpeg)
                    Case 3 'bmp
                        bmp.Save(d.FileName, Imaging.ImageFormat.Bmp)
                    Case 4 'tiff
                        bmp.Save(d.FileName, Imaging.ImageFormat.Tiff)
                    Case Else
                        bmp.Save(d.FileName, Imaging.ImageFormat.Bmp)
                End Select
                MsgBox("Image saved successfully to" & vbCrLf & d.FileName)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub ShowDeformationsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowDeformationsToolStripMenuItem.Click
        ShowDeformations = Not ShowDeformations
        DrawModel()
    End Sub

    Private Sub DrawVArrow(px As Integer, py As Integer, c As Color, downward As Boolean)
        Dim p As New Pen(c)
        Dim size As Integer = 15
        If downward = False Then
            'upward arrow
            gr.DrawLine(p, New Point(px, py), New Point(px, py + size))
            gr.DrawLine(p, New Point(px, py), New Point(CInt(px - size / 4), CInt(py + size / 4)))
            gr.DrawLine(p, New Point(px, py), New Point(CInt(px + size / 4), CInt(py + size / 4)))
        Else
            'downward arrow
            gr.DrawLine(p, New Point(px, py), New Point(px, py - size))
            gr.DrawLine(p, New Point(px, py), New Point(CInt(px - size / 4), CInt(py - size / 4)))
            gr.DrawLine(p, New Point(px, py), New Point(CInt(px + size / 4), CInt(py - size / 4)))
        End If


    End Sub



    Private Sub HandleShowResultClick(sender As Object, e As EventArgs) Handles NoneToolStripMenuItem.Click,
            SigmaXToolStripMenuItem.Click, SigmaYToolStripMenuItem.Click, TauXYToolStripMenuItem.Click,
            EpsilonXToolStripMenuItem.Click, EpsilonYToolStripMenuItem.Click, GammaXYToolStripMenuItem.Click

        ShowResult = Results.None

        If sender.Equals(NoneToolStripMenuItem) Then
            ShowResult = Results.None
        ElseIf sender.Equals(SigmaXToolStripMenuItem) Then
            ShowResult = Results.SigmaX
        ElseIf sender.Equals(SigmaYToolStripMenuItem) Then
            ShowResult = Results.SigmaY
        ElseIf sender.Equals(TauXYToolStripMenuItem) Then
            ShowResult = Results.TauXY
        ElseIf sender.Equals(EpsilonXToolStripMenuItem) Then
            ShowResult = Results.EpsilonX
        ElseIf sender.Equals(EpsilonYToolStripMenuItem) Then
            ShowResult = Results.EpsilonY
        ElseIf sender.Equals(GammaXYToolStripMenuItem) Then
            ShowResult = Results.GammaXY
        End If

        DrawModel()
    End Sub

    Private Sub pbModel_Paint(sender As Object, e As PaintEventArgs) Handles pbModel.Paint
        If bmp IsNot Nothing Then
            e.Graphics.DrawImage(bmp, 0, 0)
        End If
    End Sub



    Private Sub pbModel_Resize(sender As Object, e As EventArgs) Handles pbModel.Resize
        If bmp IsNot Nothing Then
            bmp.Dispose()
        End If
        If gr IsNot Nothing Then
            gr.Dispose()
        End If
        bmp = New Bitmap(pbModel.Width, pbModel.Height, pbModel.CreateGraphics)
        gr = Graphics.FromImage(bmp)
        DrawModel()
    End Sub

    Private Sub pbModel_MouseDown(sender As Object, e As MouseEventArgs) Handles pbModel.MouseDown
        If e.Button = MouseButtons.Left Then
            'zoom in
            zoom = zoom * 1.1
        ElseIf e.Button = MouseButtons.Right Then
            zoom = zoom / 1.1
        End If
        DrawModel()
    End Sub

    Private Sub frmbtFem_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If MsgBox("Are you sure you want to exit?", MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then
            e.Cancel = True
        End If
    End Sub

    Private Sub ShowToolStripMenuItem_DropDownOpening(sender As Object, e As EventArgs) Handles ShowToolStripMenuItem.DropDownOpening

        NoneToolStripMenuItem.Checked = False
        SigmaXToolStripMenuItem.Checked = False
        SigmaYToolStripMenuItem.Checked = False
        TauXYToolStripMenuItem.Checked = False
        EpsilonXToolStripMenuItem.Checked = False
        EpsilonYToolStripMenuItem.Checked = False
        GammaXYToolStripMenuItem.Checked = False


        Select Case ShowResult
            Case Results.None
                NoneToolStripMenuItem.Checked = True
            Case Results.SigmaX
                SigmaXToolStripMenuItem.Checked = True
            Case Results.SigmaY
                SigmaYToolStripMenuItem.Checked = True
            Case Results.TauXY
                TauXYToolStripMenuItem.Checked = True
            Case Results.EpsilonX
                EpsilonXToolStripMenuItem.Checked = True
            Case Results.EpsilonY
                EpsilonYToolStripMenuItem.Checked = True
            Case Results.GammaXY
                GammaXYToolStripMenuItem.Checked = True
        End Select

    End Sub

    Private Sub ModelToolStripMenuItem_DropDownOpening(sender As Object, e As EventArgs) Handles ModelToolStripMenuItem.DropDownOpening
        ShowModelToolStripMenuItem.Checked = ShowModel
        ShowDeformationsToolStripMenuItem.Checked = ShowDeformations
        ShowNodeNumbersToolStripMenuItem.Checked = ShowNodeNumbers
        ShowElementsOnDeformedShapeToolStripMenuItem.Checked = ShowElementsOnDeformedShape
    End Sub
End Class
