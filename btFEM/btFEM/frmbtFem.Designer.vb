<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmbtFem
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveImageAsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ModelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AnalyseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ResultsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ShowNodeNumbersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SetDeformationZoomToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ShowModelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ShowDeformationsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ShowToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NoneToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem4 = New System.Windows.Forms.ToolStripSeparator()
        Me.SigmaXToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SigmaYToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TauXYToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem3 = New System.Windows.Forms.ToolStripSeparator()
        Me.EpsilonXToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EpsilonYToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GammaXYToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ShowElementsOnDeformedShapeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.pbModel = New System.Windows.Forms.PictureBox()
        Me.hs = New System.Windows.Forms.HScrollBar()
        Me.vs = New System.Windows.Forms.VScrollBar()
        Me.btnResetZoom = New System.Windows.Forms.Button()
        Me.ToolStripMenuItem5 = New System.Windows.Forms.ToolStripSeparator()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.pbModel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.ModelToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(708, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenToolStripMenuItem, Me.SaveImageAsToolStripMenuItem, Me.ToolStripMenuItem1, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'OpenToolStripMenuItem
        '
        Me.OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
        Me.OpenToolStripMenuItem.Size = New System.Drawing.Size(159, 22)
        Me.OpenToolStripMenuItem.Text = "&Open..."
        '
        'SaveImageAsToolStripMenuItem
        '
        Me.SaveImageAsToolStripMenuItem.Name = "SaveImageAsToolStripMenuItem"
        Me.SaveImageAsToolStripMenuItem.Size = New System.Drawing.Size(159, 22)
        Me.SaveImageAsToolStripMenuItem.Text = "Save &Image As..."
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(156, 6)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(159, 22)
        Me.ExitToolStripMenuItem.Text = "E&xit"
        '
        'ModelToolStripMenuItem
        '
        Me.ModelToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AnalyseToolStripMenuItem, Me.ToolStripMenuItem2, Me.ResultsToolStripMenuItem, Me.ShowNodeNumbersToolStripMenuItem, Me.SetDeformationZoomToolStripMenuItem, Me.ShowModelToolStripMenuItem, Me.ShowDeformationsToolStripMenuItem, Me.ShowElementsOnDeformedShapeToolStripMenuItem, Me.ToolStripMenuItem5, Me.ShowToolStripMenuItem})
        Me.ModelToolStripMenuItem.Name = "ModelToolStripMenuItem"
        Me.ModelToolStripMenuItem.Size = New System.Drawing.Size(53, 20)
        Me.ModelToolStripMenuItem.Text = "&Model"
        '
        'AnalyseToolStripMenuItem
        '
        Me.AnalyseToolStripMenuItem.Name = "AnalyseToolStripMenuItem"
        Me.AnalyseToolStripMenuItem.Size = New System.Drawing.Size(262, 22)
        Me.AnalyseToolStripMenuItem.Text = "&Analyse"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(259, 6)
        '
        'ResultsToolStripMenuItem
        '
        Me.ResultsToolStripMenuItem.Name = "ResultsToolStripMenuItem"
        Me.ResultsToolStripMenuItem.Size = New System.Drawing.Size(262, 22)
        Me.ResultsToolStripMenuItem.Text = "&Results..."
        '
        'ShowNodeNumbersToolStripMenuItem
        '
        Me.ShowNodeNumbersToolStripMenuItem.Name = "ShowNodeNumbersToolStripMenuItem"
        Me.ShowNodeNumbersToolStripMenuItem.Size = New System.Drawing.Size(262, 22)
        Me.ShowNodeNumbersToolStripMenuItem.Text = "Show Node Numbers"
        '
        'SetDeformationZoomToolStripMenuItem
        '
        Me.SetDeformationZoomToolStripMenuItem.Name = "SetDeformationZoomToolStripMenuItem"
        Me.SetDeformationZoomToolStripMenuItem.Size = New System.Drawing.Size(262, 22)
        Me.SetDeformationZoomToolStripMenuItem.Text = "Set Deformation Zoom..."
        '
        'ShowModelToolStripMenuItem
        '
        Me.ShowModelToolStripMenuItem.Name = "ShowModelToolStripMenuItem"
        Me.ShowModelToolStripMenuItem.Size = New System.Drawing.Size(262, 22)
        Me.ShowModelToolStripMenuItem.Text = "Show Model"
        '
        'ShowDeformationsToolStripMenuItem
        '
        Me.ShowDeformationsToolStripMenuItem.Checked = True
        Me.ShowDeformationsToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ShowDeformationsToolStripMenuItem.Name = "ShowDeformationsToolStripMenuItem"
        Me.ShowDeformationsToolStripMenuItem.Size = New System.Drawing.Size(262, 22)
        Me.ShowDeformationsToolStripMenuItem.Text = "Show Deformations"
        '
        'ShowToolStripMenuItem
        '
        Me.ShowToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NoneToolStripMenuItem, Me.ToolStripMenuItem4, Me.SigmaXToolStripMenuItem, Me.SigmaYToolStripMenuItem, Me.TauXYToolStripMenuItem, Me.ToolStripMenuItem3, Me.EpsilonXToolStripMenuItem, Me.EpsilonYToolStripMenuItem, Me.GammaXYToolStripMenuItem})
        Me.ShowToolStripMenuItem.Name = "ShowToolStripMenuItem"
        Me.ShowToolStripMenuItem.Size = New System.Drawing.Size(262, 22)
        Me.ShowToolStripMenuItem.Text = "Show Results"
        '
        'NoneToolStripMenuItem
        '
        Me.NoneToolStripMenuItem.Name = "NoneToolStripMenuItem"
        Me.NoneToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.NoneToolStripMenuItem.Text = "None"
        '
        'ToolStripMenuItem4
        '
        Me.ToolStripMenuItem4.Name = "ToolStripMenuItem4"
        Me.ToolStripMenuItem4.Size = New System.Drawing.Size(130, 6)
        '
        'SigmaXToolStripMenuItem
        '
        Me.SigmaXToolStripMenuItem.Name = "SigmaXToolStripMenuItem"
        Me.SigmaXToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.SigmaXToolStripMenuItem.Text = "Sigma X"
        '
        'SigmaYToolStripMenuItem
        '
        Me.SigmaYToolStripMenuItem.Name = "SigmaYToolStripMenuItem"
        Me.SigmaYToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.SigmaYToolStripMenuItem.Text = "Sigma Y"
        '
        'TauXYToolStripMenuItem
        '
        Me.TauXYToolStripMenuItem.Name = "TauXYToolStripMenuItem"
        Me.TauXYToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.TauXYToolStripMenuItem.Text = "Tau XY"
        '
        'ToolStripMenuItem3
        '
        Me.ToolStripMenuItem3.Name = "ToolStripMenuItem3"
        Me.ToolStripMenuItem3.Size = New System.Drawing.Size(130, 6)
        '
        'EpsilonXToolStripMenuItem
        '
        Me.EpsilonXToolStripMenuItem.Name = "EpsilonXToolStripMenuItem"
        Me.EpsilonXToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.EpsilonXToolStripMenuItem.Text = "Epsilon X"
        '
        'EpsilonYToolStripMenuItem
        '
        Me.EpsilonYToolStripMenuItem.Name = "EpsilonYToolStripMenuItem"
        Me.EpsilonYToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.EpsilonYToolStripMenuItem.Text = "Epsilon Y"
        '
        'GammaXYToolStripMenuItem
        '
        Me.GammaXYToolStripMenuItem.Name = "GammaXYToolStripMenuItem"
        Me.GammaXYToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.GammaXYToolStripMenuItem.Text = "Gamma XY"
        '
        'ShowElementsOnDeformedShapeToolStripMenuItem
        '
        Me.ShowElementsOnDeformedShapeToolStripMenuItem.Name = "ShowElementsOnDeformedShapeToolStripMenuItem"
        Me.ShowElementsOnDeformedShapeToolStripMenuItem.Size = New System.Drawing.Size(262, 22)
        Me.ShowElementsOnDeformedShapeToolStripMenuItem.Text = "Show Elements on Deformed Shape"
        '
        'pbModel
        '
        Me.pbModel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbModel.BackColor = System.Drawing.Color.White
        Me.pbModel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pbModel.Location = New System.Drawing.Point(12, 27)
        Me.pbModel.Name = "pbModel"
        Me.pbModel.Size = New System.Drawing.Size(665, 384)
        Me.pbModel.TabIndex = 1
        Me.pbModel.TabStop = False
        '
        'hs
        '
        Me.hs.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.hs.Location = New System.Drawing.Point(12, 414)
        Me.hs.Maximum = 100000
        Me.hs.Minimum = -100000
        Me.hs.Name = "hs"
        Me.hs.Size = New System.Drawing.Size(665, 19)
        Me.hs.TabIndex = 2
        '
        'vs
        '
        Me.vs.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.vs.Location = New System.Drawing.Point(680, 27)
        Me.vs.Maximum = 100000
        Me.vs.Minimum = -100000
        Me.vs.Name = "vs"
        Me.vs.Size = New System.Drawing.Size(19, 384)
        Me.vs.TabIndex = 3
        '
        'btnResetZoom
        '
        Me.btnResetZoom.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnResetZoom.BackColor = System.Drawing.Color.Black
        Me.btnResetZoom.Location = New System.Drawing.Point(680, 414)
        Me.btnResetZoom.Name = "btnResetZoom"
        Me.btnResetZoom.Size = New System.Drawing.Size(19, 19)
        Me.btnResetZoom.TabIndex = 4
        Me.btnResetZoom.UseVisualStyleBackColor = False
        '
        'ToolStripMenuItem5
        '
        Me.ToolStripMenuItem5.Name = "ToolStripMenuItem5"
        Me.ToolStripMenuItem5.Size = New System.Drawing.Size(259, 6)
        '
        'frmbtFem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(708, 442)
        Me.Controls.Add(Me.btnResetZoom)
        Me.Controls.Add(Me.vs)
        Me.Controls.Add(Me.hs)
        Me.Controls.Add(Me.pbModel)
        Me.Controls.Add(Me.MenuStrip1)
        Me.DoubleBuffered = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmbtFem"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.Text = "Finite Element Method Framework"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.pbModel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OpenToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripSeparator
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents pbModel As PictureBox
    Friend WithEvents ModelToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ResultsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As ToolStripSeparator
    Friend WithEvents hs As HScrollBar
    Friend WithEvents vs As VScrollBar
    Friend WithEvents btnResetZoom As Button
    Friend WithEvents AnalyseToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SetDeformationZoomToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ShowDeformationsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ShowToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SigmaXToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SigmaYToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TauXYToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem3 As ToolStripSeparator
    Friend WithEvents EpsilonXToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents EpsilonYToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents GammaXYToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents NoneToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem4 As ToolStripSeparator
    Friend WithEvents ShowNodeNumbersToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ShowModelToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ShowElementsOnDeformedShapeToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveImageAsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem5 As ToolStripSeparator
End Class
