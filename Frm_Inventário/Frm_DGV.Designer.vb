<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_DGV
    Inherits System.Windows.Forms.Form

    'Descartar substituições de formulário para limpar a lista de componentes.
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

    'Exigido pelo Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'OBSERVAÇÃO: o procedimento a seguir é exigido pelo Windows Form Designer
    'Pode ser modificado usando o Windows Form Designer.  
    'Não o modifique usando o editor de códigos.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.LblLinhas = New System.Windows.Forms.Label()
        Me.DGV_Consulta = New System.Windows.Forms.DataGridView()
        Me.CMS_DGV = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ExcluirDadosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CopiarDadosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.DGV_Consulta, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CMS_DGV.SuspendLayout()
        Me.SuspendLayout()
        '
        'LblLinhas
        '
        Me.LblLinhas.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblLinhas.Location = New System.Drawing.Point(744, 435)
        Me.LblLinhas.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblLinhas.Name = "LblLinhas"
        Me.LblLinhas.Size = New System.Drawing.Size(155, 19)
        Me.LblLinhas.TabIndex = 11
        Me.LblLinhas.Text = "Total de Registros: 0"
        Me.LblLinhas.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'DGV_Consulta
        '
        Me.DGV_Consulta.AllowUserToAddRows = False
        Me.DGV_Consulta.AllowUserToDeleteRows = False
        Me.DGV_Consulta.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DGV_Consulta.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGV_Consulta.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV_Consulta.Location = New System.Drawing.Point(-28, 0)
        Me.DGV_Consulta.Name = "DGV_Consulta"
        Me.DGV_Consulta.ReadOnly = True
        Me.DGV_Consulta.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DGV_Consulta.Size = New System.Drawing.Size(927, 422)
        Me.DGV_Consulta.TabIndex = 10
        '
        'CMS_DGV
        '
        Me.CMS_DGV.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.CMS_DGV.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExcluirDadosToolStripMenuItem, Me.CopiarDadosToolStripMenuItem})
        Me.CMS_DGV.Name = "CMS_DGV"
        Me.CMS_DGV.Size = New System.Drawing.Size(146, 48)
        '
        'ExcluirDadosToolStripMenuItem
        '
        Me.ExcluirDadosToolStripMenuItem.Name = "ExcluirDadosToolStripMenuItem"
        Me.ExcluirDadosToolStripMenuItem.Size = New System.Drawing.Size(145, 22)
        Me.ExcluirDadosToolStripMenuItem.Text = "Excluir Dados"
        '
        'CopiarDadosToolStripMenuItem
        '
        Me.CopiarDadosToolStripMenuItem.Name = "CopiarDadosToolStripMenuItem"
        Me.CopiarDadosToolStripMenuItem.Size = New System.Drawing.Size(145, 22)
        Me.CopiarDadosToolStripMenuItem.Text = "Copiar Dados"
        '
        'Frm_DGV
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(911, 450)
        Me.Controls.Add(Me.LblLinhas)
        Me.Controls.Add(Me.DGV_Consulta)
        Me.Name = "Frm_DGV"
        Me.Text = "Consulta"
        CType(Me.DGV_Consulta, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CMS_DGV.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents LblLinhas As Label
    Friend WithEvents DGV_Consulta As DataGridView
    Friend WithEvents CMS_DGV As ContextMenuStrip
    Friend WithEvents ExcluirDadosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CopiarDadosToolStripMenuItem As ToolStripMenuItem
End Class
