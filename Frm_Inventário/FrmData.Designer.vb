<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmData
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmData))
        Me.DTP = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LblDataAtual = New System.Windows.Forms.Label()
        Me.BtnData = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'DTP
        '
        Me.DTP.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTP.Location = New System.Drawing.Point(70, 30)
        Me.DTP.Name = "DTP"
        Me.DTP.Size = New System.Drawing.Size(97, 20)
        Me.DTP.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(25, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Label1"
        '
        'LblDataAtual
        '
        Me.LblDataAtual.AutoSize = True
        Me.LblDataAtual.Location = New System.Drawing.Point(36, 82)
        Me.LblDataAtual.Name = "LblDataAtual"
        Me.LblDataAtual.Size = New System.Drawing.Size(39, 13)
        Me.LblDataAtual.TabIndex = 2
        Me.LblDataAtual.Text = "Label2"
        '
        'BtnData
        '
        Me.BtnData.Location = New System.Drawing.Point(119, 77)
        Me.BtnData.Name = "BtnData"
        Me.BtnData.Size = New System.Drawing.Size(75, 23)
        Me.BtnData.TabIndex = 3
        Me.BtnData.Text = "Button1"
        Me.BtnData.UseVisualStyleBackColor = True
        '
        'FrmData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(345, 128)
        Me.Controls.Add(Me.BtnData)
        Me.Controls.Add(Me.LblDataAtual)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DTP)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmData"
        Me.Text = "Data Expiração"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DTP As DateTimePicker
    Friend WithEvents Label1 As Label
    Friend WithEvents LblDataAtual As Label
    Friend WithEvents BtnData As Button
End Class
