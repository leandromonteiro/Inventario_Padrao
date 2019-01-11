Public Class Frm_DGV
    Dim I_E As New Inventário_Excel
    Public Alterado As Boolean
    Dim Linha_ID As String
    Dim Texto_Foto As String

    Private Sub Frm_DGV_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Frm_Inventário.Show()
    End Sub

    Private Sub Frm_DGV_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Consultar Dados
        I_E.Consulta_Grid(DGV_Consulta)
        LblLinhas.Text = "Total de Registros: " & DGV_Consulta.Rows.Count
    End Sub

    Private Sub DGV_Consulta_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_Consulta.CellDoubleClick
        If e.RowIndex = -1 Then
            Exit Sub
        End If
        'Limpar Array Fotos
        Try
            'Frm_Inventário.Fotos_Array.Clear()
            'Frm_Inventário.PictureBox_Consulta.ImageLocation = ""
            'Frm_Inventário.Add_Fotos_Array = 0

            'Frm_Inventário.TxtSeq_Civil.Text = DGV_Consulta.Item(0, e.RowIndex).Value
            'Frm_Inventário.TxtSeq_Desc.Text = DGV_Consulta.Item(0, e.RowIndex).Value
            'Frm_Inventário.TxtSeq_Local.Text = DGV_Consulta.Item(0, e.RowIndex).Value
            'Frm_Inventário.TxtLocal.Text = DGV_Consulta.Item(2, e.RowIndex).Value
            'Frm_Inventário.TxtODI.Text = DGV_Consulta.Item(3, e.RowIndex).Value
            'Frm_Inventário.TI = DGV_Consulta.Item(4, e.RowIndex).Value
            'Frm_Inventário.CmbTI.Text = DGV_Consulta.Item(5, e.RowIndex).Value
            'Frm_Inventário.TxtTagNovo.Text = DGV_Consulta.Item(6, e.RowIndex).Value
            'Frm_Inventário.TUC = DGV_Consulta.Item(7, e.RowIndex).Value
            'Frm_Inventário.CmbTUC.Text = DGV_Consulta.Item(8, e.RowIndex).Value
            'Frm_Inventário.A1 = DGV_Consulta.Item(9, e.RowIndex).Value
            'Frm_Inventário.CmbA1.Text = DGV_Consulta.Item(10, e.RowIndex).Value
            'Frm_Inventário.UAR = DGV_Consulta.Item(11, e.RowIndex).Value
            'Frm_Inventário.CmbUAR.Text = DGV_Consulta.Item(12, e.RowIndex).Value
            'Frm_Inventário.A2 = IIf(IsDBNull(DGV_Consulta.Item(13, e.RowIndex).Value), "", DGV_Consulta.Item(13, e.RowIndex).Value)
            'Frm_Inventário.CmbA2.Text = IIf(IsDBNull(DGV_Consulta.Item(14, e.RowIndex).Value), "", DGV_Consulta.Item(14, e.RowIndex).Value)
            'Frm_Inventário.A3 = IIf(IsDBNull(DGV_Consulta.Item(15, e.RowIndex).Value), "", DGV_Consulta.Item(15, e.RowIndex).Value)
            'Frm_Inventário.CmbA3.Text = IIf(IsDBNull(DGV_Consulta.Item(16, e.RowIndex).Value), "", DGV_Consulta.Item(16, e.RowIndex).Value)
            'Frm_Inventário.A4 = IIf(IsDBNull(DGV_Consulta.Item(17, e.RowIndex).Value), "", DGV_Consulta.Item(17, e.RowIndex).Value)
            'Frm_Inventário.CmbA4.Text = IIf(IsDBNull(DGV_Consulta.Item(18, e.RowIndex).Value), "", DGV_Consulta.Item(18, e.RowIndex).Value)
            'Frm_Inventário.A5 = IIf(IsDBNull(DGV_Consulta.Item(19, e.RowIndex).Value), "", DGV_Consulta.Item(19, e.RowIndex).Value)
            'Frm_Inventário.CmbA5.Text = IIf(IsDBNull(DGV_Consulta.Item(20, e.RowIndex).Value), "", DGV_Consulta.Item(20, e.RowIndex).Value)
            'Frm_Inventário.A6 = IIf(IsDBNull(DGV_Consulta.Item(21, e.RowIndex).Value), "", DGV_Consulta.Item(21, e.RowIndex).Value)
            'Frm_Inventário.CmbA6.Text = IIf(IsDBNull(DGV_Consulta.Item(22, e.RowIndex).Value), "", DGV_Consulta.Item(22, e.RowIndex).Value)
            'Frm_Inventário.CM1 = DGV_Consulta.Item(23, e.RowIndex).Value
            'Frm_Inventário.CmbCm1.Text = DGV_Consulta.Item(24, e.RowIndex).Value
            'Frm_Inventário.CM2 = DGV_Consulta.Item(25, e.RowIndex).Value
            'Frm_Inventário.CmbTagAntigo.Text = DGV_Consulta.Item(26, e.RowIndex).Value
            'Frm_Inventário.CM3 = DGV_Consulta.Item(27, e.RowIndex).Value
            'Frm_Inventário.CmbLocalFisico.Text = DGV_Consulta.Item(28, e.RowIndex).Value
            'Frm_Inventário.TxtDetalhada.Text = DGV_Consulta.Item(29, e.RowIndex).Value
            'Frm_Inventário.TxtFabricante.Text = IIf(IsDBNull(DGV_Consulta.Item(30, e.RowIndex).Value), "", DGV_Consulta.Item(30, e.RowIndex).Value)
            'Frm_Inventário.TxtModelo.Text = IIf(IsDBNull(DGV_Consulta.Item(31, e.RowIndex).Value), "", DGV_Consulta.Item(31, e.RowIndex).Value)
            'Frm_Inventário.TxtSerie.Text = IIf(IsDBNull(DGV_Consulta.Item(32, e.RowIndex).Value), "", DGV_Consulta.Item(32, e.RowIndex).Value)
            'Frm_Inventário.TxtObs.Text = IIf(IsDBNull(DGV_Consulta.Item(33, e.RowIndex).Value), "", DGV_Consulta.Item(33, e.RowIndex).Value)
            'Frm_Inventário.TxtQtd.Text = DGV_Consulta.Item(34, e.RowIndex).Value
            'Frm_Inventário.CmbUm.Text = DGV_Consulta.Item(35, e.RowIndex).Value
            'Frm_Inventário.CmbAno.Text = IIf(IsDBNull(DGV_Consulta.Item(36, e.RowIndex).Value), "", DGV_Consulta.Item(36, e.RowIndex).Value)
            'Frm_Inventário.CmbMes.Text = IIf(IsDBNull(DGV_Consulta.Item(37, e.RowIndex).Value), "", DGV_Consulta.Item(37, e.RowIndex).Value)
            'Frm_Inventário.CmbDia.Text = IIf(IsDBNull(DGV_Consulta.Item(38, e.RowIndex).Value), "", DGV_Consulta.Item(38, e.RowIndex).Value)
            'Frm_Inventário.CmbStatus.Text = DGV_Consulta.Item(39, e.RowIndex).Value
            'Frm_Inventário.CmbEstado.Text = DGV_Consulta.Item(40, e.RowIndex).Value
            'Frm_Inventário.TxtAltura.Text = IIf(IsDBNull(DGV_Consulta.Item(41, e.RowIndex).Value), "", DGV_Consulta.Item(41, e.RowIndex).Value)
            'Frm_Inventário.TxtLargura.Text = IIf(IsDBNull(DGV_Consulta.Item(42, e.RowIndex).Value), "", DGV_Consulta.Item(42, e.RowIndex).Value)
            'Frm_Inventário.TxtComprimento.Text = IIf(IsDBNull(DGV_Consulta.Item(43, e.RowIndex).Value), "", DGV_Consulta.Item(43, e.RowIndex).Value)
            'Frm_Inventário.TxtArea.Text = IIf(IsDBNull(DGV_Consulta.Item(44, e.RowIndex).Value), "", DGV_Consulta.Item(44, e.RowIndex).Value)
            'Frm_Inventário.TxtPe.Text = IIf(IsDBNull(DGV_Consulta.Item(45, e.RowIndex).Value), "", DGV_Consulta.Item(45, e.RowIndex).Value)
            'Frm_Inventário.TxtObsCivil.Text = IIf(IsDBNull(DGV_Consulta.Item(46, e.RowIndex).Value), "", DGV_Consulta.Item(46, e.RowIndex).Value)
            'Frm_Inventário.TxtConsultor.Text = DGV_Consulta.Item(47, e.RowIndex).Value
            'Frm_Inventário.TxtLider.Text = DGV_Consulta.Item(48, e.RowIndex).Value
            'Frm_Inventário.TxtTag.Text = DGV_Consulta.Item(50, e.RowIndex).Value
            'Frm_Inventário.TxtEsforco.Text = DGV_Consulta.Item(51, e.RowIndex).Value

            Texto_Foto = DGV_Consulta.Item(52, e.RowIndex).Value
        Catch
        End Try
        Frm_Inventário.A_Fotos_Inventario.Clear()
        Frm_Inventário.Foto = ""
        Frm_Inventário.PictureBox_Consulta.ImageLocation = ""

        If Texto_Foto <> "" Then
            Dim Palavras As String() = Texto_Foto.Split("|")
            For Each Palavra In Palavras
                Frm_Inventário.A_Fotos_Inventario.Add(Palavra)
            Next
            Frm_Inventário.PictureBox_Consulta.ImageLocation = Frm_Inventário.Caminho & "\" & Frm_Inventário.A_Fotos_Inventario(0)
        End If

        Frm_Inventário.BtnCopiar.Enabled = False
        Alterado = True
        Me.Close()
    End Sub

    Private Sub DGV_Consulta_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DGV_Consulta.CellMouseClick
        If e.Button = Windows.Forms.MouseButtons.Right AndAlso e.RowIndex >= 0 Then
            DGV_Consulta.MultiSelect = False
            DGV_Consulta.Rows(e.RowIndex).Selected = True
            Linha_ID = DGV_Consulta.Item(0, e.RowIndex).Value
            CMS_DGV.Show(DGV_Consulta, e.Location)
            CMS_DGV.Show(Cursor.Position)
            DGV_Consulta.MultiSelect = True
        End If

    End Sub

    Private Sub ExcluirDadosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExcluirDadosToolStripMenuItem.Click
        I_E.Excluir(Linha_ID)
        'Consultar Dados
        I_E.Consulta_Grid(DGV_Consulta)
        LblLinhas.Text = "Total de Registros: " & DGV_Consulta.Rows.Count
    End Sub

    Private Sub CopiarDadosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopiarDadosToolStripMenuItem.Click

    End Sub
End Class