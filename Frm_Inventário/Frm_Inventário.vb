Imports System.IO
Public Class Frm_Inventário
    Dim I_E As New Inventário_Excel
    Dim C_I As New Class_Inventario
    Public TUC As Integer
    Public TI As Integer
    Public TI_Cod_Geral_Todos As Integer

    Dim ID As Integer
    Public Sequencial As String

    Public A_Instal As New ArrayList
    Public Cod_Instal As String


    Public consultor As String
    Public lider As String
    Public Foto As String

    Public A_Fotos_Principal As New ArrayList
    Public A_Fotos_Inventario As New ArrayList
    Dim N_Foto_Principal As Integer
    Dim N_Foto_Inventario As Integer

    Dim V_Atual_TB As Integer = 0

    Public Caminho As String

    Dim Invalidos As Boolean

    Public Erro_Excluir As Boolean

    Dim F_DGV As New Frm_DGV

    Private Sub Validacao_Salvar()
        If CmbCC.Text = "" Or CmbInstall.Text = "" Or CmbLocal.Text = "" Then
            MsgBox("Dados Incompletos na aba local", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If CmbDesc.Text = "" Then
            MsgBox("Preencha Descrição Simples", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If TxtQtd.Text = "" Then
            MsgBox("Preencha a Quantidade", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If

        If CmbUm.Text = "" Then
            MsgBox("Preencha unidade de medida", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If CmbStatus.Text = "" Then
            MsgBox("Preencha status do bem", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If CmbEstado.Text = "" Then
            MsgBox("Preencha estado do bem", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If CmbConsultor.Text = "" Then
            MsgBox("Preencha o Consultor", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If CmbResponsavel.Text = "" Then
            MsgBox("Preencha o Responsável", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
    End Sub

    Private Sub Limpar_Tudo()
        TxtTagNovo.Text = ""
        CmbTagAntigo.Text = ""
        CmbDesc.Text = ""
        TxtDetalhada.Text = ""
        CmbLocalFisico.Text = ""
        CmbFabricante.Text = ""
        CmbModelo.Text = ""
        TxtSerie.Text = ""
        TxtObs.Text = ""
        TxtQtd.Text = 1
        CmbUm.Text = "UN"
        CmbAno.Text = ""
        CmbMes.Text = ""
        CmbDia.Text = ""
        CmbStatus.Text = ""
        CmbEstado.Text = ""
        TxtAltura.Text = ""
        TxtLargura.Text = ""
        TxtComprimento.Text = ""
        TxtArea.Text = ""
        TxtEsforco.Text = ""
        TxtPe.Text = ""
        TxtObsCivil.Text = ""
        A_Fotos_Inventario.Clear()
        Foto = ""
        PictureBox_Consulta.ImageLocation = ""
    End Sub

    Private Sub Limpar_Parcial()
        TxtSerie.Text = ""
        CmbStatus.Text = ""
        CmbEstado.Text = ""
        CmbAno.Text = ""
        CmbMes.Text = ""
        CmbDia.Text = ""
        TxtObs.Text = ""
        TxtObsCivil.Text = ""
        CmbTagAntigo.Text = ""
        A_Fotos_Inventario.Clear()
        PictureBox_Consulta.ImageLocation = ""
        Foto = ""
    End Sub

    Private Sub FrmInventario_Novo_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        I_E.Buscar_Data_Limite()

        If I_E.DTExpira.Rows(0)(0) <= Today Then
            MsgBox("Data Expirada para Uso do Software. Contate o administrador.", vbCritical)
            Application.Exit()
        End If

        Panel_Picture_Consulta.Controls.Add(PictureBox_Consulta)
        ID = I_E.Buscar_Ultimo_ID() + 1
        TxtSeq_Civil.Text = ID
        TxtSeq_Desc.Text = ID
        TxtSeq_Local.Text = ID

        PB_Excel.Visible = False
        'Preencher ComboBoxes
        I_E.Preencher_CMB(CmbCC, CmbInstall, CmbLocal, CmbTagAntigo, CmbDesc, CmbFabricante, CmbModelo,
                          CmbLocalFisico, CmbConsultor, CmbResponsavel, A_Instal)

    End Sub

    Private Sub BtnAnterior_Click(sender As Object, e As EventArgs) Handles BtnAnterior.Click
        N_Foto_Principal = C_I.Anterior_Foto(A_Fotos_Principal, PictureBox, N_Foto_Principal, Caminho)
    End Sub

    Private Sub BtnProximo_Click(sender As Object, e As EventArgs) Handles BtnProximo.Click
        N_Foto_Principal = C_I.Proxima_Foto(A_Fotos_Principal, PictureBox, N_Foto_Principal, Caminho)
    End Sub

    Private Sub Btn_Voltar10_Click(sender As Object, e As EventArgs) Handles Btn_Voltar10.Click
        N_Foto_Principal = C_I.Anterior_Foto(A_Fotos_Principal, PictureBox, N_Foto_Principal - 9, Caminho)
    End Sub

    Private Sub Btn_Avancar10_Click(sender As Object, e As EventArgs) Handles Btn_Avancar10.Click
        N_Foto_Principal = C_I.Proxima_Foto(A_Fotos_Principal, PictureBox, N_Foto_Principal + 9, Caminho)
    End Sub

    Private Sub BtnSalvar_Click(sender As Object, e As EventArgs) Handles BtnSalvar.Click
        'Validação de Dados
        Dim Seq_At As String = ""

        Validacao_Salvar()
        If Invalidos = True Then
            Invalidos = False
            Exit Sub
        End If

        BtnCopiar.Enabled = True

        If TxtSeq_Local.Text.Count = 1 Then
            Seq_At = "AT00000"
        ElseIf TxtSeq_Local.Text.Count = 2 Then
            Seq_At = "AT0000"
        ElseIf TxtSeq_Local.Text.Count = 3 Then
            Seq_At = "AT000"
        ElseIf TxtSeq_Local.Text.Count = 4 Then
            Seq_At = "AT00"
        ElseIf TxtSeq_Local.Text.Count = 5 Then
            Seq_At = "AT0"
        ElseIf TxtSeq_Local.Text.Count >= 6 Then
            Seq_At = "AT"
        End If
        'Update BD
        Sequencial = Seq_At & TxtSeq_Local.Text
        consultor = CmbConsultor.Text
        lider = CmbResponsavel.Text

        ID = TxtSeq_Civil.Text
        If CmbAno.Text = "" Then
            CmbAno.Text = 0
        End If
        If CmbMes.Text = "" Then
            CmbMes.Text = 0
        End If
        If CmbDia.Text = "" Then
            CmbDia.Text = 0
        End If
        If TxtAltura.Text = "" Then
            TxtAltura.Text = 0
        End If
        If TxtLargura.Text = "" Then
            TxtLargura.Text = 0
        End If
        If TxtComprimento.Text = "" Then
            TxtComprimento.Text = 0
        End If
        If TxtArea.Text = "" Then
            TxtArea.Text = 0
        End If
        If TxtPe.Text = "" Then
            TxtPe.Text = 0
        End If
        If TxtEsforco.Text = "" Then
            TxtEsforco.Text = 0
        End If

        'Arrumar foto
        If A_Fotos_Inventario.Count > 0 Then
            For i = 0 To A_Fotos_Inventario.Count - 1
                Foto = Foto & IIf(Foto = "", "", "|") & A_Fotos_Inventario(i)
            Next
        Else
            Foto = ""
        End If

        'Se o ID do cadastro já existir, faça o Update, senão Insert
        If ID <= I_E.Buscar_Ultimo_ID Then
            'Consulta e compara os dados para saber se é Validado ou Alterado
            Dim Validacao As String
            If I_E.Consulta_Validar(ID, CmbCC.Text, Cod_Instal, CmbInstall.Text, CmbLocal.Text, CmbTagAntigo.Text, TxtTagNovo.Text,
CmbDesc.Text, TxtDetalhada.Text, CmbFabricante.Text, CmbModelo.Text, TxtSerie.Text, CmbLocalFisico.Text, TxtQtd.Text, CmbUm.Text,
CmbAno.Text, CmbMes.Text, CmbDia.Text, CmbStatus.Text, CmbEstado.Text, TxtObs.Text, TxtAltura.Text, TxtLargura.Text, TxtComprimento.Text,
TxtArea.Text, TxtPe.Text, TxtEsforco.Text, TxtObsCivil.Text, CmbConsultor.Text, CmbResponsavel.Text) = True Then
                Validacao = "ALTERADO"
            Else
                Validacao = "VALIDADO"
            End If

            I_E.Update_Inventario(ID, CmbLocal.Text, CmbCC.Text, Cod_Instal, CmbInstall.Text, CmbTagAntigo.Text, TxtTagNovo.Text, CmbDesc.Text, TxtDetalhada.Text, CmbFabricante.Text, CmbModelo.Text, TxtSerie.Text, CmbLocalFisico.Text,
                              TxtObs.Text, Validacao, TxtQtd.Text, CmbUm.Text, CmbAno.Text, CmbMes.Text, CmbDia.Text, CmbStatus.Text, CmbEstado.Text, TxtAltura.Text,
                              TxtLargura.Text, TxtComprimento.Text, TxtArea.Text, TxtPe.Text, TxtEsforco.Text, TxtObsCivil.Text, Foto, consultor, lider)
        Else
            I_E.Inserir_Dados(ID, Sequencial, CmbLocal.Text, CmbCC.Text, Cod_Instal, CmbInstall.Text, CmbTagAntigo.Text,
                              TxtTagNovo.Text, CmbDesc.Text, TxtDetalhada.Text, CmbFabricante.Text, CmbModelo.Text,
                              TxtSerie.Text, CmbLocalFisico.Text, TxtObs.Text, "NOVO", TxtQtd.Text, CmbUm.Text, CmbAno.Text,
                              CmbMes.Text, CmbDia.Text, CmbStatus.Text, CmbEstado.Text, TxtAltura.Text, TxtLargura.Text,
                              TxtComprimento.Text, TxtArea.Text, TxtPe.Text, TxtEsforco.Text, TxtObsCivil.Text, Foto, consultor, lider)
        End If
        BtnCopiar.Enabled = True
        'Limpar Dados
        Limpar_Tudo()
        ID = I_E.Buscar_Ultimo_ID
        TxtSeq_Civil.Text = ID + 1
        TxtSeq_Desc.Text = ID + 1
        TxtSeq_Local.Text = ID + 1

        CmbStatus.Text = "EM USO"
        CmbEstado.Text = "BOM"
    End Sub

    Private Sub BtnCopiar_Click(sender As Object, e As EventArgs) Handles BtnCopiar.Click
        'Limpar
        Limpar_Tudo()
        'Consultar ID +1
        ID = I_E.Buscar_Ultimo_ID()
        I_E.Consulta_Descricao_Civil(ID, CmbTagAntigo, TxtTagNovo, CmbDesc, TxtDetalhada, CmbFabricante, CmbModelo, TxtObs,
                                     TxtQtd, CmbUm, CmbStatus, CmbEstado, TxtAltura, TxtLargura, TxtComprimento,
                                     TxtArea, TxtPe, TxtEsforco, TxtObsCivil, CmbLocalFisico)
        ID += 1
        TxtSeq_Civil.Text = ID
        TxtSeq_Desc.Text = ID
        TxtSeq_Local.Text = ID
    End Sub

    Private Sub TxtQtd_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtQtd.KeyPress
        If e.KeyChar = ","c Then
            e.Handled = (CType(sender, TextBox).Text.IndexOf(","c) <> -1)
        ElseIf e.KeyChar <> ControlChars.Back Then
            e.Handled = ("0123456789".IndexOf(e.KeyChar) = -1)
        End If
    End Sub

    Private Sub TxtAltura_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtAltura.KeyPress
        If e.KeyChar = ","c Then
            e.Handled = (CType(sender, TextBox).Text.IndexOf(","c) <> -1)
        ElseIf e.KeyChar <> ControlChars.Back Then
            e.Handled = ("0123456789".IndexOf(e.KeyChar) = -1)
        End If
    End Sub

    Private Sub TxtLargura_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtLargura.KeyPress
        If e.KeyChar = ","c Then
            e.Handled = (CType(sender, TextBox).Text.IndexOf(","c) <> -1)
        ElseIf e.KeyChar <> ControlChars.Back Then
            e.Handled = ("0123456789".IndexOf(e.KeyChar) = -1)
        End If
    End Sub

    Private Sub TxtEsforco_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtEsforco.KeyPress
        If e.KeyChar = ","c Then
            e.Handled = (CType(sender, TextBox).Text.IndexOf(","c) <> -1)
        ElseIf e.KeyChar <> ControlChars.Back Then
            e.Handled = ("0123456789".IndexOf(e.KeyChar) = -1)
        End If
    End Sub

    Private Sub TxtComprimento_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtComprimento.KeyPress
        If e.KeyChar = ","c Then
            e.Handled = (CType(sender, TextBox).Text.IndexOf(","c) <> -1)
        ElseIf e.KeyChar <> ControlChars.Back Then
            e.Handled = ("0123456789".IndexOf(e.KeyChar) = -1)
        End If
    End Sub

    Private Sub TxtArea_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtArea.KeyPress
        If e.KeyChar = ","c Then
            e.Handled = (CType(sender, TextBox).Text.IndexOf(","c) <> -1)
        ElseIf e.KeyChar <> ControlChars.Back Then
            e.Handled = ("0123456789".IndexOf(e.KeyChar) = -1)
        End If
    End Sub

    Private Sub TxtPe_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtPe.KeyPress
        If e.KeyChar = ","c Then
            e.Handled = (CType(sender, TextBox).Text.IndexOf(","c) <> -1)
        ElseIf e.KeyChar <> ControlChars.Back Then
            e.Handled = ("0123456789".IndexOf(e.KeyChar) = -1)
        End If
    End Sub

    Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles BtnAdd.Click
        Try
            A_Fotos_Inventario.Add(A_Fotos_Principal(N_Foto_Principal))
            PictureBox_Consulta.ImageLocation = Caminho & "\" & A_Fotos_Inventario(A_Fotos_Inventario.Count - 1)
            N_Foto_Inventario = A_Fotos_Inventario.Count - 1
        Catch
        End Try
    End Sub
    Private Sub BtnRemover_Fotos_Click(sender As Object, e As EventArgs) Handles BtnRemover_Fotos.Click
        Try
            A_Fotos_Inventario.RemoveAt(N_Foto_Inventario)
            If A_Fotos_Inventario.Count >= 1 Then
                N_Foto_Inventario -= 1
                If N_Foto_Inventario < 0 Then
                    N_Foto_Inventario = 0
                End If
                PictureBox_Consulta.ImageLocation = Caminho & "\" & A_Fotos_Inventario(N_Foto_Inventario)
            Else
                PictureBox_Consulta.ImageLocation = ""
                N_Foto_Inventario = 0
                Exit Sub
            End If

        Catch
        End Try
    End Sub

    Private Sub BtnAnterior_Consulta_Click(sender As Object, e As EventArgs) Handles BtnAnterior_Consulta.Click
        N_Foto_Inventario = C_I.Anterior_Foto(A_Fotos_Inventario, PictureBox_Consulta, N_Foto_Inventario, Caminho)
    End Sub

    Private Sub BtnProximo_Consulta_Click(sender As Object, e As EventArgs) Handles BtnProximo_Consulta.Click
        N_Foto_Inventario = C_I.Proxima_Foto(A_Fotos_Inventario, PictureBox_Consulta, N_Foto_Inventario, Caminho)
    End Sub

    Private Sub BtnConsultar_Click(sender As Object, e As EventArgs) Handles BtnConsultar.Click
        Frm_DGV.Show()
        Me.Hide()
    End Sub

    Private Sub CaminhoFotosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CaminhoFotosToolStripMenuItem.Click
        FBD.ShowDialog()
        Caminho = FBD.SelectedPath
        If Caminho = "" Then
            Exit Sub
        End If
        'Dim F_Arquivos = Directory.GetFiles(Caminho)
        Dim di As New DirectoryInfo(Caminho)
        Dim F_Arquivos = di.GetFiles()
        F_Arquivos = F_Arquivos.OrderBy(Function(x) x.CreationTime).ToArray()

        For Each A_F As FileInfo In F_Arquivos
            A_Fotos_Principal.Add(Path.GetFileName(A_F.ToString))
        Next A_F

        'Mostrar Imagem no PictureBox
        PictureBox.ImageLocation = Caminho & "\" & A_Fotos_Principal(0)
    End Sub

    Private Sub ExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExcelToolStripMenuItem.Click
        I_E.Modelo_Excel()
    End Sub

    Private Sub BtnGirar_Click(sender As Object, e As EventArgs) Handles BtnGirar.Click
        Try
            PictureBox_Consulta.Image.RotateFlip(RotateFlipType.Rotate90FlipNone)
            PictureBox_Consulta.Refresh()
        Catch
        End Try
    End Sub

    Private Sub TB_ValueChanged(sender As Object, e As EventArgs) Handles TB.ValueChanged
        If V_Atual_TB < TB.Value Then
            PictureBox_Consulta.Width += TB.Value * (20%)
            PictureBox_Consulta.Height += TB.Value * (20%)
        Else
            PictureBox_Consulta.Width -= (TB.Value + 1) * (20%)
            PictureBox_Consulta.Height -= (TB.Value + 1) * (20%)
        End If
        If TB.Value = 0 Then
            PictureBox_Consulta.Width = 530
            PictureBox_Consulta.Height = 450
        End If
        V_Atual_TB = TB.Value
    End Sub


    Private Sub BaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BaseToolStripMenuItem.Click
        'Carregar Inventário
        On Error Resume Next
        Dim dr As DialogResult = Me.OFD.ShowDialog()
        Dim CaminhoI As String
        If dr = System.Windows.Forms.DialogResult.OK Then
            CaminhoI = OFD.FileName
        Else
            Exit Sub
        End If
        'Carga
        I_E.Carga_Inventario(CaminhoI)
    End Sub

    Private Sub LayoutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LayoutToolStripMenuItem.Click
        'Carregar ComboBox
        On Error Resume Next
        Dim dr As DialogResult = Me.OFD.ShowDialog()
        Dim CaminhoC As String
        If dr = System.Windows.Forms.DialogResult.OK Then
            CaminhoC = OFD.FileName
        Else
            Exit Sub
        End If
        I_E.Carga_Cmb(CaminhoC)
    End Sub

    Private Sub CargaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CargaToolStripMenuItem.Click
        I_E.Layout_Excel_Cmb()
    End Sub

    Private Sub InventárioToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InventárioToolStripMenuItem.Click
        I_E.Layout_Excel()
    End Sub

    Private Sub CmbInstall_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbInstall.SelectedIndexChanged
        Cod_Instal = A_Instal(CmbInstall.SelectedIndex)
    End Sub

    Private Sub ConsultarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsultarToolStripMenuItem.Click
        'Colocar InputBox para escolher TAG, caso tenha faça a consulta SqLite na tela, senão dá uma mensagem falando que não foi localizado
        Dim Consulta_TAG As String
        Consulta_TAG = InputBox("Escreva o TAG a ser localizado:", "Encontrar")
        If Consulta_TAG <> "" Then

            I_E.Consulta_Tag(ID, CmbCC, CmbLocal, Cod_Instal, CmbInstall, CmbTagAntigo, TxtTagNovo, CmbDesc, TxtDetalhada, CmbConsultor,
CmbResponsavel, CmbFabricante, CmbModelo, TxtSerie, TxtObs, TxtQtd, CmbUm, CmbStatus, CmbAno, CmbMes, CmbDia, CmbEstado, TxtAltura,
TxtLargura, TxtComprimento, TxtArea, TxtPe, TxtEsforco, TxtObsCivil, CmbLocalFisico, Consulta_TAG, TxtSeq_Local, TxtSeq_Desc, TxtSeq_Civil)
        End If
    End Sub

    Private Sub InventárioToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles InventárioToolStripMenuItem1.Click
        Dim Result As DialogResult = MessageBox.Show("Deseja excluir os dados anteriores?", "Dados", MessageBoxButtons.YesNo)
        If Result = vbYes Then
            I_E.Excluir_Tudo()
            If Erro_Excluir = False Then
                MsgBox("Dados Excluídos com Sucesso", vbInformation)
            End If

            ID = 1
            TxtSeq_Civil.Text = ID
            TxtSeq_Desc.Text = ID
            TxtSeq_Local.Text = ID
        End If
    End Sub

    Private Sub CaixaDeSeleçãoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CaixaDeSeleçãoToolStripMenuItem.Click
        Dim Result As DialogResult = MessageBox.Show("Deseja excluir os dados anteriores?", "Dados", MessageBoxButtons.YesNo)
        If Result = vbYes Then
            I_E.Excluir_Carga_Cmb()
        End If
    End Sub

    Private Sub LicençaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LicençaToolStripMenuItem.Click

        If InputBox("Escreva a senha de acesso", "Senha") = "@t0M05I" Then
            FrmData.Show()
        Else
            MsgBox("Senha Incorreta", vbCritical)
        End If

    End Sub
End Class