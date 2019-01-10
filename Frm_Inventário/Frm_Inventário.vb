Imports System.IO
Public Class Frm_Inventário
    Dim I_E As New Inventário_Excel
    Dim C_I As New Class_Inventario
    Public TUC As Integer
    Public TI As Integer
    Public TI_Cod_Geral_Todos As Integer

    Dim ID As Integer
    Public Sequencial As String

    Public UAR As Integer
    Public A1 As String
    Public A2 As String
    Public A3 As String
    Public A4 As String
    Public A5 As String
    Public A6 As String

    Public CM1 As String
    Public CM2 As String
    Public CM3 As String

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

    Dim F_DGV As New Frm_DGV

    Private Sub Validacao_Salvar()
        If TxtLocal.Text = "" Or TxtODI.Text = "" Or CmbTI.Text = "" Or CmbTI_Geral.Text = "" Then
            MsgBox("Dados Incompletos na aba local", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If CmbTUC.Text = "" Then
            MsgBox("Preencha TUC", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If CmbA1.Text = "" Then
            MsgBox("Preencha Tipo de Bem", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If CmbUAR.Text = "" Then
            MsgBox("Preencha UAR", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If CmbA2.Text = "" And CmbA2.Enabled = True Then
            MsgBox("Preencha A2", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If CmbA3.Text = "" And CmbA3.Enabled = True Then
            MsgBox("Preencha A3", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If CmbA4.Text = "" And CmbA4.Enabled = True Then
            MsgBox("Preencha A4", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If CmbA5.Text = "" And CmbA5.Enabled = True Then
            MsgBox("Preencha A5", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If CmbA6.Text = "" And CmbA6.Enabled = True Then
            MsgBox("Preencha A6", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If CmbCm1.Text = "" Then
            MsgBox("Preencha cm1", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If CmbCm2.Text = "" Then
            MsgBox("Preencha cm2", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If CmbCm3.Text = "" Then
            MsgBox("Preencha cm3", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If TxtDesc.Text = "" Then
            MsgBox("Preencha descrição", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If TxtQtd.Text = "" Then
            MsgBox("Preencha quantidade", MsgBoxStyle.Exclamation)
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
        If TxtConsultor.Text = "" Then
            MsgBox("Preencha o e-mail do consultor", MsgBoxStyle.Exclamation)
            Invalidos = True
            Exit Sub
        End If
        If TxtLider.Text = "" Then
            MsgBox("Preencha o e-mail do líder", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
    End Sub

    Private Sub Limpar_Tudo()
        TxtBay.Text = ""
        CmbTUC.Text = ""
        A1 = ""
        A2 = ""
        A3 = ""
        A4 = ""
        A5 = ""
        A6 = ""
        CmbA1.Text = ""
        CmbA2.Text = ""
        CmbA3.Text = ""
        CmbA4.Text = ""
        CmbA5.Text = ""
        CmbA6.Text = ""
        CmbUAR.Text = ""
        CmbCm1.Text = ""
        CmbCm2.Text = ""
        CmbCm3.Text = ""
        TxtDesc.Text = ""
        TxtFabricante.Text = ""
        TxtModelo.Text = ""
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
        TxtTag.Text = ""
        CmbA2.Enabled = True
        CmbA3.Enabled = True
        CmbA4.Enabled = True
        CmbA5.Enabled = True
        CmbA6.Enabled = True
        CmbA2.Items.Clear()
        CmbA3.Items.Clear()
        CmbA4.Items.Clear()
        CmbA5.Items.Clear()
        CmbA6.Items.Clear()
        CmbUAR.Items.Clear()
        CmbA1.Items.Clear()
        LblA2.Text = "A2:"
        LblA3.Text = "A3:"
        LblA4.Text = "A4:"
        LblA5.Text = "A5:"
        LblA6.Text = "A6:"
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
        TxtTag.Text = ""
        A_Fotos_Inventario.Clear()
        PictureBox_Consulta.ImageLocation = ""
        Foto = ""
    End Sub

    Public Sub Ajuste_A1()
        A1 = I_E.Buscar_A1(CmbA1, TUC)
        I_E.Consulta_A2_A6(CmbA2, TUC, A1, "Desc_A2")
        I_E.Consulta_A2_A6(CmbA3, TUC, A1, "Desc_A3")
        I_E.Consulta_A2_A6(CmbA4, TUC, A1, "Desc_A4")
        I_E.Consulta_A2_A6(CmbA5, TUC, A1, "Desc_A5")
        I_E.Consulta_A2_A6(CmbA6, TUC, A1, "Desc_A6")
        I_E.Buscar_Tabela(LblA2, TUC, A1, "Desc_A2")
        I_E.Buscar_Tabela(LblA3, TUC, A1, "Desc_A3")
        I_E.Buscar_Tabela(LblA4, TUC, A1, "Desc_A4")
        I_E.Buscar_Tabela(LblA5, TUC, A1, "Desc_A5")
        I_E.Buscar_Tabela(LblA6, TUC, A1, "Desc_A6")
    End Sub

    Private Sub FrmInventario_Novo_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Panel_Picture_Consulta.Controls.Add(PictureBox_Consulta)
        'ID = I_E.Buscar_Ultimo_ID() + 1
        'TxtSeq_Civil.Text = ID
        'TxtSeq_Desc.Text = ID
        'TxtSeq_Local.Text = ID

        'I_E.Consulta_TUC(CmbTUC)
        'I_E.Consulta_CM(CmbCm1, CmbCm2, CmbCm3)
        'I_E.Consulta_TI_Geral(CmbTI_Geral)
        PB_Excel.Visible = False
        'Preencher ComboBoxes
        I_E.Preencher_CMB(CmbTI_Geral, CmbTI)

    End Sub

    Private Sub CmbTUC_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbTUC.SelectedIndexChanged
        TUC = I_E.Buscar_TUC(CmbTUC)
        I_E.Consulta_UAR(CmbUAR, TUC)
        I_E.Consulta_A1(CmbA1, TUC)
        'Limpar
        CmbUAR.Text = ""
        CmbA1.Text = ""
        CmbA2.Text = ""
        CmbA3.Text = ""
        CmbA4.Text = ""
        CmbA5.Text = ""
        CmbA6.Text = ""
        A2 = ""
        A3 = ""
        A4 = ""
        A5 = ""
        A6 = ""
        CmbA2.Items.Clear()
        CmbA3.Items.Clear()
        CmbA4.Items.Clear()
        CmbA5.Items.Clear()
        CmbA6.Items.Clear()
        CmbA2.Enabled = True
        CmbA3.Enabled = True
        CmbA4.Enabled = True
        CmbA5.Enabled = True
        CmbA6.Enabled = True
        LblA2.Text = "A2:"
        LblA3.Text = "A3:"
        LblA4.Text = "A4:"
        LblA5.Text = "A5:"
        LblA6.Text = "A6:"
        Colocar_Desc()
        If I_E.Contar_Unidade_TUC(TUC) = 1 Then
            CmbUm.Text = I_E.Unidade_TUC(TUC)
        End If
        'CADASTRO
        If I_E.Contar_Cadastro_TUC(TUC) = 1 Then
            LblCadastro.Text = I_E.Cadastro_TUC(TUC)
            Exit Sub
        End If

        If I_E.Contar_Cadastro_TUC_TI(TUC, TI) = 1 Then
            LblCadastro.Text = I_E.Cadastro_TUC_TI(TUC, TI)
        Else
            LblCadastro.Text = ""
        End If
    End Sub

    Private Sub Colocar_Desc()
        If CmbA2.Text = "" Then
            TxtDesc.Text = CmbTUC.Text & "; " & CmbA1.Text & " (" & CmbUAR.Text & ")"
            Exit Sub
        End If
        If CmbA3.Text = "" Then
            TxtDesc.Text = CmbTUC.Text & "; " & CmbA1.Text & "; " & CmbA2.Text & " (" & CmbUAR.Text & ")"
            Exit Sub
        End If
        If CmbA4.Text = "" Then
            TxtDesc.Text = CmbTUC.Text & "; " & CmbA1.Text & "; " & CmbA2.Text & "; " & CmbA3.Text & " (" & CmbUAR.Text & ")"
            Exit Sub
        End If
        If CmbA5.Text = "" Then
            TxtDesc.Text = CmbTUC.Text & "; " & CmbA1.Text & "; " & CmbA2.Text & "; " & CmbA3.Text & "; " & CmbA4.Text & " (" & CmbUAR.Text & ")"
            Exit Sub
        End If
        If CmbA6.Text = "" Then
            TxtDesc.Text = CmbTUC.Text & "; " & CmbA1.Text & "; " & CmbA2.Text & "; " & CmbA3.Text & "; " & CmbA4.Text & "; " &
                CmbA5.Text & " (" & CmbUAR.Text & ")"
        Else
            TxtDesc.Text = CmbTUC.Text & "; " & CmbA1.Text & "; " & CmbA2.Text & "; " & CmbA3.Text & "; " & CmbA4.Text & "; " &
                CmbA5.Text & "; " & CmbA6.Text & " (" & CmbUAR.Text & ")"
        End If
    End Sub
    Private Sub CmbCm1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCm1.SelectedIndexChanged
        If CmbCm1.Text = "NÃO APLICÁVEL" Then
            CM1 = "9"
        Else
            CM1 = I_E.Buscar_CM1(CmbCm1)
        End If
    End Sub

    Private Sub CmbCm2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCm2.SelectedIndexChanged
        If CmbCm2.Text = "NÃO APLICÁVEL" Then
            CM2 = "9"
        Else
            CM2 = I_E.Buscar_CM2(CmbCm2)
        End If
    End Sub

    Private Sub CmbCm3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCm3.SelectedIndexChanged
        If CmbCm3.Text = "NÃO APLICÁVEL" Then
            CM3 = "9"
        Else
            CM3 = I_E.Buscar_CM3(CmbCm3)
        End If
    End Sub

    Private Sub CmbTI_Geral_Click(sender As Object, e As EventArgs) Handles CmbTI_Geral.Click
        CmbTI.Text = ""
    End Sub

    Private Sub CmbUAR_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbUAR.SelectedIndexChanged
        UAR = I_E.Buscar_UAR(CmbUAR, TUC)
        Colocar_Desc()
    End Sub

    Private Sub CmbA1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbA1.SelectedIndexChanged
        Me.Cursor = Cursors.WaitCursor
        Ajuste_A1()
        'Limpar
        CmbA2.Text = ""
        CmbA3.Text = ""
        CmbA4.Text = ""
        CmbA5.Text = ""
        CmbA6.Text = ""
        A2 = ""
        A3 = ""
        A4 = ""
        A5 = ""
        A6 = ""
        Colocar_Desc()
        If I_E.Contar_Unidade_A1(TUC, A1) = 1 Then
            CmbUm.Text = I_E.Unidade_A1(TUC, A1)
        End If

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub CmbA2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbA2.SelectedIndexChanged
        A2 = I_E.Buscar_A2_A6(CmbA2)
        Colocar_Desc()
        'CADASTRO
        If I_E.Contar_Cadastro_TUC_A1_A2(TUC, A1, A2) = 1 Then
            LblCadastro.Text = I_E.Cadastro_TUC_A1_A2(TUC, A1, A2)
            Exit Sub
        End If

        If I_E.Contar_Cadastro_TUC_A2(TUC, A2) = 1 Then
            LblCadastro.Text = I_E.Cadastro_TUC_A2(TUC, A2)
            Exit Sub
        End If

        If I_E.Contar_Cadastro_TUC_TI_A2(TUC, TI, A2) = 1 Then
            LblCadastro.Text = I_E.Cadastro_TUC_TI_A2(TUC, TI, A2)
            Exit Sub
        End If
        If I_E.Contar_Cadastro_TUC_TI(TUC, TI) = 1 Then
            LblCadastro.Text = I_E.Cadastro_TUC_TI(TUC, TI)
        End If
    End Sub

    Private Sub CmbA3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbA3.SelectedIndexChanged
        A3 = I_E.Buscar_A2_A6(CmbA3)
        Colocar_Desc()
        If I_E.Contar_Unidade_A3(TUC, A1, A3) = 1 Then
            CmbUm.Text = I_E.Unidade_A3(TUC, A1, A3)
        End If
        'CADASTRO
        If I_E.Contar_Cadastro_TUC_A3(TUC, A3) = 1 Then
            LblCadastro.Text = I_E.Cadastro_TUC_A3(TUC, A3)
            Exit Sub
        End If
        If I_E.Contar_Cadastro_TUC_TI_A3(TUC, TI, A3) = 1 Then
            LblCadastro.Text = I_E.Cadastro_TUC_TI_A3(TUC, TI, A3)
            Exit Sub
        End If
        If I_E.Contar_Cadastro_TUC_TI(TUC, TI) = 1 Then
            LblCadastro.Text = I_E.Cadastro_TUC_TI(TUC, TI)
        End If
    End Sub

    Private Sub CmbA4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbA4.SelectedIndexChanged
        A4 = I_E.Buscar_A2_A6(CmbA4)
        Colocar_Desc()
        If I_E.Contar_Unidade_A4(TUC, A4) = 1 Then
            CmbUm.Text = I_E.Unidade_A4(TUC, A4)
        End If
    End Sub

    Private Sub CmbA5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbA5.SelectedIndexChanged
        A5 = I_E.Buscar_A2_A6(CmbA5)
        Colocar_Desc()
    End Sub

    Private Sub CmbA6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbA6.SelectedIndexChanged
        A6 = I_E.Buscar_A2_A6(CmbA6)
        Colocar_Desc()
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

    'Private Sub BtnZoom_Click(sender As Object, e As EventArgs) Handles BtnZoom.Click
    '    Dim p As New Process()
    '    p.StartInfo.FileName = "rundll32.exe"
    '    p.StartInfo.Arguments = Path.Combine(Environment.SystemDirectory, "shimgvw.dll" & ",ImageView_Fullscreen " & PictureBox.ImageLocation)
    '    p.Start()
    'End Sub

    Private Sub BtnSalvar_Click(sender As Object, e As EventArgs) Handles BtnSalvar.Click
        'Validação de Dados
        Validacao_Salvar()
        If Invalidos = True Then
            Invalidos = False
            Exit Sub
        End If

        BtnS_Multi.Enabled = True
        BtnCopiar.Enabled = True

        'Update BD
        Sequencial = TxtLocal.Text & " - " & TxtSeq_Local.Text
        consultor = TxtConsultor.Text
        lider = TxtLider.Text

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
            I_E.Update_Inventario(ID, Sequencial, TxtLocal.Text, TxtODI.Text, TI, CmbTI.Text, TxtBay.Text, TUC, CmbTUC.Text, A1, CmbA1.Text,
                              UAR, CmbUAR.Text, A2, CmbA2.Text, A3, CmbA3.Text, A4, CmbA4.Text, A5, CmbA5.Text, A6, CmbA6.Text, CM1, CmbCm1.Text,
                              CM2, CmbCm2.Text, CM3, CmbCm3.Text, TxtDesc.Text, TxtFabricante.Text, TxtModelo.Text, TxtSerie.Text, TxtObs.Text,
                              TxtQtd.Text, CmbUm.Text, CmbAno.Text, CmbMes.Text, CmbDia.Text, CmbStatus.Text, CmbEstado.Text, TxtAltura.Text,
                              TxtLargura.Text, TxtComprimento.Text, TxtArea.Text, TxtPe.Text, TxtObsCivil.Text, Foto, consultor, lider, TxtTag.Text, TxtEsforco.Text)
        Else
            I_E.Inserir_Dados(ID, Sequencial, TxtLocal.Text, TxtODI.Text, TI, CmbTI.Text, TxtBay.Text, TUC, CmbTUC.Text, A1, CmbA1.Text,
                              UAR, CmbUAR.Text, A2, CmbA2.Text, A3, CmbA3.Text, A4, CmbA4.Text, A5, CmbA5.Text, A6, CmbA6.Text, CM1, CmbCm1.Text,
                              CM2, CmbCm2.Text, CM3, CmbCm3.Text, TxtDesc.Text, TxtFabricante.Text, TxtModelo.Text, TxtSerie.Text, TxtObs.Text,
                              TxtQtd.Text, CmbUm.Text, CmbAno.Text, CmbMes.Text, CmbDia.Text, CmbStatus.Text, CmbEstado.Text, TxtAltura.Text,
                              TxtLargura.Text, TxtComprimento.Text, TxtArea.Text, TxtPe.Text, TxtEsforco.Text, TxtObsCivil.Text, consultor, lider, TxtTag.Text, Foto)
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

    Private Sub BtnS_Multi_Click(sender As Object, e As EventArgs) Handles BtnS_Multi.Click
        'Validação de Dados
        Validacao_Salvar()
        If Invalidos = True Then
            Invalidos = False
            Exit Sub
        End If

        Dim N_cadastros As String
        N_cadastros = InputBox("Número de Cadastros. Limite 10!", "Cadastros")
        If N_cadastros = "" Then
            Exit Sub
        End If

        If IsNumeric(N_cadastros) Then
            If N_cadastros > 10 Then
                MsgBox("Insira valores menores ou iguais a 10!")
                Exit Sub
            End If
            If N_cadastros <= 0 Then
                MsgBox("Insira valores maiores que 0!")
                Exit Sub
            End If
        Else
            MsgBox("Insira dados numéricos!")
            Exit Sub
        End If

        'Update BD
        consultor = TxtConsultor.Text
        lider = TxtLider.Text

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
        For i = ID To (ID + N_cadastros - 1)
            Sequencial = TxtLocal.Text & " - " & i
            I_E.Inserir_Dados(i, Sequencial, TxtLocal.Text, TxtODI.Text, TI, CmbTI.Text, TxtBay.Text, TUC, CmbTUC.Text, A1, CmbA1.Text,
                                  UAR, CmbUAR.Text, A2, CmbA2.Text, A3, CmbA3.Text, A4, CmbA4.Text, A5, CmbA5.Text, A6, CmbA6.Text, CM1, CmbCm1.Text,
                                  CM2, CmbCm2.Text, CM3, CmbCm3.Text, TxtDesc.Text, TxtFabricante.Text, TxtModelo.Text, TxtSerie.Text, TxtObs.Text,
                                  TxtQtd.Text, CmbUm.Text, CmbAno.Text, CmbMes.Text, CmbDia.Text, CmbStatus.Text, CmbEstado.Text, TxtAltura.Text,
                                  TxtLargura.Text, TxtComprimento.Text, TxtArea.Text, TxtPe.Text, TxtEsforco.Text, TxtObsCivil.Text, consultor, lider, TxtTag.Text, Foto)

        Next i
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
        I_E.Consulta_Descricao_Civil(ID, TxtBay, TUC, CmbTUC, A1, CmbA1, UAR, CmbUAR, CmbA2, CmbA3, CmbA4,
                                     CmbA5, CmbA6, CM1, CmbCm1, CM2, CmbCm2, CM3, CmbCm3, TxtDesc, TxtFabricante, TxtModelo, TxtObs,
                                     TxtQtd, CmbUm, CmbAno, CmbMes, CmbDia, CmbStatus, CmbEstado, TxtAltura, TxtLargura, TxtComprimento,
                                     TxtArea, TxtPe, TxtObsCivil, TxtEsforco, TxtSerie, TxtTag)
        'Consultar A2 a A6
        Ajuste_A1()
        I_E.Consulta_UAR(CmbUAR, TUC)
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

    Private Sub ExcluirDadosAnterioresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExcluirDadosAnterioresToolStripMenuItem.Click
        Dim Result As DialogResult = MessageBox.Show("Deseja excluir os dados anteriores?", "Dados", MessageBoxButtons.YesNo)
        If Result = vbYes Then
            I_E.Excluir_Tudo()
            ID = 1
            TxtSeq_Civil.Text = ID
            TxtSeq_Desc.Text = ID
            TxtSeq_Local.Text = ID
        End If
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

    Private Sub CarregarBaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CarregarBaseToolStripMenuItem.Click
        On Error Resume Next
        Dim dr As DialogResult = Me.OFD.ShowDialog()
        Dim Caminho As String
        If dr = System.Windows.Forms.DialogResult.OK Then
            Caminho = OFD.FileName
        Else
            Exit Sub
        End If
        'Carga
        I_E.Carga_Sistema(Caminho)
    End Sub
End Class