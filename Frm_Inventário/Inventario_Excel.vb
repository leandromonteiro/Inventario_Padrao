Imports System.Net.Mail
Imports Microsoft.Office.Interop

Public Class Inventário_Excel
    Dim connstr As String = "Data Source=C:\Users\Public\INVENTARIO_PADRAO.db;;Version=3;New=True;Compress=True;Pooling=True"
    Public DS As New DataSet

    Public Sub Preencher_CMB(CmbCC As ComboBox, CmbInstall As ComboBox, CmbLocal As ComboBox, CmbTag As ComboBox,
                             CmbDesc As ComboBox, CmbFabricante As ComboBox, CmbModelo As ComboBox,
                             CmbLocalFisico As ComboBox, CmbConsultor As ComboBox, CmbResponsavel As ComboBox, A_Install As ArrayList)
        Dim DT As New DataTable
        Try
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim DA As New SQLite.SQLiteDataAdapter
            connection.Open()
            'DT.Load(Command.ExecuteReader(CommandBehavior.CloseConnection))
            DA.SelectCommand = New SQLite.SQLiteCommand("select * from Carga_Cmb;", connection)
            DA.Fill(DT)
            connection.Close()
            connection.Dispose()
            DA.Dispose()
            GC.Collect()
        Catch
            MsgBox("Erro na consulta da Carga", MsgBoxStyle.Critical)
        End Try

        'Inserir nos CMB
        Try
            For i = 0 To DT.Rows.Count
                If Not DT.Rows(i)(0) = "" Then
                    CmbCC.Items.Add(DT.Rows(i)(0))
                End If
                If Not DT.Rows(i)(1) = "" Then
                    CmbInstall.Items.Add(DT.Rows(i)(2))
                    A_Install.Add(DT.Rows(i)(1))
                End If
                If Not DT.Rows(i)(3) = "" Then
                    CmbLocal.Items.Add(DT.Rows(i)(3))
                End If
                If Not DT.Rows(i)(4) = "" Then
                    CmbTag.Items.Add(DT.Rows(i)(4))
                End If
                If Not DT.Rows(i)(5) = "" Then
                    CmbDesc.Items.Add(DT.Rows(i)(5))
                End If
                If Not DT.Rows(i)(6) = "" Then
                    CmbFabricante.Items.Add(DT.Rows(i)(6))
                End If
                If Not DT.Rows(i)(7) = "" Then
                    CmbModelo.Items.Add(DT.Rows(i)(7))
                End If
                If Not DT.Rows(i)(8) = "" Then
                    CmbLocalFisico.Items.Add(DT.Rows(i)(8))
                End If
                If Not DT.Rows(i)(9) = "" Then
                    CmbConsultor.Items.Add(DT.Rows(i)(9))
                End If
                If Not DT.Rows(i)(10) = "" Then
                    CmbResponsavel.Items.Add(DT.Rows(i)(10))
                End If
            Next i
        Catch
        End Try
    End Sub

    Public Sub Carga_Cmb(Caminho As String)
        'Coloca os dados do Excel no DT
        Dim DT As New DataTable
        Try
            Dim connection As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=YES;IMEX=1';DATA SOURCE=" & Caminho & ";")
            Dim DA As New OleDb.OleDbDataAdapter
            connection.Open()
            'DT.Load(Command.ExecuteReader(CommandBehavior.CloseConnection))
            DA.SelectCommand = New OleDb.OleDbCommand("select * from [Carga$];", connection)
            DA.Fill(DT)
            connection.Close()
            connection.Dispose()
            DA.Dispose()
            GC.Collect()
        Catch
            MsgBox("Erro na consulta da Carga", MsgBoxStyle.Critical)
        End Try

        'Limpa a base e insere com os dados DT
        Try
            Dim connectionS As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connectionS.Open()
            cmd.Connection = connectionS
            cmd.CommandText = "delete from Carga_Cmb;"
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            connectionS.Close()
            connectionS.Dispose()
        Catch
            MsgBox("Erro ao limpar Base", MsgBoxStyle.Critical)
        End Try

        Try
            Dim connectionS As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connectionS.Open()
            cmd.Connection = connectionS
            For i = 0 To DT.Rows.Count - 1
                cmd.CommandText = "insert into Carga_Cmb (Centro_Custo,Cod_Instalacao,Desc_Instalacao,Local," &
                    "Tag,Descricao,Fabricante,Modelo,Local_Fisico,Consultor,Responsável) values ('" &
                    DT.Rows(i)(0) & "','" & DT.Rows(i)(1) & "','" & DT.Rows(i)(2) & "','" & DT.Rows(i)(3) & "','" &
                    DT.Rows(i)(4) & "','" & DT.Rows(i)(5) & "','" & DT.Rows(i)(6) & "','" & DT.Rows(i)(7) & "','" &
                    DT.Rows(i)(8) & "','" & DT.Rows(i)(9) & "','" & DT.Rows(i)(10) & "');"
                cmd.ExecuteNonQuery()
            Next i
            cmd.Dispose()
            connectionS.Close()
            connectionS.Dispose()
            MsgBox("Base carregada com Sucesso", MsgBoxStyle.Information)
        Catch
            MsgBox("Erro ao inserir Base", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Carga_Inventario(Caminho As String)
        'Coloca os dados do Excel no DT
        Dim DT As New DataTable
        Try
            Dim connection As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=YES;IMEX=1';DATA SOURCE=" & Caminho & ";")
            Dim DA As New OleDb.OleDbDataAdapter
            connection.Open()
            'DT.Load(Command.ExecuteReader(CommandBehavior.CloseConnection))
            DA.SelectCommand = New OleDb.OleDbCommand("select * from [Carga$];", connection)
            DA.Fill(DT)
            connection.Close()
            connection.Dispose()
            DA.Dispose()
            GC.Collect()
        Catch
            MsgBox("Erro na consulta da Carga", MsgBoxStyle.Critical)
            Exit Sub
        End Try

        'Limpa a base e insere com os dados DT
        Excluir_Tudo()

        'Inserir Dados

        Using connectionS As New SQLite.SQLiteConnection(connstr)
            SQLite.SQLiteConnection.ClearAllPools()
            Dim cmd As New SQLite.SQLiteCommand
            If connectionS.State = ConnectionState.Closed Then
                connectionS.Open()
            End If
            cmd.Connection = connectionS
            'Variaveis
            Dim V_ID As Integer
            Dim V_Seq As String
            Dim V_CC As String
            Dim V_C_Inst As String
            Dim V_D_Inst As String
            Dim V_Local As String
            Dim V_Tag As String
            Dim V_D_Simp As String
            Dim V_D_Det As String
            Dim V_Fab As String
            Dim V_Mod As String
            Dim V_Serie As String
            Dim V_LocalF As String
            Dim V_Qtd As Decimal
            Dim V_Um As String
            Dim V_Ano As Integer
            Dim V_Mes As Integer
            Dim V_Dia As Integer
            Dim V_Status As String
            Dim V_Estado As String
            Dim V_Obs As String
            Dim V_Alt As Decimal
            Dim V_Larg As Decimal
            Dim V_Comp As Decimal
            Dim V_Area As Decimal
            Dim V_Pe As Decimal
            Dim V_Esforco As Decimal
            Dim V_Obs_C As String
            Dim V_Consultor As String
            Dim V_Responsavel As String

            For i = 0 To DT.Rows.Count - 1
                Try
                    V_ID = DT.Rows(i)(0)
                    V_Seq = DT.Rows(i)(1)
                    V_CC = DT.Rows(i)(2)
                    V_C_Inst = DT.Rows(i)(3)
                    V_D_Inst = DT.Rows(i)(4)
                    V_Local = DT.Rows(i)(5)
                    V_Tag = DT.Rows(i)(6)
                    V_D_Simp = DT.Rows(i)(7)
                    V_D_Det = IIf(IsDBNull(DT.Rows(i)(8)), "", DT.Rows(i)(8))
                    V_Fab = IIf(IsDBNull(DT.Rows(i)(9)), "", DT.Rows(i)(9))
                    V_Mod = IIf(IsDBNull(DT.Rows(i)(10)), "", DT.Rows(i)(10))
                    V_Serie = IIf(IsDBNull(DT.Rows(i)(11)), "", DT.Rows(i)(11))
                    V_LocalF = DT.Rows(i)(12)
                    V_Qtd = DT.Rows(i)(13)
                    V_Um = DT.Rows(i)(14)
                    V_Ano = IIf(IsDBNull(DT.Rows(i)(15)), 0, DT.Rows(i)(15))
                    V_Mes = IIf(IsDBNull(DT.Rows(i)(16)), 0, DT.Rows(i)(16))
                    V_Dia = IIf(IsDBNull(DT.Rows(i)(17)), 0, DT.Rows(i)(17))
                    V_Status = DT.Rows(i)(18)
                    V_Estado = DT.Rows(i)(19)
                    V_Obs = IIf(IsDBNull(DT.Rows(i)(20)), "", DT.Rows(i)(20))
                    V_Alt = IIf(IsDBNull(DT.Rows(i)(21)), 0, DT.Rows(i)(21))
                    V_Larg = IIf(IsDBNull(DT.Rows(i)(22)), 0, DT.Rows(i)(22))
                    V_Comp = IIf(IsDBNull(DT.Rows(i)(23)), 0, DT.Rows(i)(23))
                    V_Area = IIf(IsDBNull(DT.Rows(i)(24)), 0, DT.Rows(i)(24))
                    V_Pe = IIf(IsDBNull(DT.Rows(i)(25)), 0, DT.Rows(i)(25))
                    V_Esforco = IIf(IsDBNull(DT.Rows(i)(26)), 0, DT.Rows(i)(26))
                    V_Obs_C = IIf(IsDBNull(DT.Rows(i)(27)), "", DT.Rows(i)(27))
                    V_Consultor = DT.Rows(i)(28)
                    V_Responsavel = DT.Rows(i)(29)

                    cmd.CommandText = "insert into Inventario (ID,Sequencial,Centro_Custo,Cod_Instalacao,Desc_Instalacao," &
                    "Local,Tag_Antigo,Tag_Novo,Desc_Simples,Desc_Detalhada,Fabricante,Modelo,Serie,Desc_Unificada," &
                    "Local_Fisico,Quantidade,Um,Ano,Mes,Dia,Status,Estado,Validacao,Observacao,Altura,Largura," &
                    "Comprimento,Area,Pe,Esforco,Obs_Civil,Foto,Consultor,Responsavel,Data_Hora" &
                    ") values(" & V_ID & ", '" & V_Seq & "', '" & V_CC & "', '" & V_C_Inst & "', '" & V_D_Inst & "', '" & V_Local &
                "', '" & V_Tag & "','', '" & V_D_Simp & "', '" & V_D_Det & "', '" & V_Fab & "', '" & V_Mod & "', '" & V_Serie & "','', '" &
                V_LocalF & "', " & V_Qtd & ", '" & V_Um & "', " & V_Ano & ", " & V_Mes & ", " & V_Dia & ", '" &
                V_Status & "', '" & V_Estado & "', 'INSERIDO','" & V_Obs & "', " & V_Alt & ", " & V_Larg & ", " & V_Comp & ", " &
                V_Area & ", " & V_Pe & ", " & V_Esforco & ", '" & V_Obs_C & "','','" & V_Consultor & "', '" & V_Responsavel & "','" & Now & "');"

                    cmd.ExecuteNonQuery()
                Catch
                    MsgBox("Erro na Carga", vbCritical)
                    Exit Sub
                End Try
            Next i
            cmd.Dispose()
            connectionS.Close()
            connectionS.Dispose()
            GC.Collect()
        End Using

        MsgBox("Dados Inseridos Com Sucesso", vbInformation)
    End Sub

    Public Sub Consulta_Descricao_Civil(ID As Integer, TxtBay As TextBox, cod_tuc As Integer, Cmbtuc As ComboBox, cod_tipo_bem As String, cmba1 As ComboBox,
                                        cod_uar As Integer, Cmbuar As ComboBox, Cmba2 As ComboBox,
                                        Cmba3 As ComboBox, Cmba4 As ComboBox, Cmba5 As ComboBox,
                                        Cmba6 As ComboBox, cod_cm1 As String, Cmbcm1 As ComboBox,
                                        cod_cm2 As String, Cmbcm2 As ComboBox, cod_cm3 As String, Cmbcm3 As ComboBox, txtdesc As RichTextBox, txtfabric As TextBox,
                                        txtmodelo As TextBox, txtobs As TextBox, txtqtd As TextBox, cmbun As ComboBox, cmbano As ComboBox, cmbmes As ComboBox,
                                        cmbdia As ComboBox, cmbstatus As ComboBox, cmbestado As ComboBox, txtaltura As TextBox, txtlarg As TextBox, txtcomp As TextBox,
                                        txtarea As TextBox, txtpe As TextBox, txtobs_civil As TextBox, txtesforco As TextBox, txtserie As TextBox, txtTag As TextBox)
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select bay,cod_tuc,desc_tuc,cod_tipo_bem,desc_tipo_bem,cod_uar,desc_uar,cod_a2,desc_a2," &
                    "cod_a3,desc_a3,cod_a4,desc_a4,cod_a5,desc_a5,cod_a6,desc_a6,cod_cm1,desc_cm1,cod_cm2," &
                    "desc_cm2,cod_cm3,desc_cm3,descricao,fabricante,modelo,observacao,quantidade,unidade_medida," &
                    "ano,mes,dia,status_bem,estado_bem,altura,largura,comprimento,area,pe_direito,obs_civil,esforco " &
                    "from INVENTARIO where ID=" & ID & ";"
            leitor = cmd.ExecuteReader
            leitor.Read()
            TxtBay.Text = leitor("bay")
            cod_tuc = leitor("cod_tuc")
            Cmbtuc.Text = leitor("desc_tuc")
            cod_tipo_bem = leitor("cod_tipo_bem")
            cmba1.Text = leitor("desc_tipo_bem")
            cod_uar = leitor("cod_uar")
            Cmbuar.Text = leitor("desc_uar")
            'Frm_Inventário.A2 = leitor("cod_a2")
            'Cmba2.Text = leitor("desc_a2")
            'Frm_Inventário.A3 = leitor("cod_a3")
            'Cmba3.Text = leitor("desc_a3")
            'Frm_Inventário.A4 = leitor("cod_a4")
            'Cmba4.Text = leitor("desc_a4")
            'Frm_Inventário.A5 = leitor("cod_a5")
            'Cmba5.Text = leitor("desc_a5")
            'Frm_Inventário.A6 = leitor("cod_a6")
            Cmba6.Text = leitor("desc_a6")
            cod_cm1 = leitor("cod_cm1")
            Cmbcm1.Text = leitor("desc_cm1")
            cod_cm2 = leitor("cod_cm2")
            Cmbcm2.Text = leitor("desc_cm2")
            cod_cm3 = leitor("cod_cm3")
            Cmbcm3.Text = leitor("desc_cm3")
            txtdesc.Text = leitor("descricao")
            txtfabric.Text = leitor("fabricante")
            txtmodelo.Text = leitor("modelo")
            txtserie.Text = ""
            txtTag.Text = ""
            txtobs.Text = leitor("observacao")
            txtqtd.Text = leitor("quantidade")
            cmbun.Text = leitor("unidade_medida")
            cmbano.Text = ""
            cmbmes.Text = ""
            cmbdia.Text = ""
            cmbstatus.Text = leitor("status_bem")
            cmbestado.Text = leitor("estado_bem")
            txtaltura.Text = leitor("altura")
            txtlarg.Text = leitor("largura")
            txtcomp.Text = leitor("comprimento")
            txtarea.Text = leitor("area")
            txtpe.Text = leitor("pe_direito")
            txtobs_civil.Text = leitor("obs_civil")
            txtesforco.Text = leitor("esforco")

            cmd.Dispose()
            connection.Close()
            connection.Dispose()

        Catch
            MsgBox("Erro ao consultar descricao_civil ", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Consulta_TUC(cmb As ComboBox)
        cmb.Items.Clear()
        'Try
        Dim leitor As SQLite.SQLiteDataReader
        Dim connection As New SQLite.SQLiteConnection(connstr)
        Dim cmd As New SQLite.SQLiteCommand
        connection.Open()
        cmd.Connection = connection
        cmd.CommandText = "select descricao from tuc;"
        leitor = cmd.ExecuteReader

        Do While leitor.Read
            cmb.Items.Add(leitor("Descricao"))
        Loop
        cmd.Dispose()
        connection.Close()
        connection.Dispose()
        'Catch
        'MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
        'End Try
    End Sub

    Public Sub Consulta_CM(cmb1 As ComboBox, cmb2 As ComboBox, cmb3 As ComboBox)
        cmb1.Items.Clear()
        cmb2.Items.Clear()
        cmb3.Items.Clear()
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select CM1,CM2,CM3 from CM;"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                If Not leitor("CM1") = "" Then
                    cmb1.Items.Add(leitor("CM1"))
                End If
                cmb2.Items.Add(leitor("CM2"))
                If Not leitor("CM3") = "" Then
                    cmb3.Items.Add(leitor("CM3"))
                End If
            Loop
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            GC.Collect()
        Catch
            MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
        End Try
    End Sub


    Public Sub Inserir_Dados(ID As Integer, Sequencial As String, Centro_Custo As String, Cod_Inst As String,
                             Desc_Inst As Integer, Local As String, Tag_Antiga As String, Tag_Nova As String, Desc_Simples As Integer,
                             Desc_Detalhada As String, fabricante As String, modelo As String, serie As String,
                             Local_Fisico As String, observacao As String, quantidade As Decimal, unidade As String,
                             ano As Integer, mes As Integer, dia As Integer, status_bem As String,
                             estado_bem As String, altura As Decimal, largura As Decimal, comprimento As Decimal,
                             area As Decimal, pe As Decimal, esforco As Decimal, obs_civil As String, Validacao As String,
                             Consultor As String, Responsavel As String, Foto As String,
                             Desc_Unificada As String)
        'Try
        Using connection As New SQLite.SQLiteConnection(connstr)
                SQLite.SQLiteConnection.ClearAllPools()
                Dim cmd As New SQLite.SQLiteCommand
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                cmd.Connection = connection
                cmd.CommandText = "insert into Inventario (ID,Sequencial,Centro_Custo,Cod_Instalacao,Desc_Instalacao," &
                    "Local,Tag_Antigo,Tag_Novo,Desc_Simples,Desc_Detalhada,Fabricante,Modelo,Serie,Desc_Unificada," &
                    "Local_Fisico,Quantidade,Um,Ano,Mes,Dia,Status,Estado,Validacao,Observacao,Altura,Largura," &
                    "Comprimento,Area,Pe,Esforco,Obs_Civil,Foto,Consultor,Responsavel,Data_Hora " &
                    ") values(" & ID & ",'" & Sequencial & "','" & Centro_Custo & "','" & Cod_Inst & "','" & Desc_Inst &
                    "','" & Local & "','" & Tag_Antiga & "','" & Tag_Nova & "','" & Desc_Simples & "','" & Desc_Detalhada &
                    "','" & fabricante & "', '" & modelo & "','" & serie & "','" & Desc_Unificada & "','" &
                    Local_Fisico & "'," & Replace(CStr(quantidade), ",", ".") & ", '" & unidade & "'," & ano & ", " &
                    mes & "," & dia & ",'" & status_bem & "','" & estado_bem & "','" & Validacao & "','" & observacao &
                    "'," & Replace(CStr(altura), ",", ".") & "," & Replace(CStr(largura), ",", ".") & "," &
                    Replace(CStr(comprimento), ",", ".") & "," & Replace(CStr(area), ",", ".") & "," &
                    Replace(CStr(pe), ",", ".") & "," & Replace(CStr(esforco), ",", ".") & ",'" & obs_civil & "','" &
                    Foto & "','" & "','" & Consultor & "','" & Responsavel & "','" & Now & "');"
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connection.Close()
                connection.Dispose()
                GC.Collect()
            End Using
        'Catch
        '    MsgBox("Erro ao inserir dados", MsgBoxStyle.Critical)
        'End Try
    End Sub

    Public Sub Update_Inventario(ID As Integer, seq As String, local As String, odi As String, cod_ti As String, ti As String, bay As String,
                                 cod_tuc As String, tuc As String, cod_tipo_bem As String, tipo_bem As String, cod_uar As String, uar As String,
                                 cod_a2 As String, a2 As String, cod_a3 As String, a3 As String, cod_a4 As String, a4 As String, cod_a5 As String,
                                 a5 As String, cod_a6 As String, a6 As String, cod_cm1 As String, cm1 As String, cod_cm2 As String, cm2 As String,
                                 cod_cm3 As String, cm3 As String, descricao As String, fabricante As String, modelo As String, serie As String,
                                 obs As String, qtd As Decimal, um As String, ano As String, mes As String, dia As String, status As String,
                                 estado_bem As String, altura As Decimal, largura As Decimal, comprimento As Decimal, area As Decimal, pe As Decimal,
                                 obs_civil As String, foto As String, consultor As String, lider As String, TAG As String, esforco As Decimal)
        Try
            Using connection As New SQLite.SQLiteConnection(connstr)
                SQLite.SQLiteConnection.ClearAllPools()
                Dim cmd As New SQLite.SQLiteCommand
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                cmd.Connection = connection
                cmd.CommandText = "update Inventario set Sequencial='" & seq & "',Local='" & local & "',ODI='" & odi & "',cod_ti=" &
                    cod_ti & ",ti='" & ti & "',Bay='" & bay & "',cod_tuc='" & cod_tuc & "',desc_tuc='" & tuc & "',cod_tipo_bem='" &
                    cod_tipo_bem & "',desc_tipo_bem='" & tipo_bem & "',cod_uar='" & cod_uar & "',desc_uar='" & uar &
                    "',Cod_A2='" & cod_a2 & "',desc_A2='" & a2 & "',Cod_A3='" & cod_a3 & "',desc_A3='" & a3 & "',cod_A4='" &
                    cod_a4 & "',Desc_A4='" & a4 & "',Cod_A5='" & cod_a5 & "',Desc_A5='" & a5 & "',Cod_A6='" & cod_a6 &
                    "',Desc_A6='" & a6 & "',Cod_CM1='" & cod_cm1 & "',Desc_CM1='" & cm1 & "',Cod_CM2='" & cod_cm2 &
                    "',Desc_CM2='" & cm2 & "',Cod_CM3='" & cod_cm3 & "',Desc_CM3='" & cm3 & "',Descricao='" & descricao &
                    "',Fabricante='" & fabricante & "',Modelo='" & modelo & "',serie='" & serie & "',Observacao='" & obs & "',Quantidade=" &
                    Replace(CStr(qtd), ",", ".") & ",Unidade_Medida='" & um & "',Ano='" & ano & "',Mes='" & mes & "',Dia='" & dia &
                    "',Status_Bem='" & status & "',Estado_Bem='" & estado_bem & "',Altura=" & Replace(CStr(altura), ",", ".") & ",Largura=" & Replace(CStr(largura), ",", ".") & ",Comprimento=" &
                    Replace(CStr(comprimento), ",", ".") & ",area=" & Replace(CStr(area), ",", ".") & ",Pe_direito=" & Replace(CStr(pe), ",", ".") & ",Obs_Civil='" & obs_civil & "',foto='" & foto & "',Consultor='" & consultor & "',Lider='" &
                    lider & "',Data_Hora='" & Now & "',Numero_Manutencao='" & TAG & "',esforco=" & Replace(CStr(esforco), ",", ".") & " where ID=" & ID & ";"

                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connection.Close()
                connection.Dispose()
            End Using
            'MsgBox("Dados Salvos com Sucesso", MsgBoxStyle.Information)
        Catch
            MsgBox("Erro ao atualizar dados", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Excluir_Carga_Cmb()
        Try
            Using connection As New SQLite.SQLiteConnection(connstr)
                SQLite.SQLiteConnection.ClearAllPools()
                Dim cmd As New SQLite.SQLiteCommand
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                cmd.Connection = connection
                cmd.CommandText = "Delete from Carga_Cmb;"
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connection.Close()
                connection.Dispose()
            End Using
            MsgBox("Dados Excluídos com Sucesso", MsgBoxStyle.Information)
        Catch
            MsgBox("Erro ao excluir dados", MsgBoxStyle.Critical)
            Frm_Inventário.Erro_Excluir = True
        End Try
    End Sub

    Public Sub Excluir_Tudo()
        Try
            Using connection As New SQLite.SQLiteConnection(connstr)
                SQLite.SQLiteConnection.ClearAllPools()
                Dim cmd As New SQLite.SQLiteCommand
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                cmd.Connection = connection
                cmd.CommandText = "Delete from Inventario;"
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connection.Close()
                connection.Dispose()
            End Using
            'MsgBox("Dados Excluídos com Sucesso", MsgBoxStyle.Information)
        Catch
            MsgBox("Erro ao excluir dados", MsgBoxStyle.Critical)
            Frm_Inventário.Erro_Excluir = True
        End Try
    End Sub

    Public Sub Excluir(ID As Integer)
        Try
            Using connection As New SQLite.SQLiteConnection(connstr)
                SQLite.SQLiteConnection.ClearAllPools()
                Dim cmd As New SQLite.SQLiteCommand
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                cmd.Connection = connection
                cmd.CommandText = "Delete from Inventario where ID=" & ID & ";"
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connection.Close()
                connection.Dispose()
            End Using
            MsgBox("Dados Excluídos com Sucesso", MsgBoxStyle.Information)
        Catch
            MsgBox("Erro ao excluir dados", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Consulta_Excel()
        Try
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim DA As New SQLite.SQLiteDataAdapter
            connection.Open()
            DA.SelectCommand = New SQLite.SQLiteCommand("select * from Inventario;", connection)
            DA.Fill(DS)
            connection.Close()
            connection.Dispose()
            DA.Dispose()
            GC.Collect()
        Catch
            MsgBox("Erro na consulta", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Consulta_Grid(DGV As DataGridView)
        Try
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim DA As New SQLite.SQLiteDataAdapter
            Dim DT As New DataTable
            connection.Open()
            DA.SelectCommand = New SQLite.SQLiteCommand("select * from Inventario;", connection)
            DA.Fill(DT)
            DGV.DataSource = DT
            connection.Close()
            connection.Dispose()
            DA.Dispose()
            GC.Collect()
        Catch
            MsgBox("Erro na consulta", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Function Buscar_Ultimo_ID()
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select ID from Inventario order by ID DESC;"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Dim a As String
            a = leitor("ID")
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            Return a
        Catch
            Return 0
        End Try
    End Function

    Public Sub Modelo_Excel()
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim Sh_T As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Consulta_Excel()

        Try
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            Sh_T = xlWorkBook.Sheets(1)
            Sh_T.Name = "Inventario"
            Sh_T.Range("a1").Value = "ID"
            Sh_T.Range("b1").Value = "Sequencial"
            Sh_T.Range("c1").Value = "Centro de Custo"
            Sh_T.Range("d1").Value = "Código de Instalação"
            Sh_T.Range("e1").Value = "Descrição de Instalação"
            Sh_T.Range("f1").Value = "Local"
            Sh_T.Range("g1").Value = "N° de Manutenção Antigo"
            Sh_T.Range("h1").Value = "N° de Manutenção Novo"
            Sh_T.Range("i1").Value = "Descrição Simples"
            Sh_T.Range("j1").Value = "Descrição Detalhada"
            Sh_T.Range("k1").Value = "Fabricante"
            Sh_T.Range("l1").Value = "Modelo"
            Sh_T.Range("m1").Value = "N° de Série"
            Sh_T.Range("n1").Value = "Descrição Unificada"
            Sh_T.Range("o1").Value = "Local Físico"
            Sh_T.Range("p1").Value = "Quantidade"
            Sh_T.Range("q1").Value = "Unidade de Medida"
            Sh_T.Range("r1").Value = "Ano de Fabricação"
            Sh_T.Range("s1").Value = "Mês de Fabricação"
            Sh_T.Range("t1").Value = "Dia de Fabricação"
            Sh_T.Range("u1").Value = "Status do Bem"
            Sh_T.Range("v1").Value = "Estado do Bem"
            Sh_T.Range("w1").Value = "Validação"
            Sh_T.Range("x1").Value = "Observação"
            Sh_T.Range("y1").Value = "Altura"
            Sh_T.Range("aa1").Value = "Largura"
            Sh_T.Range("ab1").Value = "Comprimento"
            Sh_T.Range("ac1").Value = "Área"
            Sh_T.Range("ad1").Value = "Pé Direito"
            Sh_T.Range("ae1").Value = "Esforço"
            Sh_T.Range("af1").Value = "Observação Civil"
            Sh_T.Range("ag1").Value = "Foto"
            Sh_T.Range("ah1").Value = "Consultor"
            Sh_T.Range("ai1").Value = "Responsável"
            Sh_T.Range("aj1").Value = "Data/Hora"

            Sh_T.Range("a1:aj1").Font.Bold = True
            Sh_T.Range("a1:aj1").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue)

            'Arrumar colunas
            'DS.Tables(0).Columns(50).SetOrdinal(33)
            'DS.Tables(0).Columns(51).SetOrdinal(47)
            Dim Linhas As Integer
            Linhas = DS.Tables(0).Rows.Count
            Frm_Inventário.PB_Excel.Value = 0
            Frm_Inventário.PB_Excel.Visible = True
            'Inserir linhas
            For i = 0 To Linhas - 1
                For j = 0 To DS.Tables(0).Columns.Count - 1
                    'Colunas anteriores
                    If j <= 30 Or j >= 32 Then
                        Sh_T.Cells(i + 2, j + 1) = DS.Tables(0).Rows(i).Item(j)
                    End If
                    'fotos
                    If j = 31 Then
                        Sh_T.Cells(i + 2, j + 1) = DS.Tables(0).Rows(i).Item(31).ToString.Replace(".bmp", "").Replace(".jpg", "").Replace(".png", "").Replace("|", ", ")
                    End If
                Next
                Frm_Inventário.PB_Excel.Value = ((i + 1) / Linhas) * 100
            Next

            Sh_T.Columns.AutoFit()
            xlApp.Visible = True
            Frm_Inventário.PB_Excel.Visible = False
        Catch
            MsgBox("Erro ao Carregar Excel!", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Layout_Excel()
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim Sh_T As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        Try
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            Sh_T = xlWorkBook.Sheets(1)
            Sh_T.Name = "Carga"
            Sh_T.Range("a1").Value = "ID"
            Sh_T.Range("b1").Value = "Sequencial"
            Sh_T.Range("c1").Value = "Centro de Custo"
            Sh_T.Range("d1").Value = "Código de Instalação"
            Sh_T.Range("e1").Value = "Descrição de Instalação"
            Sh_T.Range("f1").Value = "Local"
            Sh_T.Range("g1").Value = "Tag"
            Sh_T.Range("h1").Value = "Descrição Simples"
            Sh_T.Range("i1").Value = "Descrição Detalhada"
            Sh_T.Range("j1").Value = "Fabricante"
            Sh_T.Range("k1").Value = "Modelo"
            Sh_T.Range("l1").Value = "Série"
            Sh_T.Range("m1").Value = "Local Físico"
            Sh_T.Range("n1").Value = "Quantidade"
            Sh_T.Range("o1").Value = "Unidade de Medida"
            Sh_T.Range("p1").Value = "Ano de Fabricação"
            Sh_T.Range("q1").Value = "Mês de Fabricação"
            Sh_T.Range("r1").Value = "Dia de Fabricação"
            Sh_T.Range("s1").Value = "Status do Bem"
            Sh_T.Range("t1").Value = "Estado do Bem"
            Sh_T.Range("u1").Value = "Observação"
            Sh_T.Range("v1").Value = "Altura"
            Sh_T.Range("w1").Value = "Largura"
            Sh_T.Range("x1").Value = "Comprimento"
            Sh_T.Range("y1").Value = "Área"
            Sh_T.Range("z1").Value = "Pé Direito"
            Sh_T.Range("aa1").Value = "Esforço"
            Sh_T.Range("ab1").Value = "Observação Civil"
            Sh_T.Range("ac1").Value = "Consultor"
            Sh_T.Range("ad1").Value = "Responsável"

            Sh_T.Range("a1:ad1").Font.Bold = True
            Sh_T.Range("a1:ad1").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue)


            Sh_T.Columns.AutoFit()
            xlApp.Visible = True
        Catch
            MsgBox("Erro ao Carregar Excel!", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Layout_Excel_Cmb()
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim Sh_T As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        Try
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            Sh_T = xlWorkBook.Sheets(1)
            Sh_T.Name = "Carga"
            Sh_T.Range("a1").Value = "Centro de Custo"
            Sh_T.Range("b1").Value = "Código de Instalação"
            Sh_T.Range("c1").Value = "Descrição de Instalação"
            Sh_T.Range("d1").Value = "Local"
            Sh_T.Range("e1").Value = "Tag"
            Sh_T.Range("f1").Value = "Descrição"
            Sh_T.Range("g1").Value = "Fabricante"
            Sh_T.Range("h1").Value = "Modelo"
            Sh_T.Range("i1").Value = "Local Físico"
            Sh_T.Range("j1").Value = "Consultor"
            Sh_T.Range("k1").Value = "Responsável"

            Sh_T.Range("a1:k1").Font.Bold = True
            Sh_T.Range("a1:k1").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue)

            'Arrumar colunas
            'DS.Tables(0).Columns(50).SetOrdinal(33)
            'DS.Tables(0).Columns(51).SetOrdinal(47)

            Sh_T.Columns.AutoFit()
            xlApp.Visible = True
        Catch
            MsgBox("Erro ao Carregar Excel!", MsgBoxStyle.Critical)
        End Try
    End Sub
End Class
