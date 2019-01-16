Imports System.Net.Mail
Imports Microsoft.Office.Interop

Public Class Inventário_Excel
    Dim connstr As String = "Data Source=C:\Users\Public\INVENTARIO_PADRAO.db;;Version=3;New=True;Compress=True;Pooling=True"
    Public DS As New DataSet
    Public DTExpira As New DataTable

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
            Dim connectionD As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connectionD.Open()
            cmd.Connection = connectionD
            cmd.CommandText = "delete from Carga_Cmb;"
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            connectionD.Close()
            connectionD.Dispose()
        Catch
            MsgBox("Erro ao limpar Base", MsgBoxStyle.Critical)
        End Try

        Try
            Dim connectionS As New SQLite.SQLiteConnection(connstr)
            Dim cmdS As New SQLite.SQLiteCommand
            connectionS.Open()
            cmdS.Connection = connectionS
            For i = 0 To DT.Rows.Count - 1
                cmdS.CommandText = "insert into Carga_Cmb (Centro_Custo,Cod_Instalacao,Desc_Instalacao,Local," &
                        "Tag,Descricao,Fabricante,Modelo,Local_Fisico,Consultor,Responsável) values ('" &
                        DT.Rows(i)(0) & "','" & DT.Rows(i)(1) & "','" & DT.Rows(i)(2) & "','" & DT.Rows(i)(3) & "','" &
                        UCase(DT.Rows(i)(4)) & "','" & DT.Rows(i)(5) & "','" & DT.Rows(i)(6) & "','" & DT.Rows(i)(7) & "','" &
                        DT.Rows(i)(8) & "','" & DT.Rows(i)(9) & "','" & DT.Rows(i)(10) & "');"
                cmdS.ExecuteNonQuery()
            Next i
            cmdS.Dispose()
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
                    V_Tag = UCase(DT.Rows(i)(6))
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

    Public Sub Consulta_Descricao_Civil(ID As Integer, CmbTag_A As ComboBox, TxtTag_N As TextBox, CmbDesc As ComboBox, TxtDetal As RichTextBox,
                                        Cmbfabric As ComboBox, Cmbmodelo As ComboBox, txtobs As RichTextBox, txtqtd As TextBox, cmbun As ComboBox,
                                        cmbstatus As ComboBox, cmbestado As ComboBox, txtaltura As TextBox, txtlarg As TextBox, txtcomp As TextBox,
                                        txtarea As TextBox, txtpe As TextBox, txtesforco As TextBox, txtobs_civil As TextBox, CmbLocal_F As ComboBox)
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select ID,Tag_Antigo,Tag_Novo,Desc_Simples,Desc_Detalhada,Fabricante,Modelo," &
                        "Local_Fisico,Quantidade,Um,Status,Estado,Observacao,Altura,Largura," &
                        "Comprimento,Area,Pe,Esforco,Obs_Civil " &
                    "from INVENTARIO where ID=" & ID & ";"
            leitor = cmd.ExecuteReader
            leitor.Read()
            CmbLocal_F.Text = leitor("Local_Fisico")
            CmbTag_A.Text = leitor("Tag_Antigo")
            TxtTag_N.Text = leitor("Tag_Novo")
            CmbDesc.Text = leitor("Desc_Simples")
            TxtDetal.Text = leitor("Desc_Detalhada")
            Cmbfabric.Text = leitor("Fabricante")
            Cmbmodelo.Text = leitor("Modelo")
            txtobs.Text = leitor("Observacao")
            txtqtd.Text = leitor("Quantidade")
            cmbun.Text = leitor("Um")
            cmbstatus.Text = leitor("Status")
            cmbestado.Text = leitor("Estado")
            txtaltura.Text = leitor("altura")
            txtlarg.Text = leitor("largura")
            txtcomp.Text = leitor("comprimento")
            txtarea.Text = leitor("area")
            txtpe.Text = leitor("Pe")
            txtobs_civil.Text = leitor("Obs_Civil")
            txtesforco.Text = leitor("Esforco")

            cmd.Dispose()
            connection.Close()
            connection.Dispose()

        Catch
            MsgBox("Erro ao copiar Dados ", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Consulta_Tag(ID As Integer, CmbCC As ComboBox, CmbLocal As ComboBox, Cod_Inst As String, CmbInst As ComboBox, CmbTag_A As ComboBox, TxtTag_N As TextBox,
                                        CmbDesc As ComboBox, TxtDetal As RichTextBox, CmbConsultor As ComboBox, CmbResp As ComboBox,
                                        Cmbfabric As ComboBox, Cmbmodelo As ComboBox, TxtSerie As TextBox, txtobs As RichTextBox, txtqtd As TextBox, cmbun As ComboBox,
                                        cmbstatus As ComboBox, cmbano As ComboBox, cmbmes As ComboBox, cmbdia As ComboBox,
                                        cmbestado As ComboBox, txtaltura As TextBox, txtlarg As TextBox, txtcomp As TextBox,
                                        txtarea As TextBox, txtpe As TextBox, txtesforco As TextBox, txtobs_civil As TextBox, CmbLocal_F As ComboBox, Tag As String,
                                        TxtIDLocal As TextBox, TxtIDDesc As TextBox, TxtIDCivil As TextBox)
        Try
            Dim Texto_Foto As String

            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select ID,Tag_Antigo,Tag_Novo,Desc_Simples,Desc_Detalhada,Fabricante,Modelo,Serie," &
                        "Local_Fisico,Quantidade,Um,Status,Estado,Ano,Mes,Dia,Observacao,Altura,Largura," &
                        "Comprimento,Area,Pe,Esforco,Obs_Civil,Centro_Custo,Cod_Instalacao,Desc_Instalacao,Local,Consultor,Responsavel,Foto " &
                    "from INVENTARIO where Tag_Antigo='" & Tag & "' or Tag_Novo='" & Tag & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            CmbCC.Text = leitor("Centro_Custo")
            Cod_Inst = leitor("Cod_Instalacao")
            CmbInst.Text = leitor("Desc_Instalacao")
            CmbLocal.Text = leitor("Local")
            CmbConsultor.Text = leitor("Consultor")
            CmbResp.Text = leitor("Responsavel")
            CmbLocal_F.Text = leitor("Local_Fisico")
            CmbTag_A.Text = leitor("Tag_Antigo")
            TxtTag_N.Text = leitor("Tag_Novo")
            CmbDesc.Text = leitor("Desc_Simples")
            TxtDetal.Text = leitor("Desc_Detalhada")
            Cmbfabric.Text = leitor("Fabricante")
            Cmbmodelo.Text = leitor("Modelo")
            TxtSerie.Text = leitor("Serie")
            txtobs.Text = leitor("Observacao")
            txtqtd.Text = leitor("Quantidade")
            cmbun.Text = leitor("Um")
            cmbstatus.Text = leitor("Status")
            cmbestado.Text = leitor("Estado")
            cmbano.Text = leitor("Ano")
            cmbmes.Text = leitor("Mes")
            cmbdia.Text = leitor("Dia")
            txtaltura.Text = leitor("altura")
            txtlarg.Text = leitor("largura")
            txtcomp.Text = leitor("comprimento")
            txtarea.Text = leitor("area")
            txtpe.Text = leitor("Pe")
            txtobs_civil.Text = leitor("Obs_Civil")
            txtesforco.Text = leitor("Esforco")
            TxtIDLocal.Text = leitor("ID")
            TxtIDDesc.Text = leitor("ID")
            TxtIDCivil.Text = leitor("ID")
            Texto_Foto = leitor("Foto")

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

            cmd.Dispose()
            connection.Close()
            connection.Dispose()

        Catch
            MsgBox("Tag não localizado.", MsgBoxStyle.Information)
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

    Public Function Consulta_Validar(ID As Integer, CC As String, Cod_Inst As String, Desc_Inst As String, Local As String, Tag_A As String,
                                     Tag_N As String, Desc_S As String, Desc_D As String, Fabricante As String,
                                     Modelo As String, Serie As String, Local_F As String, Qtd As Decimal, Um As String,
                                     Ano As Integer, Mes As Integer, Dia As Integer, Status As String, Estado As String,
                                     Obs As String, Altura As Decimal, Largura As Decimal, Comp As Decimal, Area As Decimal,
                                     Pe As Decimal, Esforco As Decimal, Obs_C As String, Consultor As String, Resp As String) As Boolean
        Try
            Consulta_Validar = False

            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select Centro_Custo,Cod_Instalacao,Desc_Instalacao," &
                        "Local,Tag_Antigo,Tag_Novo,Desc_Simples,Desc_Detalhada,Fabricante,Modelo,Serie," &
                        "Local_Fisico,Quantidade,Um,Ano,Mes,Dia,Status,Estado,Observacao,Altura,Largura," &
                        "Comprimento,Area,Pe,Esforco,Obs_Civil,Consultor,Responsavel from Inventario Where ID=" & ID & ";"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                If Not leitor("Centro_Custo") = CC Then
                    Consulta_Validar = True
                End If
                If Not leitor("Cod_Instalacao") = Cod_Inst Then
                    Consulta_Validar = True
                End If
                If Not leitor("Desc_Instalacao") = Desc_Inst Then
                    Consulta_Validar = True
                End If
                If Not leitor("Local") = Local Then
                    Consulta_Validar = True
                End If
                If Not leitor("Tag_Antigo") = Tag_A Then
                    Consulta_Validar = True
                End If
                If Not leitor("Tag_Novo") = Tag_N Then
                    Consulta_Validar = True
                End If
                If Not leitor("Desc_Simples") = Desc_S Then
                    Consulta_Validar = True
                End If
                If Not leitor("Desc_Detalhada") = Desc_D Then
                    Consulta_Validar = True
                End If
                If Not leitor("Fabricante") = Fabricante Then
                    Consulta_Validar = True
                End If
                If Not leitor("Modelo") = Modelo Then
                    Consulta_Validar = True
                End If
                If Not leitor("Serie") = Serie Then
                    Consulta_Validar = True
                End If
                If Not leitor("Local_Fisico") = Local_F Then
                    Consulta_Validar = True
                End If
                If Not leitor("Quantidade") = Qtd Then
                    Consulta_Validar = True
                End If
                If Not leitor("Um") = Um Then
                    Consulta_Validar = True
                End If
                If Not leitor("Ano") = Ano Then
                    Consulta_Validar = True
                End If
                If Not leitor("Mes") = Mes Then
                    Consulta_Validar = True
                End If
                If Not leitor("Dia") = Dia Then
                    Consulta_Validar = True
                End If
                If Not leitor("Status") = Status Then
                    Consulta_Validar = True
                End If
                If Not leitor("Estado") = Estado Then
                    Consulta_Validar = True
                End If
                If Not leitor("Observacao") = Obs Then
                    Consulta_Validar = True
                End If
                If Not leitor("Altura") = Altura Then
                    Consulta_Validar = True
                End If
                If Not leitor("Largura") = Largura Then
                    Consulta_Validar = True
                End If
                If Not leitor("Comprimento") = Comp Then
                    Consulta_Validar = True
                End If
                If Not leitor("Area") = Area Then
                    Consulta_Validar = True
                End If
                If Not leitor("Pe") = Pe Then
                    Consulta_Validar = True
                End If
                If Not leitor("Esforco") = Esforco Then
                    Consulta_Validar = True
                End If
                If Not leitor("Obs_Civil") = Obs_C Then
                    Consulta_Validar = True
                End If
                If Not leitor("Consultor") = Consultor Then
                    Consulta_Validar = True
                End If
                If Not leitor("Responsavel") = Resp Then
                    Consulta_Validar = True
                End If
            Loop
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            GC.Collect()
        Catch
            Consulta_Validar = False
            'MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
        End Try
    End Function


    Public Sub Inserir_Dados(ID As Integer, Sequencial As String, local As String, CC As String, cod_Inst As String, Desc_Inst As String, Tag_A As String,
                                 Tag_N As String, Desc_S As String, Desc_D As String, fabricante As String, modelo As String, serie As String,
                                 Local_F As String, obs As String, Validacao As String,
                                 qtd As Decimal, um As String, ano As String, mes As String, dia As String, status As String,
                                 estado_bem As String, altura As Decimal, largura As Decimal, comprimento As Decimal, area As Decimal, pe As Decimal,
                                 esforco As Decimal, obs_civil As String, foto As String, consultor As String, Resp As String)
        Try
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
                        ") values(" & ID & ",'" & Sequencial & "','" & CC & "','" & cod_Inst & "','" & Desc_Inst &
                        "','" & local & "','" & UCase(Tag_A) & "','" & UCase(Tag_N) & "','" & Desc_S & "','" & Desc_D &
                        "','" & fabricante & "', '" & modelo & "','" & serie & "','" & Desc_S & " | " & Desc_D & " | " & fabricante & " | " & modelo & " | " & serie & "','" &
                        Local_F & "'," & Replace(CStr(qtd), ",", ".") & ", '" & um & "'," & ano & ", " &
                        mes & "," & dia & ",'" & status & "','" & estado_bem & "','" & Validacao & "','" & obs &
                        "'," & Replace(CStr(altura), ",", ".") & "," & Replace(CStr(largura), ",", ".") & "," &
                        Replace(CStr(comprimento), ",", ".") & "," & Replace(CStr(area), ",", ".") & "," &
                        Replace(CStr(pe), ",", ".") & "," & Replace(CStr(esforco), ",", ".") & ",'" & obs_civil & "','" &
                        foto & "','" & consultor & "','" & Resp & "','" & Now & "');"
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connection.Close()
                connection.Dispose()
                GC.Collect()
            End Using
        Catch
            MsgBox("Erro ao inserir dados", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Update_Inventario(ID As Integer, local As String, CC As String, cod_Inst As String, Desc_Inst As String, Tag_A As String,
                                 Tag_N As String, Desc_S As String, Desc_D As String, fabricante As String, modelo As String, serie As String,
                                 Local_F As String, obs As String, Validacao As String,
                                 qtd As Decimal, um As String, ano As String, mes As String, dia As String, status As String,
                                 estado_bem As String, altura As Decimal, largura As Decimal, comprimento As Decimal, area As Decimal, pe As Decimal,
                                 esforco As Decimal, obs_civil As String, foto As String, consultor As String, Resp As String)
        Try
            Using connection As New SQLite.SQLiteConnection(connstr)
                SQLite.SQLiteConnection.ClearAllPools()
                Dim cmd As New SQLite.SQLiteCommand
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                cmd.Connection = connection
                cmd.CommandText = "update Inventario set Local='" & local & "',Local_Fisico='" & Local_F & "',Centro_Custo='" & CC & "',Cod_Instalacao='" &
                        cod_Inst & "',Desc_Instalacao='" & Desc_Inst & "',Tag_Antigo='" & UCase(Tag_A) & "',Tag_Novo='" & UCase(Tag_N) & "',Desc_Simples='" & Desc_S & "',Desc_Detalhada='" &
                        Desc_D & "',Fabricante='" & fabricante & "',Modelo='" & modelo & "',serie='" & serie & "',Observacao='" & obs & "',Quantidade=" &
                        Replace(CStr(qtd), ",", ".") & ",Um='" & um & "',Ano='" & ano & "',Mes='" & mes & "',Dia='" & dia &
                        "',Status='" & status & "',Estado='" & estado_bem & "',Altura=" & Replace(CStr(altura), ",", ".") & ",Largura=" & Replace(CStr(largura), ",", ".") & ",Comprimento=" &
                        Replace(CStr(comprimento), ",", ".") & ",area=" & Replace(CStr(area), ",", ".") & ",Pe=" & Replace(CStr(pe), ",", ".") & ",Obs_Civil='" & obs_civil & "',Foto='" & foto & "',Consultor='" & consultor & "',Responsavel='" &
                        Resp & "',Data_Hora='" & Now & "',Validacao='" & Validacao & "',esforco=" & Replace(CStr(esforco), ",", ".") & ", Desc_Unificada='" &
                        Desc_S & " | " & Desc_D & " | " & fabricante & " | " & modelo & " | " & serie & "' where ID=" & ID & ";"

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
            Sh_T.Range("z1").Value = "Largura"
            Sh_T.Range("aa1").Value = "Comprimento"
            Sh_T.Range("ab1").Value = "Área"
            Sh_T.Range("ac1").Value = "Pé Direito"
            Sh_T.Range("ad1").Value = "Esforço"
            Sh_T.Range("ae1").Value = "Observação Civil"
            Sh_T.Range("af1").Value = "Foto"
            Sh_T.Range("ag1").Value = "Consultor"
            Sh_T.Range("ah1").Value = "Responsável"
            Sh_T.Range("ai1").Value = "Data/Hora"

            Sh_T.Range("a1:ai1").Font.Bold = True
            Sh_T.Range("a1:ai1").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue)

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
                        If j = 34 Then
                            Sh_T.Cells(i + 2, j + 1) = Mid(DS.Tables(0).Rows(i).Item(j).ToString, 4, 2) & "/" &
                                Mid(DS.Tables(0).Rows(i).Item(j).ToString, 1, 2) & "/" &
                                Mid(DS.Tables(0).Rows(i).Item(j).ToString, 7, 4) &
                                Mid(DS.Tables(0).Rows(i).Item(j).ToString, 11, 20)
                        Else
                            Sh_T.Cells(i + 2, j + 1) = DS.Tables(0).Rows(i).Item(j)
                        End If
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

    Public Sub Buscar_Data_Limite()
        Try
            Dim connection As New SQLite.SQLiteConnection(connstr)
            connection.Open()
            Dim DA As New SQLite.SQLiteDataAdapter("select Data from DataLimite;", connection)
            DTExpira.Clear()
            DA.Fill(DTExpira)
            DA.Dispose()
            connection.Close()
            connection.Dispose()

        Catch
            'MsgBox("Erro ao buscar data limite", MsgBoxStyle.Critical)
        End Try
    End Sub



    Public Sub Update_Data_Limite(DTP As DateTimePicker)
        Try
            Using connection As New SQLite.SQLiteConnection(connstr)
                SQLite.SQLiteConnection.ClearAllPools()
                Dim cmd As New SQLite.SQLiteCommand
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                cmd.Connection = connection
                cmd.CommandText = "update DataLimite set Data='" & DTP.Value.ToShortDateString & "' where ID=1;"

                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connection.Close()
                connection.Dispose()
            End Using
            MsgBox("Dados Salvos com Sucesso", MsgBoxStyle.Information)
        Catch
            MsgBox("Erro ao atualizar dados", MsgBoxStyle.Critical)
        End Try
    End Sub
End Class
